"""结构化申报采集编排。

定位关系（定义名称 → 关系文件 → 业务表头），按企业编号而不是单元格行号关联三张表，
沿公式依赖图做受信重算，并把循环依赖、跨企业引用、不支持外链等问题隔离到具体申报，
不连带丢弃无关企业。

输出中金额一律为 ``Decimal``（JSON 序列化层再转字符串），企业编号保留文本形态。
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from dates import DateError, convert_serial
from formula_engine import (
    ExternalReferenceError,
    FormulaEngine,
    FormulaError,
)
from workbook_parser import Cell, Workbook, parse_ref_text, parse_workbook

ORG_ROLE, COST_ROLE, APPLY_ROLE = '企业表', '费用表', '申报表'
ROLE_NAMES = (ORG_ROLE, COST_ROLE, APPLY_ROLE)
ROLE_MARKERS = {
    ORG_ROLE: ('企业名称',),
    COST_ROLE: ('费用项', '金额'),
    APPLY_ROLE: ('合计', '日期'),
}
ID_HEADER = '企业编号'

# 问题代码
P_CYCLE = '循环依赖'
P_CROSS = '跨企业引用'
P_EXTERNAL = '外链公式'
P_UNSUPPORTED = '不支持的公式'
P_ERROR = '错误值'
P_FAKE_LEAP = '虚构闰日'
P_BAD_DATE = '非法日期'
P_UNREGISTERED = '企业编号未登记'
P_MACRO = '宏工作簿拒收'
P_FEE_MISMATCH = '费用复核不一致'
P_EVAL_FAILED = '公式重算失败'
W_STALE_CACHE = '缓存与重算不一致'
W_NO_CACHE = '缺少缓存值'
W_EXTERNAL_REL = '工作簿含外部关系'


@dataclass
class RoleTarget:
    role: str
    sheet_name: str
    located_by: str
    start_row: int
    end_row: int
    start_column: int
    end_column: int
    headers: dict[str, int]


def _column_letters(index: int) -> str:
    index += 1
    letters = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def cell_reference(row: int, column: int) -> str:
    return _column_letters(column) + str(row + 1)


def classify(cell: Cell | None) -> str:
    """返回四态等分类：空单元格 / 空字符串 / 错误值 / 零 / 数值 / 文本 / 布尔。"""
    if cell is None:
        return '空单元格'
    if cell.kind == 'error':
        return '错误值'
    if cell.kind == 'string':
        return '空字符串' if cell.is_empty_string else '文本'
    if cell.kind == 'boolean':
        return '布尔'
    if cell.number is None:
        return '空单元格'
    return '零' if cell.number == 0 else '数值'


def _header_map(workbook: Workbook, sheet_name: str,
                start_row: int | None = None) -> tuple[dict[str, int], int]:
    sheet = workbook.sheet_by_name(sheet_name)
    scan_from = 0 if start_row is None else start_row
    for row_index in sorted(sheet.rows):
        if row_index < scan_from:
            continue
        headers = {}
        for reference in sheet.rows.get(row_index, ()):
            cell = sheet.cells[reference]
            if cell.kind == 'string' and cell.text:
                headers[cell.text] = _ref_column(cell.reference)
        if ID_HEADER in headers:
            return headers, row_index
    raise ValueError(f'工作表 {sheet_name} 找不到业务表头（{ID_HEADER}）')


def _ref_column(reference: str) -> int:
    from workbook_parser import split_reference
    return split_reference(reference)[0]


def _locate_roles(workbook: Workbook) -> dict[str, RoleTarget]:
    targets: dict[str, RoleTarget] = {}
    for role in ROLE_NAMES:
        defined = workbook.defined_names.get(role)
        if defined is not None:
            sheet_name, start_col, start_row, end_col, end_row = parse_ref_text(defined)
            if workbook.sheet_by_name(sheet_name) is None:
                raise ValueError(f'定义名称 {role} 指向不存在的工作表')
            headers, header_row = _header_map(workbook, sheet_name, start_row)
            targets[role] = RoleTarget(role, sheet_name, f'定义名称:{role}',
                                       header_row + 1, end_row, start_col, end_col, headers)
            continue
        # 关系文件枚举所有内部工作表，用业务表头兜底定位：按角色标记词的重叠数评分。
        candidates = []
        taken = {targets[r].sheet_name for r in targets}
        for sheet in workbook.sheets:
            if sheet.external or sheet.name in taken:
                continue
            try:
                headers, header_row = _header_map(workbook, sheet.name)
            except ValueError:
                continue
            if ID_HEADER not in headers:
                continue
            score = sum(1 for marker in ROLE_MARKERS[role] if marker in headers)
            candidates.append((score, sheet, headers, header_row))
        if not candidates or max(item[0] for item in candidates) == 0:
            raise ValueError(f'无法通过定义名称或业务表头定位 {role}')
        _, sheet, headers, header_row = max(candidates, key=lambda item: item[0])
        max_row = max(sheet.rows) if sheet.rows else header_row
        max_col = max(headers.values())
        targets[role] = RoleTarget(role, sheet.name, '业务表头',
                                   header_row + 1, max_row, 0, max_col, headers)
    return targets


def _row_owner_id(workbook: Workbook, target: RoleTarget, row: int) -> str | None:
    sheet = workbook.sheet_by_name(target.sheet_name)
    id_column = target.headers[ID_HEADER]
    cell = sheet.cell(cell_reference(row, id_column))
    if cell is None:
        return None
    if cell.kind == 'string':
        return cell.text if cell.text else None
    if cell.kind == 'number' and cell.number is not None:
        return format(cell.number, 'f')
    return None


def _build_owners(workbook: Workbook, targets: dict[str, RoleTarget]) -> dict[tuple[str, int], str]:
    owners: dict[tuple[str, int], str] = {}
    for target in targets.values():
        for row in range(target.start_row, target.end_row + 1):
            owner = _row_owner_id(workbook, target, row)
            if owner is not None:
                owners[(target.sheet_name, row)] = owner
    return owners


def _find_cycle_nodes(graph: dict[tuple[str, int, int], set[tuple[str, int, int]]]) -> set:
    """三色 DFS，收集所有处在环上的公式单元格。"""
    WHITE, GRAY, BLACK = 0, 1, 2
    color = {node: WHITE for node in graph}
    cyclic: set = set()

    def visit(node, stack):
        color[node] = GRAY
        for dependency in graph.get(node, ()):
            if dependency not in color:
                continue  # 非公式单元格不进入环检测
            if color[dependency] == GRAY:
                index = stack.index(dependency)
                cyclic.update(stack[index:])
                cyclic.add(dependency)
            elif color[dependency] == WHITE:
                visit(dependency, stack + [dependency])
        color[node] = BLACK

    for node in list(color):
        if color[node] == WHITE:
            visit(node, [node])
    return cyclic


def _build_graph(workbook: Workbook):
    engine = FormulaEngine(workbook)
    graph: dict[tuple[str, int, int], set[tuple[str, int, int]]] = {}
    parse_problems: dict[tuple[str, int, int], dict] = {}
    for sheet in workbook.sheets:
        if sheet.external:
            continue
        for reference, cell in sheet.cells.items():
            if cell.formula is None:
                continue
            key = (sheet.name, *_ref_column_row(reference))
            if cell.formula_type not in (None, 'normal'):
                parse_problems[key] = {'代码': P_UNSUPPORTED,
                                       '说明': f'不支持的公式类型 {cell.formula_type}'}
                graph[key] = set()
                continue
            try:
                graph[key] = engine.dependencies(sheet.name, reference)
            except ExternalReferenceError as exc:
                parse_problems[key] = {'代码': P_EXTERNAL, '说明': str(exc)}
                graph[key] = set()
            except FormulaError as exc:
                parse_problems[key] = {'代码': P_UNSUPPORTED, '说明': str(exc)}
                graph[key] = set()
    cyclic = _find_cycle_nodes(graph)
    return engine, graph, cyclic, parse_problems


def _ref_column_row(reference: str) -> tuple[int, int]:
    from workbook_parser import split_reference
    column, row = split_reference(reference)
    return row, column


def _cache_check(cell: Cell, value: Decimal) -> dict:
    if cell.cached_text is None:
        return {'原始文本': None, '核对': '缺失'}
    try:
        cached = Decimal(cell.cached_text)
    except Exception:
        return {'原始文本': cell.cached_text, '核对': '无法比较'}
    return {'原始文本': cell.cached_text,
            '核对': '一致' if cached == value else '不一致'}


def _jsonable(value):
    if isinstance(value, Decimal):
        return format(value, 'f')
    if isinstance(value, dict):
        return {key: _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    return value


def collect(path: str, digest: str | None = None) -> dict:
    """采集一个工作簿，返回每家企业的可信字段、公式来源与拒收原因。"""
    workbook = parse_workbook(path)
    targets = _locate_roles(workbook)
    owners = _build_owners(workbook, targets)
    engine, graph, cyclic, parse_problems = _build_graph(workbook)
    apply_target = targets[APPLY_ROLE]
    org_target = targets[ORG_ROLE]
    cost_target = targets[COST_ROLE]

    registered = {
        owners[(org_target.sheet_name, row)]
        for row in range(org_target.start_row, org_target.end_row + 1)
        if (org_target.sheet_name, row) in owners
    }
    file_warnings = []
    if workbook.external_parts:
        file_warnings.append({'代码': W_EXTERNAL_REL,
                              '说明': '工作簿关系中存在外部目标: ' + ', '.join(workbook.external_parts)})

    declarations = []
    for row in range(apply_target.start_row, apply_target.end_row + 1):
        owner = owners.get((apply_target.sheet_name, row))
        if owner is None:
            continue
        record = _collect_row(
            workbook, targets, owners, engine, graph, cyclic,
            parse_problems, row, owner, owner in registered)
        declarations.append(record)

    result = {
        '来源文件': path,
        '原文摘要': digest,
        '日期制': '1904' if workbook.date1904 else '1900',
        '工作表定位': {
            role: {
                '工作表': target.sheet_name,
                '定位方式': target.located_by,
                '区域': f"{_column_letters(target.start_column)}{target.start_row}:"
                        f"{_column_letters(target.end_column)}{target.end_row + 1}",
                '表头': list(target.headers),
            }
            for role, target in targets.items()
        },
        '文件级问题': file_warnings,
        '宏工作簿': workbook.has_macros,
        '企业申报': sorted(declarations, key=lambda item: item['企业编号']),
    }
    if workbook.has_macros:
        for record in result['企业申报']:
            record['状态'] = '拒收'
            record['拒收原因'].insert(0, {'代码': P_MACRO, '说明': '容器含 vbaProject，按禁止执行宏规则整体拒收'})
    return _jsonable(result)


def _collect_row(workbook, targets, owners, engine, graph, cyclic,
                 parse_problems, row, owner, registered) -> dict:
    apply_target = targets[APPLY_ROLE]
    cost_target = targets[COST_ROLE]
    sheet = workbook.sheet_by_name(apply_target.sheet_name)
    blocking: list[dict] = []
    warnings: list[dict] = []
    fields = []
    formula_sources = []
    if not registered:
        blocking.append({'代码': P_UNREGISTERED, '字段': ID_HEADER,
                         '说明': f'企业编号 {owner} 未出现在企业表中'})

    for header, column in apply_target.headers.items():
        reference = cell_reference(row, column)
        cell = sheet.cell(reference)
        field = {'名称': header, '单元格': f'{apply_target.sheet_name}!{reference}',
                 '分类': classify(cell), '来源': '录入', '值': None}
        if header == ID_HEADER:
            field['值'] = cell.text if cell and cell.kind == 'string' else owner
            fields.append(field)
            continue
        if cell is None or cell.formula is None:
            _fill_plain_field(field, cell, header, blocking)
        else:
            _fill_formula_field(
                field, cell, workbook, engine, owners, cyclic, parse_problems,
                graph, apply_target, row, column, owner, blocking, warnings, formula_sources)
        if header == '日期' and field['值'] is not None and field['分类'] not in ('错误值', '公式不可信'):
            _convert_date_field(field, cell, workbook, blocking)
        fields.append(field)

    fee_total, fee_notes = _sum_costs(workbook, cost_target, owner, engine, cyclic, parse_problems)
    declared = next((field['值'] for field in fields
                     if field['名称'] == '合计' and field['值'] is not None), None)
    cross_check = {'费用表合计': fee_total, '申报合计': declared,
                   '一致': None, '备注': fee_notes}
    if isinstance(fee_total, Decimal) and isinstance(declared, Decimal):
        cross_check['一致'] = fee_total == declared
        if fee_total != declared and not any(
                reason['代码'] == P_FEE_MISMATCH for reason in blocking):
            blocking.append({'代码': P_FEE_MISMATCH, '字段': '合计',
                             '说明': f'申报合计 {declared} 与费用表重算合计 {fee_total} 不符'})
    return {
        '企业编号': owner,
        '状态': '拒收' if blocking else '可信',
        '字段': fields,
        '费用复核': _jsonable(cross_check),
        '公式来源': formula_sources,
        '拒收原因': blocking,
        '问题提示': warnings,
    }


def _fill_plain_field(field, cell, header, blocking):
    kind = classify(cell)
    if kind == '错误值':
        field['值'] = cell.error
        blocking.append({'代码': P_ERROR, '字段': header,
                         '说明': f'{field["单元格"]} 是错误值 {cell.error}'})
        return
    if kind in ('空单元格', '空字符串'):
        field['值'] = None
        return
    if cell.kind == 'boolean':
        field['值'] = cell.boolean
        return
    if cell.kind == 'number':
        field['值'] = cell.number
        return
    field['值'] = cell.text


def _convert_date_field(field, cell, workbook, blocking):
    if field.get('来源') == '公式':
        serial = field.get('值')
    else:
        serial = cell.number if cell is not None and cell.kind == 'number' else None
    if serial is None:
        return
    field['序列值'] = format(serial, 'f')
    try:
        converted = convert_serial(serial, workbook.date1904)
    except DateError as exc:
        field['值'] = None
        field['分类'] = '日期不可信'
        code = P_FAKE_LEAP if '虚构闰日' in str(exc) else P_BAD_DATE
        blocking.append({'代码': code, '字段': '日期', '说明': str(exc)})
        return
    field['值'] = converted.isoformat()
    field['分类'] = '日期'


def _upstream_taint(key, workbook, graph, cyclic, parse_problems, owners, owner):
    """沿依赖图寻找该公式单元格上游已确定的拒收根因。"""
    visited: set = set()
    queue = list(graph.get(key, ()))
    while queue:
        dependency = queue.pop()
        if dependency in visited:
            continue
        visited.add(dependency)
        dep_sheet, dep_row, dep_col = dependency
        if dependency in cyclic:
            return {'代码': P_CYCLE, '说明': f'上游 {dep_sheet}!{cell_reference(dep_row, dep_col)} 处于循环引用'}
        if dependency in parse_problems:
            origin = parse_problems[dependency]
            return {'代码': origin['代码'], '说明': f"上游 {dep_sheet}!{cell_reference(dep_row, dep_col)} {origin['说明']}"}
        dep_owner = owners.get((dep_sheet, dep_row), '无主行')
        if dep_owner not in ('无主行', owner):
            return {'代码': P_CROSS,
                    '说明': f'上游 {dep_sheet}!{cell_reference(dep_row, dep_col)} 归属企业 {dep_owner}'}
        target_sheet = workbook.sheet_by_name(dep_sheet)
        target_cell = target_sheet.cell(cell_reference(dep_row, dep_col))
        if target_cell is not None and target_cell.kind == 'error':
            return {'代码': P_ERROR,
                    '说明': f'上游 {dep_sheet}!{cell_reference(dep_row, dep_col)} 是错误值 {target_cell.error}'}
        queue.extend(graph.get(dependency, ()))
    return None


def _fill_formula_field(field, cell, workbook, engine, owners, cyclic, parse_problems,
                        graph, apply_target, row, column, owner, blocking, warnings, formula_sources):
    key = (apply_target.sheet_name, row, column)
    field['来源'] = '公式'
    field['公式'] = cell.formula
    dependencies = sorted(
        (dep_sheet, cell_reference(dep_row, dep_col), owners.get((dep_sheet, dep_row), '无主行'))
        for dep_sheet, dep_row, dep_col in graph.get(key, set())
    )
    provenance = {'字段': field['名称'], '单元格': field['单元格'], '公式': cell.formula,
                  '依赖': [{'工作表': item[0], '单元格': item[1], '归属企业': item[2]}
                          for item in dependencies]}
    formula_sources.append(provenance)
    field['缓存核对'] = {'原始文本': cell.cached_text, '核对': '未采信'}
    provenance['缓存核对'] = field['缓存核对']

    if key in parse_problems:
        problem = dict(parse_problems[key], 字段=field['名称'])
        blocking.append(problem)
    elif key in cyclic:
        blocking.append({'代码': P_CYCLE, '字段': field['名称'],
                         '说明': f'{field["单元格"]} 与依赖构成循环引用'})
    else:
        cross = [dep for dep in dependencies if dep[2] not in ('无主行', owner)]
        if cross:
            blocking.append({'代码': P_CROSS, '字段': field['名称'],
                             '说明': '引用了其他企业的行: ' +
                                    ', '.join(f'{dep[0]}!{dep[1]}({dep[2]})' for dep in cross)})
        else:
            taint = _upstream_taint(key, workbook, graph, cyclic, parse_problems, owners, owner)
            if taint is not None:
                blocking.append({**taint, '字段': field['名称']})
            else:
                try:
                    value = engine.evaluate_cell(apply_target.sheet_name, cell.reference)
                except FormulaError as exc:
                    blocking.append({'代码': P_EVAL_FAILED, '字段': field['名称'], '说明': str(exc)})
                else:
                    field['值'] = value
                    field['分类'] = '零' if value == 0 else '数值'
                    cache = _cache_check(cell, value)
                    field['缓存核对'] = cache
                    provenance['重算值'] = format(value, 'f')
                    provenance['缓存核对'] = cache
                    if cache['核对'] == '不一致':
                        warnings.append({'代码': W_STALE_CACHE, '字段': field['名称'],
                                         '说明': f'缓存 {cache["原始文本"]} 与重算值 {format(value, "f")} 不符，已采用重算值'})
                    elif cache['核对'] == '缺失':
                        warnings.append({'代码': W_NO_CACHE, '字段': field['名称'],
                                         '说明': '公式没有缓存值，已现场重算'})
                    return
    field['值'] = None
    field['分类'] = '公式不可信'


def _sum_costs(workbook, cost_target, owner, engine, cyclic, parse_problems):
    sheet = workbook.sheet_by_name(cost_target.sheet_name)
    money_column = cost_target.headers.get('金额')
    if money_column is None:
        return None, ['费用表缺少金额表头']
    total = Decimal(0)
    notes = []
    for row in range(cost_target.start_row, cost_target.end_row + 1):
        row_owner = _row_owner_id(workbook, cost_target, row)
        if row_owner != owner:
            continue
        cell = sheet.cell(cell_reference(row, money_column))
        if cell is None:
            continue
        if cell.kind == 'error':
            notes.append(f'{cell.reference} 为错误值 {cell.error}')
            continue
        if cell.formula is not None:
            key = (cost_target.sheet_name, row, money_column)
            if key in cyclic or key in parse_problems:
                notes.append(f'{cell.reference} 公式不可信，已跳过')
                continue
            try:
                total += engine.evaluate_cell(cost_target.sheet_name, cell.reference)
            except FormulaError as exc:
                notes.append(f'{cell.reference} 重算失败: {exc}')
            continue
        if cell.kind == 'number' and cell.number is not None:
            total += cell.number
    return total, notes
