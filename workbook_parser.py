"""只读 OOXML 工作簿解析。

只解压 zip 并读取 XML，绝不执行公式、宏，也不刷新外链。输出的内存模型保留
单元格的原始形态：数值（Decimal）、文本、布尔、错误值、空字符串，以及缺失单元格
（由调用方按空单元格处理），从而让业务层区分空单元格、空字符串、错误值和零。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import PurePosixPath
from zipfile import ZipFile
from xml.etree import ElementTree as ET

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

_CELL_REF_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')
_REF_TEXT_RE = re.compile(r"^(?:'((?:[^']|'')*)'|([^'!]+))!(.+)$")


def column_index(letters: str) -> int:
    """A -> 0，B -> 1，AA -> 26。"""
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def split_reference(reference: str) -> tuple[int, int]:
    """把 ``$C$4`` 之类的引用拆成零基 (列, 行)。"""
    match = _CELL_REF_RE.match(reference.replace('$', '')) if '$' in reference else _CELL_REF_RE.match(reference)
    if match is None:
        raise ValueError(f'无法识别的单元格引用: {reference}')
    return column_index(match.group(1)), int(match.group(2)) - 1


def parse_ref_text(text: str) -> tuple[str | None, int, int, int, int]:
    """解析定义名称/公式里的区域文本，如 ``费用明细!$C$2:$C$3``。

    返回 (工作表名, 起始列, 起始行, 结束列, 结束行)，行列均为零基。
    """
    match = _REF_TEXT_RE.match(text.strip())
    if match is None:
        sheet_name = None
        rest = text.strip()
    else:
        sheet_name = match.group(1) if match.group(1) is not None else match.group(2)
        if sheet_name is not None:
            sheet_name = sheet_name.replace("''", "'")
        rest = match.group(3)
    positions = rest.split(':')
    start_col, start_row = split_reference(positions[0])
    if len(positions) == 1:
        end_col, end_row = start_col, start_row
    else:
        end_col, end_row = split_reference(positions[1])
    return sheet_name, start_col, start_row, end_col, end_row


@dataclass
class Cell:
    reference: str
    kind: str                      # number | string | boolean | error
    number: Decimal | None = None
    text: str | None = None
    boolean: bool | None = None
    error: str | None = None
    formula: str | None = None
    formula_type: str | None = None   # normal | shared | array
    cached_text: str | None = None    # <v> 的原始文本，没有缓存即 None

    @property
    def is_empty_string(self) -> bool:
        return self.kind == 'string' and self.text == ''


@dataclass
class SheetPart:
    name: str
    rel_id: str
    target: str                 # 关系文件里的原始 Target
    part_path: str              # 容器内的实际部件路径
    external: bool
    cells: dict[str, Cell] = field(default_factory=dict)
    rows: dict[int, list[str]] = field(default_factory=dict)

    def cell(self, reference: str) -> Cell | None:
        return self.cells.get(reference.upper())


@dataclass
class Workbook:
    path: str
    members: int
    date1904: bool
    sheets: list[SheetPart]
    defined_names: dict[str, str]
    external_parts: list[str]
    has_macros: bool

    def sheet_by_name(self, name: str) -> SheetPart | None:
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None


def _validate_container(archive: ZipFile) -> None:
    entries = archive.infolist()
    if len({entry.filename for entry in entries}) != len(entries):
        raise ValueError('容器出现重名成员')
    if sum(entry.file_size for entry in entries) > 8 * 1024 * 1024:
        raise ValueError('工作簿展开尺寸超限')
    for entry in entries:
        name = PurePosixPath(entry.filename)
        if name.is_absolute() or '..' in name.parts or '\\' in entry.filename:
            raise ValueError('容器路径不安全')


def _resolve_part(target: str) -> str:
    if target.startswith('/'):
        return target.lstrip('/')
    return str(PurePosixPath('xl').joinpath(target))


def _read_shared_strings(archive: ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    strings = []
    for item in root.findall(f'{{{MAIN}}}si'):
        strings.append(''.join(node.text or '' for node in item.iter(f'{{{MAIN}}}t')))
    return strings


def _parse_cell(element: ET.Element, shared: list[str]) -> Cell:
    reference = element.attrib['r']
    cell_type = element.attrib.get('t', 'n')
    formula_node = element.find(f'{{{MAIN}}}f')
    formula = formula_node.text if formula_node is not None else None
    formula_type = formula_node.attrib.get('t') if formula_node is not None else None
    value_node = element.find(f'{{{MAIN}}}v')
    cached = value_node.text if value_node is not None else None
    cell = Cell(reference=reference, kind='number', formula=formula,
                formula_type=formula_type, cached_text=cached)
    if cell_type == 'inlineStr':
        inline = element.find(f'{{{MAIN}}}is')
        text = '' if inline is None else ''.join(
            node.text or '' for node in inline.iter(f'{{{MAIN}}}t'))
        cell.kind = 'string'
        cell.text = text
    elif cell_type == 's':
        cell.kind = 'string'
        if cached is None:
            cell.text = ''
        else:
            index = int(cached)
            cell.text = shared[index] if 0 <= index < len(shared) else ''
    elif cell_type == 'str':
        cell.kind = 'string'
        cell.text = cached if cached is not None else ''
    elif cell_type == 'b':
        cell.kind = 'boolean'
        cell.boolean = cached == '1'
    elif cell_type == 'e':
        cell.kind = 'error'
        cell.error = cached or '#ERR'
    else:
        cell.kind = 'number'
        if cached is not None:
            cell.number = Decimal(cached)
    return cell


def _parse_worksheet(xml_bytes: bytes, shared: list[str]) -> tuple[dict[str, Cell], dict[int, list[str]]]:
    root = ET.fromstring(xml_bytes)
    cells: dict[str, Cell] = {}
    rows: dict[int, list[str]] = {}
    for row_node in root.iter(f'{{{MAIN}}}row'):
        row_index = int(row_node.attrib['r']) - 1
        refs = []
        for cell_node in row_node.findall(f'{{{MAIN}}}c'):
            parsed = _parse_cell(cell_node, shared)
            cells[parsed.reference.upper()] = parsed
            refs.append(parsed.reference.upper())
        rows[row_index] = refs
    return cells, rows


def parse_workbook(path: str) -> Workbook:
    """解析本地 xlsx 容器，返回结构化工作簿模型。"""
    with ZipFile(path) as archive:
        _validate_container(archive)
        names = set(archive.namelist())
        has_macros = any(PurePosixPath(name).name.startswith('vbaProject') for name in names)
        shared = _read_shared_strings(archive)
        workbook_root = ET.fromstring(archive.read('xl/workbook.xml'))
        rels_root = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        rels = {entry.attrib['Id']: entry.attrib for entry in rels_root}
        external_parts = [
            entry.attrib['Target'] for entry in rels_root
            if entry.attrib.get('TargetMode') == 'External'
        ]
        properties = workbook_root.find(f'{{{MAIN}}}workbookPr')
        date1904 = properties is not None and properties.get('date1904', '0') == '1'
        defined_names = {
            node.attrib['name']: node.text or ''
            for node in workbook_root.findall(f'.//{{{MAIN}}}definedName')
        }
        sheets: list[SheetPart] = []
        for node in workbook_root.findall(f'.//{{{MAIN}}}sheet'):
            rel_id = node.attrib[f'{{{REL}}}id']
            relation = rels[rel_id]
            external = relation.get('TargetMode') == 'External'
            target = relation['Target']
            part = SheetPart(
                name=node.attrib['name'], rel_id=rel_id, target=target,
                part_path='' if external else _resolve_part(target), external=external)
            if not external:
                part.cells, part.rows = _parse_worksheet(archive.read(part.part_path), shared)
            sheets.append(part)
        return Workbook(
            path=path, members=len(names), date1904=date1904, sheets=sheets,
            defined_names=defined_names, external_parts=external_parts, has_macros=has_macros)
