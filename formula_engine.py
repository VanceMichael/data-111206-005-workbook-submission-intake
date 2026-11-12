"""受限 Excel 公式引擎。

支持范围严格限定为：
  * 四则运算 ``+ - * /``、括号、正负号、数字常量；
  * 同簿单元格引用（``B2``、``费用明细!C4``、``$A$1``）；
  * 区域引用，且只允许出现在 ``SUM`` 内（含命名区域展开后的区域）；
  * ``SUM`` 与工作簿定义名称（命名区域）。

外链、其他函数、共享/数组公式、百分比、乘方等一律视为不支持，由业务层隔离申报，
而不是偷偷采用缓存值。引擎本身不做业务决策（循环由依赖图显式报告）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, getcontext

from workbook_parser import Workbook, parse_ref_text, split_reference

getcontext().prec = 28

# 引用/名称优先：允许带工作表前缀、绝对引用 $、含中文与点的名称，以及带括号的外链。
_TOKEN_RE = re.compile(
    r"""
    (?P<space>\s+)
  | (?P<number>\d+(?:\.\d+)?|\.\d+)
  | (?P<ref>(?:(?:'(?:[^']|'')*'|[A-Za-z0-9_\.一-鿿]+)!)?
          \$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<name>[A-Za-z_][A-Za-z0-9_\.]*|[一-鿿][A-Za-z0-9_\.一-鿿]*)
  | (?P<op>[+\-*/(),:])
    """,
    re.VERBOSE,
)
_REF_PART_RE = re.compile(r'^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$')


class FormulaError(Exception):
    """公式不受支持或无法求值。"""


class ExternalReferenceError(FormulaError):
    """引用了外部工作簿。"""


@dataclass(frozen=True)
class CellRef:
    sheet: str | None
    column: int
    row: int


@dataclass(frozen=True)
class RangeRef:
    sheet: str | None
    start_column: int
    start_row: int
    end_column: int
    end_row: int


@dataclass(frozen=True)
class NameRef:
    name: str


@dataclass(frozen=True)
class Number:
    value: Decimal


@dataclass(frozen=True)
class Unary:
    operator: str
    operand: object


@dataclass(frozen=True)
class Binary:
    operator: str
    left: object
    right: object


@dataclass(frozen=True)
class SumCall:
    argument: object


def tokenize(formula: str) -> list[tuple[str, str]]:
    if '[' in formula or ']' in formula:
        raise ExternalReferenceError(f'公式包含外部工作簿引用: {formula}')
    tokens: list[tuple[str, str]] = []
    position = 0
    while position < len(formula):
        match = _TOKEN_RE.match(formula, position)
        if match is None:
            raise FormulaError(f'无法解析公式片段: {formula[position:]}')
        kind = match.lastgroup
        text = match.group()
        position = match.end()
        if kind != 'space':
            tokens.append((kind, text))
    return tokens


class _Parser:
    def __init__(self, tokens: list[tuple[str, str]], defined_names: dict[str, str]):
        self.tokens = tokens
        self.index = 0
        self.defined_names = defined_names

    def _peek(self) -> tuple[str, str] | None:
        return self.tokens[self.index] if self.index < len(self.tokens) else None

    def _take(self) -> tuple[str, str]:
        token = self._peek()
        if token is None:
            raise FormulaError('公式意外结束')
        self.index += 1
        return token

    def parse(self) -> object:
        node = self._expression()
        if self._peek() is not None:
            raise FormulaError(f'公式存在多余记号: {self._peek()[1]}')
        return node

    def _expression(self) -> object:
        node = self._term()
        while self._peek() is not None and self._peek()[1] in ('+', '-'):
            operator = self._take()[1]
            node = Binary(operator, node, self._term())
        return node

    def _term(self) -> object:
        node = self._factor()
        while self._peek() is not None and self._peek()[1] in ('*', '/'):
            operator = self._take()[1]
            node = Binary(operator, node, self._factor())
        return node

    def _factor(self) -> object:
        token = self._peek()
        if token is None:
            raise FormulaError('公式意外结束')
        if token[1] in ('+', '-'):
            operator = self._take()[1]
            return Unary(operator, self._factor())
        if token[1] == '(':
            self._take()
            node = self._expression()
            if self._peek() is None or self._take()[1] != ')':
                raise FormulaError('缺少右括号')
            return node
        if token[0] == 'number':
            self._take()
            try:
                return Number(Decimal(token[1]))
            except InvalidOperation as exc:
                raise FormulaError(f'非法数字常量: {token[1]}') from exc
        if token[0] == 'name':
            return self._name_token()
        if token[0] == 'ref':
            self._take()
            reference = self._build_ref(token[1])
            if isinstance(reference, RangeRef):
                raise FormulaError(f'区域引用 {token[1]} 只能用于 SUM')
            return reference
        raise FormulaError(f'不支持的记号: {token[1]}')

    def _name_token(self) -> object:
        name = self._take()[1]
        upper = name.upper()
        if upper == 'SUM' and self._peek() is not None and self._peek()[1] == '(':
            self._take()
            argument = self._expression()
            if self._peek() is not None and self._peek()[1] == ',':
                raise FormulaError('SUM 仅支持单个区域参数')
            if self._peek() is None or self._take()[1] != ')':
                raise FormulaError('SUM 缺少右括号')
            if not isinstance(argument, (RangeRef, NameRef)):
                raise FormulaError('SUM 参数必须是区域或命名区域')
            return SumCall(argument)
        if upper in ('TRUE', 'FALSE'):
            raise FormulaError('不支持布尔常量')
        if name not in self.defined_names:
            raise FormulaError(f'未定义的名称或不支持的函数: {name}')
        return NameRef(name)

    def _build_ref(self, text: str) -> CellRef | RangeRef:
        if '!' in text:
            sheet, start_col, start_row, end_col, end_row = parse_ref_text(text)
        else:
            sheet = None
            positions = text.split(':')
            start_col, start_row = split_reference(positions[0])
            if len(positions) == 1:
                end_col, end_row = start_col, start_row
            else:
                end_col, end_row = split_reference(positions[1])
        if start_col == end_col and start_row == end_row:
            return CellRef(sheet, start_col, start_row)
        return RangeRef(sheet, start_col, start_row, end_col, end_row)


def parse_formula(formula: str, defined_names: dict[str, str]) -> object:
    return _Parser(tokenize(formula), defined_names).parse()


def _expand_name(workbook: Workbook, name: str) -> RangeRef:
    if name not in workbook.defined_names:
        raise FormulaError(f'未定义的名称: {name}')
    sheet, start_col, start_row, end_col, end_row = parse_ref_text(workbook.defined_names[name])
    if sheet is None:
        raise FormulaError(f'命名区域缺少工作表: {name}')
    return RangeRef(sheet, start_col, start_row, end_col, end_row)


def _resolve_sheet(workbook: Workbook, sheet: str | None, current_sheet: str) -> str:
    name = current_sheet if sheet is None else sheet
    if workbook.sheet_by_name(name) is None:
        raise FormulaError(f'引用了不存在的工作表: {name}')
    return name


def iter_range(workbook: Workbook, reference: RangeRef, current_sheet: str):
    sheet_name = _resolve_sheet(workbook, reference.sheet, current_sheet)
    sheet = workbook.sheet_by_name(sheet_name)
    for row in range(reference.start_row, reference.end_row + 1):
        for col in range(reference.start_column, reference.end_column + 1):
            yield sheet_name, row, col, _column_letters(col) + str(row + 1)


def _column_letters(index: int) -> str:
    index += 1
    letters = ''
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def collect_dependencies(node: object, workbook: Workbook, current_sheet: str) -> set[tuple[str, int, int]]:
    """返回 AST 直接或经命名区域展开后依赖的 (工作表, 行, 列) 集合。"""
    dependencies: set[tuple[str, int, int]] = set()

    def walk(current: object) -> None:
        if isinstance(current, CellRef):
            sheet_name = _resolve_sheet(workbook, current.sheet, current_sheet)
            dependencies.add((sheet_name, current.row, current.column))
        elif isinstance(current, RangeRef):
            for sheet_name, row, col, _ in iter_range(workbook, current, current_sheet):
                dependencies.add((sheet_name, row, col))
        elif isinstance(current, NameRef):
            for sheet_name, row, col, _ in iter_range(workbook, _expand_name(workbook, current.name), current_sheet):
                dependencies.add((sheet_name, row, col))
        elif isinstance(current, Unary):
            walk(current.operand)
        elif isinstance(current, Binary):
            walk(current.left)
            walk(current.right)
        elif isinstance(current, SumCall):
            walk(current.argument)
        elif isinstance(current, Number):
            return

    walk(node)
    return dependencies


class FormulaEngine:
    """对已解析工作簿做受信重算，不读取也不信任缓存值。"""

    def __init__(self, workbook: Workbook):
        self.workbook = workbook
        self._visiting: set[tuple[str, int, int]] = set()

    def dependencies(self, sheet_name: str, cell_reference: str) -> set[tuple[str, int, int]]:
        sheet = self.workbook.sheet_by_name(sheet_name)
        cell = sheet.cell(cell_reference)
        if cell is None or cell.formula is None:
            return set()
        node = parse_formula(cell.formula, self.workbook.defined_names)
        return collect_dependencies(node, self.workbook, sheet_name)

    def evaluate_cell(self, sheet_name: str, cell_reference: str) -> Decimal:
        sheet = self.workbook.sheet_by_name(sheet_name)
        cell = sheet.cell(cell_reference)
        if cell is None:
            return Decimal(0)
        return self._evaluate(sheet_name, cell)

    def _evaluate(self, sheet_name: str, cell) -> Decimal:
        if cell.kind == 'error':
            raise FormulaError(f'依赖单元格是错误值 {cell.error}')
        if cell.formula is None:
            if cell.kind == 'number':
                return cell.number if cell.number is not None else Decimal(0)
            if cell.kind == 'boolean':
                return Decimal(1 if cell.boolean else 0)
            if cell.kind == 'string' and cell.text == '':
                # 空字符串参与算术按 0 处理，与 Excel 兼容；空单元格同理。
                return Decimal(0)
            if cell.kind == 'string':
                raise FormulaError('文本无法参与算术运算')
            return Decimal(0)
        column, row = split_reference(cell.reference)
        guard = (sheet_name, row, column)
        if guard in self._visiting:
            raise FormulaError(f'检测到循环引用: {sheet_name}!{cell.reference}')
        self._visiting.add(guard)
        try:
            node = parse_formula(cell.formula, self.workbook.defined_names)
            return self._evaluate_node(node, sheet_name)
        finally:
            self._visiting.discard(guard)

    def _cell_at(self, sheet_name: str, row: int, column: int):
        sheet = self.workbook.sheet_by_name(sheet_name)
        return sheet.cell(_column_letters(column) + str(row + 1))

    def _evaluate_node(self, node: object, current_sheet: str) -> Decimal:
        if isinstance(node, Number):
            return node.value
        if isinstance(node, Unary):
            value = self._evaluate_node(node.operand, current_sheet)
            return value if node.operator == '+' else -value
        if isinstance(node, Binary):
            left = self._evaluate_node(node.left, current_sheet)
            right = self._evaluate_node(node.right, current_sheet)
            if node.operator == '+':
                return left + right
            if node.operator == '-':
                return left - right
            if node.operator == '*':
                return left * right
            if node.operator == '/':
                if right == 0:
                    raise FormulaError('除数为零')
                return left / right
        if isinstance(node, CellRef):
            sheet_name = _resolve_sheet(self.workbook, node.sheet, current_sheet)
            cell = self._cell_at(sheet_name, node.row, node.column)
            if cell is None:
                return Decimal(0)
            return self._evaluate(sheet_name, cell)
        if isinstance(node, RangeRef):
            raise FormulaError('区域不能直接参与四则运算')
        if isinstance(node, NameRef):
            target = _expand_name(self.workbook, node.name)
            if target.start_column != target.end_column or target.start_row != target.end_row:
                raise FormulaError('命名区域不能作为单个值参与四则运算')
            sheet_name = _resolve_sheet(self.workbook, target.sheet, current_sheet)
            cell = self._cell_at(sheet_name, target.start_row, target.start_column)
            if cell is None:
                return Decimal(0)
            return self._evaluate(sheet_name, cell)
        if isinstance(node, SumCall):
            return self._evaluate_sum(node.argument, current_sheet)
        raise FormulaError(f'不支持的表达式: {node!r}')

    def _evaluate_sum(self, argument: object, current_sheet: str) -> Decimal:
        if isinstance(argument, NameRef):
            target = _expand_name(self.workbook, argument.name)
        elif isinstance(argument, RangeRef):
            target = argument
        else:
            raise FormulaError('SUM 参数必须是区域或命名区域')
        total = Decimal(0)
        for sheet_name, row, column, _ in iter_range(self.workbook, target, current_sheet):
            cell = self._cell_at(sheet_name, row, column)
            if cell is None:
                continue
            if cell.kind == 'error':
                raise FormulaError(f'求和区域包含错误值 {cell.error}')
            if cell.formula is not None:
                total += self._evaluate(sheet_name, cell)
            elif cell.kind == 'number':
                total += cell.number if cell.number is not None else Decimal(0)
            # 文本（含空字符串）、布尔与空单元格在 SUM 中忽略。
        return total
