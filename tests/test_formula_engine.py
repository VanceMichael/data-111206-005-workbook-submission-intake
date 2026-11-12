import unittest
from decimal import Decimal

from formula_engine import (
    ExternalReferenceError,
    FormulaEngine,
    FormulaError,
    parse_formula,
    tokenize,
)
from workbook_parser import parse_workbook


class FormulaParserTests(unittest.TestCase):
    def setUp(self):
        self.defined_names = {'甲费用': '费用明细!$C$2:$C$3'}

    def test_arithmetic_ast(self):
        node = parse_formula('1+2*3', self.defined_names)
        self.assertEqual(node.operator, '+')

    def test_named_range(self):
        node = parse_formula('SUM(甲费用)', self.defined_names)
        self.assertEqual(node.argument.name, '甲费用')

    def test_absolute_and_sheet_prefixed_references(self):
        node = parse_formula('费用明细!$C$4+B2', self.defined_names)
        self.assertEqual(node.left.sheet, '费用明细')
        self.assertEqual(node.left.row, 3)
        self.assertIsNone(node.right.sheet)

    def test_external_reference_rejected_at_tokenize(self):
        with self.assertRaises(ExternalReferenceError):
            tokenize("'[remote.xlsx]Sheet1'!A1")

    def test_unsupported_function(self):
        with self.assertRaises(FormulaError):
            parse_formula('IF(A1>0,1,2)', self.defined_names)

    def test_unsupported_operator(self):
        with self.assertRaises(FormulaError):
            tokenize('A1^2')

    def test_range_only_inside_sum(self):
        with self.assertRaises(FormulaError):
            parse_formula('A1:B2+1', self.defined_names)


class FormulaEngineTests(unittest.TestCase):
    def setUp(self):
        self.workbook = parse_workbook('fixtures/cached.xlsx')
        self.engine = FormulaEngine(self.workbook)

    def test_sum_named_range_recomputes(self):
        # 缓存写的是 999.00，受信重算必须得到费用明细 120.10+20.20。
        self.assertEqual(self.engine.evaluate_cell('申请登记', 'B2'), Decimal('140.30'))

    def test_sheet_prefixed_reference(self):
        self.assertEqual(self.engine.evaluate_cell('申请登记', 'B3'), Decimal('15.00'))

    def test_arithmetic_chain(self):
        self.assertEqual(self.engine.evaluate_cell('申请登记', 'D2'), Decimal('141.30'))

    def test_blank_cell_is_zero(self):
        # D3 在申报表中缺失，引用空单元格按 0 处理。
        self.assertEqual(self.engine.evaluate_cell('申请登记', 'D3'), Decimal(0))

    def test_division_by_zero(self):
        node = parse_formula('1/0', {})
        with self.assertRaises(FormulaError):
            self.engine._evaluate_node(node, '申请登记')

    def test_dependency_graph_includes_named_range_expansion(self):
        dependencies = self.engine.dependencies('申请登记', 'B2')
        self.assertEqual(dependencies, {('费用明细', 1, 2), ('费用明细', 2, 2)})

    def test_cycle_is_guarded(self):
        cyclic = parse_workbook('fixtures/cycle.xlsx')
        engine = FormulaEngine(cyclic)
        with self.assertRaises(FormulaError):
            engine.evaluate_cell('申请登记', 'B2')

    def test_external_formula_fails_without_touching_cache(self):
        external = parse_workbook('fixtures/external.xlsx')
        with self.assertRaises(ExternalReferenceError):
            FormulaEngine(external).dependencies('申请登记', 'B2')


if __name__ == '__main__':
    unittest.main()
