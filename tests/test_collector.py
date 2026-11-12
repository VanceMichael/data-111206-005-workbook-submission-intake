import unittest
from decimal import Decimal

from collector import classify, collect
from workbook_parser import parse_workbook


def record_by_id(result, enterprise_id):
    return next(record for record in result['企业申报'] if record['企业编号'] == enterprise_id)


def field_by_name(record, name):
    return next(field for field in record['字段'] if field['名称'] == name)


class CachedValueTests(unittest.TestCase):
    def setUp(self):
        self.result = collect('fixtures/cached.xlsx', digest='cached-digest')

    def test_stale_cache_is_not_trusted_for_total(self):
        record = record_by_id(self.result, '0007')
        total = field_by_name(record, '合计')
        # 供应商读到的 999.00 是上一版缓存；重算必须以费用明细为准。
        self.assertEqual(total['值'], '140.30')
        self.assertEqual(total['缓存核对'], {'原始文本': '999.00', '核对': '不一致'})
        stale = [warning for warning in record['问题提示'] if warning['代码'] == '缓存与重算不一致']
        self.assertEqual(len(stale), 2)

    def test_cache_matching_recompute_is_marked_consistent(self):
        record = record_by_id(self.result, '0008')
        total = field_by_name(record, '合计')
        self.assertEqual(total['值'], '15.00')
        self.assertEqual(total['缓存核对']['核对'], '一致')
        self.assertEqual(record['状态'], '可信')

    def test_leading_zero_identifier_is_text(self):
        record = record_by_id(self.result, '0007')
        identifier = field_by_name(record, '企业编号')
        self.assertEqual(identifier['值'], '0007')
        self.assertEqual(identifier['分类'], '文本')

    def test_money_is_decimal_not_float(self):
        workbook = parse_workbook('fixtures/cached.xlsx')
        sheet = workbook.sheet_by_name('费用明细')
        self.assertEqual(sheet.cell('C2').number, Decimal('120.10'))
        self.assertIsInstance(sheet.cell('C2').number, Decimal)

    def test_formula_provenance_lists_dependencies(self):
        record = record_by_id(self.result, '0007')
        source = next(item for item in record['公式来源'] if item['字段'] == '合计')
        self.assertEqual(source['公式'], 'SUM(甲费用)')
        self.assertEqual({(dep['工作表'], dep['单元格']) for dep in source['依赖']},
                         {('费用明细', 'C2'), ('费用明细', 'C3')})
        self.assertEqual(source['重算值'], '140.30')


class DateEpochRegressionTests(unittest.TestCase):
    def test_1900_fictional_leap_day_rejected(self):
        result = collect('fixtures/cached.xlsx', digest='d')
        record = record_by_id(result, '0007')
        date_field = field_by_name(record, '日期')
        self.assertIsNone(date_field['值'])
        self.assertIn('虚构闰日', [reason['代码'] for reason in record['拒收原因']])

    def test_1904_same_serial_is_valid_and_different(self):
        result_1904 = collect('fixtures/epoch1904.xlsx', digest='d')
        result_1900 = collect('fixtures/cached.xlsx', digest='d')
        date_1904 = field_by_name(record_by_id(result_1904, '0007'), '日期')
        date_1900 = field_by_name(record_by_id(result_1900, '0007'), '日期')
        self.assertEqual(date_1904['值'], '1904-03-01')
        self.assertEqual(date_1904['序列值'], '60')
        self.assertIsNone(date_1900['值'])
        self.assertEqual(result_1904['日期制'], '1904')
        self.assertEqual(result_1900['日期制'], '1900')


class CycleAndIsolationTests(unittest.TestCase):
    def test_cycle_isolated_to_affected_enterprise(self):
        result = collect('fixtures/cycle.xlsx', digest='d')
        tainted = record_by_id(result, '0007')
        clean = record_by_id(result, '0008')
        self.assertEqual(tainted['状态'], '拒收')
        codes = [reason['代码'] for reason in tainted['拒收原因']]
        self.assertIn('循环依赖', codes)
        for field in tainted['字段']:
            if field['来源'] == '公式':
                self.assertEqual(field['缓存核对']['核对'], '未采信')
        # 无关企业不得被连带丢弃。
        self.assertEqual(clean['状态'], '可信')
        self.assertEqual(field_by_name(clean, '合计')['值'], '15.00')

    def test_three_node_cycle_is_detected(self):
        import tempfile
        import zipfile
        from pathlib import Path
        main_ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

        def formula_cell(ref, expression):
            return f'<c r="{ref}"><f>{expression}</f><v>0</v></c>'

        apply_xml = (
            f'<worksheet xmlns="{main_ns}"><sheetData>'
            '<row r="1"><c r="A1" t="inlineStr"><is><t>企业编号</t></is></c>'
            '<c r="B1" t="inlineStr"><is><t>合计</t></is></c></row>'
            f'<row r="2"><c r="A2" t="inlineStr"><is><t>0007</t></is></c>'
            f'{formula_cell("B2", "B3")}{formula_cell("C2", "B3")}</row>'
            f'<row r="3"><c r="A3" t="inlineStr"><is><t>0007</t></is></c>{formula_cell("B3", "C2")}</row>'
            f'<row r="4"><c r="A4" t="inlineStr"><is><t>0008</t></is></c>{formula_cell("B4", "费用明细!C4")}</row>'
            '</sheetData></worksheet>'
        )
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'ring.xlsx'
            with zipfile.ZipFile('fixtures/cached.xlsx') as source, \
                    zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as output:
                for item in source.infolist():
                    body = source.read(item.filename)
                    if item.filename == 'xl/worksheets/apply.xml':
                        body = apply_xml.encode()
                    elif item.filename == 'xl/workbook.xml':
                        body = body.replace(b'$A$1:$D$3', b'$A$1:$D$4')
                    output.writestr(item, body)
            result = collect(str(target), digest='d')
            tainted = [record for record in result['企业申报'] if record['企业编号'] == '0007']
            clean = record_by_id(result, '0008')
            self.assertTrue(tainted)
            self.assertTrue(
                all('循环依赖' in [reason['代码'] for reason in record['拒收原因']]
                    for record in tainted))
            self.assertEqual(clean['状态'], '可信')

    def test_external_link_isolated_and_other_enterprise_survives(self):
        result = collect('fixtures/external.xlsx', digest='d')
        tainted = record_by_id(result, '0007')
        clean = record_by_id(result, '0008')
        self.assertEqual(tainted['状态'], '拒收')
        reasons = [reason['代码'] for reason in tainted['拒收原因']]
        self.assertIn('外链公式', reasons)
        # 下游公式继承同一根因，而不是另算一个笼统的重算失败。
        self.assertIn('外链公式', reasons)
        self.assertEqual(clean['状态'], '可信')

    def test_cross_enterprise_reference_rejected_only_for_owner(self):
        result = collect('fixtures/cross.xlsx', digest='d')
        clean = record_by_id(result, '0007')
        tainted = record_by_id(result, '0008')
        self.assertEqual(clean['状态'], '可信')
        self.assertEqual(tainted['状态'], '拒收')
        self.assertIn('跨企业引用', [reason['代码'] for reason in tainted['拒收原因']])


class EmptyStateTests(unittest.TestCase):
    def test_four_cell_states_are_distinct(self):
        result = collect('fixtures/blanks.xlsx', digest='d')
        first = record_by_id(result, '0007')
        second = record_by_id(result, '0008')
        # 零、空字符串、空单元格、错误值四种形态必须可区分。
        self.assertEqual(field_by_name(second, '合计')['分类'], '零')
        self.assertEqual(field_by_name(second, '合计')['值'], '0')
        self.assertEqual(field_by_name(first, '校验')['分类'], '空字符串')
        self.assertEqual(field_by_name(second, '日期')['分类'], '空单元格')
        self.assertEqual(field_by_name(second, '校验')['分类'], '错误值')
        self.assertEqual(field_by_name(second, '校验')['值'], '#DIV/0!')
        self.assertIn('错误值', [reason['代码'] for reason in second['拒收原因']])

    def test_missing_cache_is_recomputed_and_flagged(self):
        result = collect('fixtures/blanks.xlsx', digest='d')
        record = record_by_id(result, '0007')
        total = field_by_name(record, '合计')
        self.assertEqual(record['状态'], '可信')
        self.assertEqual(total['值'], '140.30')
        self.assertEqual(total['缓存核对']['核对'], '缺失')
        self.assertIn('缺少缓存值', [warning['代码'] for warning in record['问题提示']])

    def test_fee_cross_check_catches_declared_zero(self):
        result = collect('fixtures/blanks.xlsx', digest='d')
        second = record_by_id(result, '0008')
        self.assertFalse(second['费用复核']['一致'])
        self.assertIn('费用复核不一致', [reason['代码'] for reason in second['拒收原因']])


class SheetLocationTests(unittest.TestCase):
    def test_renamed_sheets_located_via_relationships_and_defined_names(self):
        result = collect('fixtures/renamed.xlsx', digest='d')
        located = result['工作表定位']
        self.assertEqual(located['费用表']['工作表'], 'cost_2026')
        self.assertEqual(located['申报表']['工作表'], 'apply_2026')
        clean = record_by_id(result, '0008')
        self.assertEqual(clean['状态'], '可信')
        self.assertEqual(field_by_name(clean, '合计')['值'], '15.00')

    def test_header_fallback_when_defined_names_removed(self):
        # 构造一个删除定义名称的工作簿，验证业务表头兜底定位。
        from zipfile import ZipFile
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'noname.xlsx'
            with ZipFile('fixtures/cached.xlsx') as source, ZipFile(target, 'w') as output:
                for item in source.infolist():
                    body = source.read(item.filename)
                    if item.filename == 'xl/workbook.xml':
                        text = body.decode()
                        start = text.index('<definedNames>')
                        body = (text[:start] + text[text.index('</definedNames>') + len('</definedNames>'):]).encode()
                    output.writestr(item, body)
            result = collect(str(target), digest='d')
            self.assertEqual(result['工作表定位']['申报表']['定位方式'], '业务表头')
            self.assertEqual(record_by_id(result, '0008')['状态'], '可信')


class CellStateClassificationTests(unittest.TestCase):
    def test_classify_distinguishes_blank_empty_string_error_and_zero(self):
        from workbook_parser import Cell
        from decimal import Decimal
        self.assertEqual(classify(None), '空单元格')
        self.assertEqual(classify(Cell('A1', 'string', text='')), '空字符串')
        self.assertEqual(classify(Cell('A1', 'error', error='#DIV/0!')), '错误值')
        self.assertEqual(classify(Cell('A1', 'number', number=Decimal('0'))), '零')
        self.assertEqual(classify(Cell('A1', 'number', number=Decimal('0.00'))), '零')
        self.assertEqual(classify(Cell('A1', 'number', number=Decimal('120.10'))), '数值')


if __name__ == '__main__':
    unittest.main()
