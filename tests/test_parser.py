import tempfile
import unittest
import zipfile
from pathlib import Path

from workbook_parser import parse_workbook

MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def build_shared_string_book(directory: str) -> Path:
    path = Path(directory) / 'shared.xlsx'
    shared_strings = (
        f'<sst xmlns="{MAIN}" count="2" uniqueCount="2">'
        '<si><t>0007</t></si><si><t>企业编号</t></si></sst>'
    )
    sheet_xml = (
        f'<worksheet xmlns="{MAIN}"><sheetData>'
        '<row r="1"><c r="A1" t="s"><v>1</v></c></row>'
        '<row r="2"><c r="A2" t="s"><v>0</v></c></row>'
        '</sheetData></worksheet>'
    )
    workbook = (
        f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><workbookPr date1904="0"/>'
        '<sheets><sheet name="表一" sheetId="1" r:id="r1"/></sheets></workbook>'
    )
    rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="r1" Type="{REL}/worksheet" Target="worksheets/s1.xml"/></Relationships>')
    root_rels = ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 f'<Relationship Id="rBook" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    content_types = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/s1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        '</Types>'
    )
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', content_types)
        archive.writestr('_rels/.rels', root_rels)
        archive.writestr('xl/workbook.xml', workbook)
        archive.writestr('xl/_rels/workbook.xml.rels', rels)
        archive.writestr('xl/worksheets/s1.xml', sheet_xml)
        archive.writestr('xl/sharedStrings.xml', shared_strings)
    return path


class SharedStringsTests(unittest.TestCase):
    def test_shared_string_index_resolves_to_text(self):
        with tempfile.TemporaryDirectory() as directory:
            path = build_shared_string_book(directory)
            workbook = parse_workbook(str(path))
            sheet = workbook.sheets[0]
            self.assertEqual(sheet.cell('A1').text, '企业编号')
            # 前导零编号即使走共享字符串也必须保留文本语义。
            self.assertEqual(sheet.cell('A2').kind, 'string')
            self.assertEqual(sheet.cell('A2').text, '0007')


if __name__ == '__main__':
    unittest.main()
