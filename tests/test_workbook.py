import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from workbook_source import inventory


class WorkbookTests(unittest.TestCase):
    def test_sheet_relationships(self):
        result = inventory('fixtures/cached.xlsx')
        self.assertEqual(len(result['sheets']),3)
        self.assertEqual(result['sheets'][0]['target'],'worksheets/org.xml')

    def test_epoch(self):
        self.assertFalse(inventory('fixtures/cached.xlsx')['date1904'])
        self.assertTrue(inventory('fixtures/epoch1904.xlsx')['date1904'])

    def test_formula_and_cache_are_distinct(self):
        with ZipFile('fixtures/cached.xlsx') as archive:
            raw = archive.read('xl/worksheets/apply.xml').decode()
        self.assertIn('999.00',raw)
        self.assertIn('SUM(甲费用)',raw)

    def test_path_traversal_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'invalid.xlsx'
            with ZipFile(path,'w') as archive:
                archive.writestr('../escape','data')
            with self.assertRaises(ValueError):
                inventory(path)

    def test_text_identifiers_preserved_in_file(self):
        with ZipFile('fixtures/cached.xlsx') as archive:
            self.assertIn(b'0007',archive.read('xl/worksheets/org.xml'))
