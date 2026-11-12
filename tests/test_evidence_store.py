import tempfile
import unittest
import zipfile
from pathlib import Path

from collector import collect
from evidence_store import EvidenceStore


def modify_cost_value(path: Path) -> bytes:
    """改动费用明细依赖（120.10 -> 130.10），模拟同一文件依赖被修改。"""
    buffer = tempfile.SpooledTemporaryFile()
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as output:
        for item in source.infolist():
            body = source.read(item.filename)
            if item.filename == 'xl/worksheets/cost.xml':
                body = body.replace(b'<v>120.10</v>', b'<v>130.10</v>', 1)
            output.writestr(item, body)
    buffer.seek(0)
    return buffer.read()


class EvidenceStoreTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.db_path = Path(self.directory.name) / 'evidence.sqlite3'
        self.store = EvidenceStore(self.db_path)
        self.workbook = Path('fixtures/cached.xlsx')

    def tearDown(self):
        self.store.close()
        self.directory.cleanup()

    def _result(self, body, digest):
        source = Path(self.directory.name) / 'book.xlsx'
        source.write_bytes(body)
        return collect(str(source), digest=digest)

    def test_repeated_import_reuses_batch_and_keeps_observations(self):
        body = self.workbook.read_bytes()
        first = self.store.save_batch('book.xlsx', self._result(body, 'digest-1'), '2026-11-01T00:00:00Z')
        second = self.store.save_batch('book.xlsx', self._result(body, 'digest-1'), '2026-11-02T00:00:00Z')
        self.assertEqual(first['批次号'], second['批次号'])
        self.assertEqual(first['处理'], '新建批次')
        self.assertEqual(second['处理'], '复用批次')
        self.assertEqual(len(self.store.list_batches()), 1)
        history = self.store.evidence_for('0007')
        self.assertEqual(len(history), 1)

    def test_modified_dependency_creates_new_evidence(self):
        original = self.workbook.read_bytes()
        modified = modify_cost_value(self.workbook)
        first = self.store.save_batch('book.xlsx', self._result(original, 'digest-1'), '2026-11-01T00:00:00Z')
        second = self.store.save_batch('book.xlsx', self._result(modified, 'digest-2'), '2026-11-03T00:00:00Z')
        self.assertNotEqual(first['批次号'], second['批次号'])
        batches = self.store.list_batches()
        self.assertEqual(len(batches), 2)
        history = self.store.evidence_for('0007')
        self.assertEqual(len(history), 2)
        totals = [
            next(field for field in item['记录']['字段'] if field['名称'] == '合计')['值']
            for item in history
        ]
        self.assertEqual(totals, ['140.30', '150.30'])

    def test_rejected_enterprise_is_persisted_with_reason(self):
        result = self._result(self.workbook.read_bytes(), 'digest-1')
        self.store.save_batch('book.xlsx', result, '2026-11-01T00:00:00Z')
        history = self.store.evidence_for('0007')
        self.assertEqual(history[0]['状态'], '拒收')
        self.assertTrue(history[0]['记录']['拒收原因'])

    def test_batch_round_trip(self):
        result = self._result(self.workbook.read_bytes(), 'digest-1')
        saved = self.store.save_batch('book.xlsx', result, '2026-11-01T00:00:00Z')
        loaded = self.store.get_batch(saved['批次号'])
        self.assertEqual(loaded['原文摘要'], 'digest-1')
        self.assertEqual(len(loaded['企业申报']), 2)


class MacroAndContainerTests(unittest.TestCase):
    def test_macro_part_marks_workbook(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'macro.xlsx'
            with zipfile.ZipFile('fixtures/cached.xlsx') as source, zipfile.ZipFile(target, 'w') as output:
                for item in source.infolist():
                    output.writestr(item, source.read(item.filename))
                output.writestr('xl/vbaProject.bin', b'fake macro project')
            result = collect(str(target), digest='macro')
            self.assertTrue(result['宏工作簿'])
            self.assertTrue(all(record['状态'] == '拒收' for record in result['企业申报']))
            self.assertTrue(
                all(any(reason['代码'] == '宏工作簿拒收' for reason in record['拒收原因'])
                    for record in result['企业申报']))


if __name__ == '__main__':
    unittest.main()
