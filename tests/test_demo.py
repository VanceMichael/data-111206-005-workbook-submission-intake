import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


class DemoTests(unittest.TestCase):
    def test_real_demo(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'capture.sqlite3'
            process = subprocess.run([sys.executable, 'app.py', 'demo', '--output', str(output)],
                                     capture_output=True, text=True, timeout=30)
            self.assertEqual(process.returncode, 0, process.stderr)
            self.assertIsInstance(json.loads(process.stdout), dict)
            self.assertTrue(output.is_file())

    def test_collect_command_and_batch_reuse(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'capture.sqlite3'
            command = [sys.executable, 'app.py', 'collect', 'fixtures/cached.xlsx',
                       'fixtures/cycle.xlsx', '--output', str(output)]
            first = subprocess.run(command, capture_output=True, text=True, timeout=30)
            self.assertEqual(first.returncode, 0, first.stderr)
            payload = json.loads(first.stdout)
            self.assertEqual(len(payload['采集批次']), 2)
            by_file = {item['来源文件']: item for item in payload['采集批次']}
            self.assertEqual(by_file['fixtures/cached.xlsx']['处理'], '新建批次')
            cached_enterprises = {
                item['企业编号']: item for item in by_file['fixtures/cached.xlsx']['企业结果']
            }
            self.assertEqual(cached_enterprises['0007']['状态'], '拒收')
            self.assertIn('虚构闰日', cached_enterprises['0007']['拒收原因'])
            self.assertIn('缓存与重算不一致', cached_enterprises['0007']['问题提示'])
            self.assertEqual(cached_enterprises['0008']['状态'], '可信')
            cycle_enterprises = {
                item['企业编号']: item for item in by_file['fixtures/cycle.xlsx']['企业结果']
            }
            self.assertEqual(cycle_enterprises['0007']['状态'], '拒收')
            self.assertEqual(cycle_enterprises['0008']['状态'], '可信')

            # 同一文件重复导入必须查到原批次。
            second = subprocess.run(command, capture_output=True, text=True, timeout=30)
            self.assertEqual(second.returncode, 0, second.stderr)
            self.assertTrue(
                all(item['处理'] == '复用批次' for item in json.loads(second.stdout)['采集批次']))

            listing = subprocess.run([sys.executable, 'app.py', 'batches', '--output', str(output)],
                                     capture_output=True, text=True, timeout=30)
            self.assertEqual(listing.returncode, 0, listing.stderr)
            batches = json.loads(listing.stdout)['批次目录']
            self.assertEqual(len(batches), 2)
            self.assertEqual(sum(batch['拒收数'] for batch in batches), 2)
            self.assertEqual(sum(batch['可信数'] for batch in batches), 2)

    def test_collect_runs_under_isolated_python(self):
        # 下载文件视为不可信数据，采集入口必须能在 -I 隔离模式下运行。
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'capture.sqlite3'
            process = subprocess.run(
                [sys.executable, '-I', 'app.py', 'collect', 'fixtures/external.xlsx',
                 '--output', str(output)],
                capture_output=True, text=True, timeout=30)
            self.assertEqual(process.returncode, 0, process.stderr)
            payload = json.loads(process.stdout)
            result = payload['采集批次'][0]['企业结果']
            statuses = {item['企业编号']: item['状态'] for item in result}
            self.assertEqual(statuses, {'0007': '拒收', '0008': '可信'})


if __name__ == '__main__':
    unittest.main()
