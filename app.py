import argparse
import hashlib
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

# python -I 不会把脚本目录注入 sys.path，显式加入受信的脚本所在目录，
# 保证对不可信工作簿做隔离运行时仍能导入本项目模块。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from collector import collect
from evidence_store import EvidenceStore
from raw_store import RawStore


def _sha256(body: bytes) -> str:
    return hashlib.sha256(body).hexdigest()


def demo(store):
    observed_at = datetime.now(timezone.utc).isoformat()
    from workbook_source import inventory
    result = []
    for path in sorted(Path('fixtures').glob('*.xlsx')):
        found = inventory(path)
        digest = store.save(path.name, path.read_bytes(), found, observed_at)
        result.append({'文件': path.name, '工作表': found['sheets'],
                       '日期制1904': found['date1904'], '原文摘要': digest})
    return {'工作簿目录': result, '说明': '不读取公式计算结果'}


def collect_files(paths, raw_store: RawStore, evidence_store: EvidenceStore):
    observed_at = datetime.now(timezone.utc).isoformat()
    batches = []
    for path in paths:
        source = str(path)
        body = path.read_bytes()
        digest = _sha256(body)
        # 先存原文，再做结构化采集，保证拒收文件也能追溯到原始字节。
        raw_store.save(source, body, {'阶段': '结构化采集'}, observed_at)
        result = collect(source, digest=digest)
        batch = evidence_store.save_batch(source, result, observed_at)
        summary = {
            '来源文件': source,
            '批次号': batch['批次号'],
            '处理': batch['处理'],
            '日期制': result['日期制'],
            '文件级问题': result['文件级问题'],
            '企业结果': [
                {
                    '企业编号': record['企业编号'],
                    '状态': record['状态'],
                    '拒收原因': [reason['代码'] for reason in record['拒收原因']],
                    '问题提示': [warning['代码'] for warning in record['问题提示']],
                    '缓存不一致字段': [
                        field['名称'] for field in record['字段']
                        if field.get('缓存核对', {}).get('核对') == '不一致'
                    ],
                }
                for record in result['企业申报']
            ],
        }
        batches.append(summary)
    return {'采集批次': batches,
            '说明': '公式沿依赖重算并核对缓存；拒收仅隔离受影响企业，不丢弃同簿其他企业'}


def list_batches(evidence_store: EvidenceStore):
    return {'批次目录': evidence_store.list_batches()}


def main():
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--output', default='output/raw.sqlite3')
    parser = argparse.ArgumentParser(description='企业申报工作簿采集入口', parents=[common])
    sub = parser.add_subparsers(dest='command', required=True)

    sub.add_parser('demo', parents=[common])
    collect_parser = sub.add_parser('collect', parents=[common], help='对一个或多个 xlsx 做结构化采集')
    collect_parser.add_argument('paths', nargs='+', help='xlsx 文件路径')
    sub.add_parser('batches', parents=[common], help='列出已保存的采集批次与企业结果')
    args = parser.parse_args()

    raw_store = RawStore(args.output)
    evidence_store = EvidenceStore(args.output)
    try:
        if args.command == 'demo':
            result = demo(raw_store)
        elif args.command == 'collect':
            result = collect_files([Path(item) for item in args.paths], raw_store, evidence_store)
        else:
            result = list_batches(evidence_store)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        raw_store.close()
        evidence_store.close()


if __name__ == '__main__':
    main()
