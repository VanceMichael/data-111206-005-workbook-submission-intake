import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from raw_store import RawStore


def demo(store):
    observed_at = datetime.now(timezone.utc).isoformat()
    from workbook_source import inventory
    result=[]
    for path in sorted(Path('fixtures').glob('*.xlsx')):
        found=inventory(path)
        digest=store.save(path.name,path.read_bytes(),found,observed_at)
        result.append({'文件':path.name,'工作表':found['sheets'],'日期制1904':found['date1904'],'原文摘要':digest})
    return {'工作簿目录':result,'说明':'不读取公式计算结果'}


def main():
    parser = argparse.ArgumentParser(description='本地原文采集演示')
    parser.add_argument('command', choices=['demo'])
    parser.add_argument('--output', default='output/raw.sqlite3')
    args = parser.parse_args()
    store = RawStore(args.output)
    try:
        result = demo(store)
        print(json.dumps(result, ensure_ascii=False, indent=2))
    finally:
        store.close()


if __name__ == '__main__':
    main()
