"""采集批次与证据持久化。

在原文库之外记录每次导入的批次与每家企业的结构化证据：

* 同一文件字节（同一 SHA-256 摘要）重复导入复用原批次，只刷新观察时间，不制造新证据；
* 任一依赖变化导致字节摘要不同，就产生新的批次 ID 与新证据行；
* 证据 JSON 中金额以文本形式保存十进制值，企业编号保留前导零。
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
from datetime import datetime, timezone
from pathlib import Path


def canonical_json(payload) -> str:
    return json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False)


class EvidenceStore:
    def __init__(self, path: str):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(path, timeout=10)
        self.connection.execute('PRAGMA journal_mode=WAL')
        self.connection.execute('PRAGMA foreign_keys=ON')
        self.connection.executescript("""
            CREATE TABLE IF NOT EXISTS batches (
                batch_id TEXT PRIMARY KEY,
                source TEXT NOT NULL,
                digest TEXT NOT NULL,
                date_system TEXT NOT NULL,
                result_json TEXT NOT NULL,
                first_seen TEXT NOT NULL,
                last_seen TEXT NOT NULL,
                UNIQUE(source, digest));
            CREATE TABLE IF NOT EXISTS evidence (
                batch_id TEXT NOT NULL REFERENCES batches(batch_id),
                enterprise_id TEXT NOT NULL,
                status TEXT NOT NULL,
                record_json TEXT NOT NULL,
                PRIMARY KEY(batch_id, enterprise_id));
            CREATE TABLE IF NOT EXISTS batch_observations (
                source TEXT NOT NULL,
                digest TEXT NOT NULL,
                observed_at TEXT NOT NULL,
                PRIMARY KEY(source, digest, observed_at));
        """)

    @staticmethod
    def build_batch_id(source: str, digest: str) -> str:
        return hashlib.sha256(f'{source}|{digest}'.encode()).hexdigest()[:16]

    def save_batch(self, source: str, result: dict, observed_at: str | None = None) -> dict:
        seen = observed_at or datetime.now(timezone.utc).isoformat()
        digest = result['原文摘要']
        if not digest:
            raise ValueError('采集结果缺少原文摘要')
        batch_id = self.build_batch_id(source, digest)
        encoded = canonical_json(result)
        with self.connection:
            existing = self.connection.execute(
                'SELECT batch_id, first_seen FROM batches WHERE batch_id=?', (batch_id,)).fetchone()
            if existing is None:
                self.connection.execute(
                    'INSERT INTO batches VALUES(?,?,?,?,?,?,?)',
                    (batch_id, source, digest, result['日期制'], encoded, seen, seen))
                for record in result['企业申报']:
                    self.connection.execute(
                        'INSERT INTO evidence VALUES(?,?,?,?)',
                        (batch_id, record['企业编号'], record['状态'], canonical_json(record)))
                action = '新建批次'
            else:
                self.connection.execute(
                    'UPDATE batches SET last_seen=? WHERE batch_id=?', (seen, batch_id))
                action = '复用批次'
            self.connection.execute(
                'INSERT OR IGNORE INTO batch_observations VALUES(?,?,?)', (source, digest, seen))
        return {'批次号': batch_id, '处理': action, '原文摘要': digest, '观察时间': seen}

    def list_batches(self) -> list[dict]:
        rows = self.connection.execute("""
            SELECT b.batch_id, b.source, b.digest, b.date_system,
                   b.first_seen, b.last_seen,
                   COUNT(e.enterprise_id),
                   SUM(CASE WHEN e.status='可信' THEN 1 ELSE 0 END),
                   SUM(CASE WHEN e.status='拒收' THEN 1 ELSE 0 END)
            FROM batches b LEFT JOIN evidence e ON e.batch_id=b.batch_id
            GROUP BY b.batch_id ORDER BY b.first_seen, b.source
        """).fetchall()
        return [
            {'批次号': row[0], '来源': row[1], '原文摘要': row[2], '日期制': row[3],
             '首次导入': row[4], '最近导入': row[5], '企业数': row[6],
             '可信数': row[7] or 0, '拒收数': row[8] or 0}
            for row in rows
        ]

    def get_batch(self, batch_id: str) -> dict | None:
        row = self.connection.execute(
            'SELECT result_json FROM batches WHERE batch_id=?', (batch_id,)).fetchone()
        if row is None:
            return None
        return json.loads(row[0])

    def evidence_for(self, enterprise_id: str) -> list[dict]:
        rows = self.connection.execute("""
            SELECT b.batch_id, b.source, b.first_seen, e.status, e.record_json
            FROM evidence e JOIN batches b ON b.batch_id=e.batch_id
            WHERE e.enterprise_id=? ORDER BY b.first_seen
        """, (enterprise_id,)).fetchall()
        return [
            {'批次号': row[0], '来源': row[1], '导入时间': row[2],
             '状态': row[3], '记录': json.loads(row[4])}
            for row in rows
        ]

    def close(self):
        self.connection.close()
