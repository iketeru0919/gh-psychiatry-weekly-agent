"""永続化層: 取り込んだブックとセル上書きを SQLite に保存する。"""
import json
import os
import sqlite3
import time

_SCHEMA = """
CREATE TABLE IF NOT EXISTS facilities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    year INTEGER,
    month INTEGER,
    filename TEXT NOT NULL,
    created REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS overrides (
    facility_id INTEGER NOT NULL,
    sheet TEXT NOT NULL,
    col INTEGER NOT NULL,
    row INTEGER NOT NULL,
    value TEXT,             -- JSON エンコード(null=セルを空にする)
    updated REAL NOT NULL,
    PRIMARY KEY (facility_id, sheet, col, row)
);
"""


class Store:
    def __init__(self, root):
        self.root = root
        os.makedirs(root, exist_ok=True)
        os.makedirs(os.path.join(root, 'uploads'), exist_ok=True)
        os.makedirs(os.path.join(root, 'cache'), exist_ok=True)
        self.db_path = os.path.join(root, 'shift.db')
        with self._db() as db:
            db.executescript(_SCHEMA)

    def _db(self):
        db = sqlite3.connect(self.db_path)
        db.row_factory = sqlite3.Row
        return db

    @property
    def cache_dir(self):
        return os.path.join(self.root, 'cache')

    def upload_path(self, filename):
        return os.path.join(self.root, 'uploads', filename)

    # --- facilities ---
    def add_facility(self, name, year, month, file_bytes):
        fname = f'{int(time.time())}-{name[:40].replace("/", "_")}.xlsx'
        with open(self.upload_path(fname), 'wb') as fh:
            fh.write(file_bytes)
        with self._db() as db:
            cur = db.execute(
                'INSERT INTO facilities (name, year, month, filename, created) VALUES (?,?,?,?,?)',
                (name, year, month, fname, time.time()))
            return cur.lastrowid

    def facilities(self):
        with self._db() as db:
            return [dict(r) for r in db.execute(
                'SELECT * FROM facilities ORDER BY name, year, month')]

    def facility(self, fid):
        with self._db() as db:
            r = db.execute('SELECT * FROM facilities WHERE id=?', (fid,)).fetchone()
            return dict(r) if r else None

    def delete_facility(self, fid):
        f = self.facility(fid)
        with self._db() as db:
            db.execute('DELETE FROM overrides WHERE facility_id=?', (fid,))
            db.execute('DELETE FROM facilities WHERE id=?', (fid,))
        if f:
            try:
                os.remove(self.upload_path(f['filename']))
            except OSError:
                pass

    # --- overrides ---
    def set_override(self, fid, sheet, col, row, value):
        with self._db() as db:
            db.execute(
                'INSERT INTO overrides (facility_id, sheet, col, row, value, updated) '
                'VALUES (?,?,?,?,?,?) '
                'ON CONFLICT(facility_id, sheet, col, row) DO UPDATE SET value=excluded.value, updated=excluded.updated',
                (fid, sheet, col, row, json.dumps(value, ensure_ascii=False), time.time()))

    def overrides(self, fid):
        with self._db() as db:
            return [(r['sheet'], r['col'], r['row'], json.loads(r['value']))
                    for r in db.execute('SELECT * FROM overrides WHERE facility_id=?', (fid,))]
