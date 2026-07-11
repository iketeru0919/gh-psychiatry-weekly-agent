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
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    pw_hash TEXT NOT NULL,      -- salt$hash (pbkdf2-sha256)
    role TEXT NOT NULL,         -- admin / area / facility / staff
    name TEXT,
    created REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS user_facilities (
    user_id INTEGER NOT NULL,
    facility_id INTEGER NOT NULL,
    PRIMARY KEY (user_id, facility_id)
);
CREATE TABLE IF NOT EXISTS summaries (
    facility_id INTEGER PRIMARY KEY,
    data TEXT NOT NULL,         -- JSON
    computed_at REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS csv_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id INTEGER NOT NULL,
    kind TEXT NOT NULL,          -- kintai / timee
    filename TEXT NOT NULL,
    content BLOB NOT NULL,
    uploaded REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS snapshots (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    facility_id INTEGER NOT NULL,
    ts REAL NOT NULL,
    username TEXT,
    data TEXT NOT NULL           -- グリッドのJSON
);
CREATE TABLE IF NOT EXISTS edit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts REAL NOT NULL,
    username TEXT,
    facility_id INTEGER NOT NULL,
    sheet TEXT NOT NULL,
    col INTEGER NOT NULL,
    row INTEGER NOT NULL,
    value TEXT
);
"""


def _hash_pw(password, salt=None):
    import hashlib
    import os as _os
    salt = salt or _os.urandom(16).hex()
    h = hashlib.pbkdf2_hmac('sha256', password.encode(), bytes.fromhex(salt), 200_000).hex()
    return f'{salt}${h}'


def _verify_pw(password, stored):
    salt = stored.split('$', 1)[0]
    return _hash_pw(password, salt) == stored


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
    def set_override(self, fid, sheet, col, row, value, username=None):
        now = time.time()
        with self._db() as db:
            db.execute(
                'INSERT INTO overrides (facility_id, sheet, col, row, value, updated) '
                'VALUES (?,?,?,?,?,?) '
                'ON CONFLICT(facility_id, sheet, col, row) DO UPDATE SET value=excluded.value, updated=excluded.updated',
                (fid, sheet, col, row, json.dumps(value, ensure_ascii=False), now))
            db.execute(
                'INSERT INTO edit_log (ts, username, facility_id, sheet, col, row, value) VALUES (?,?,?,?,?,?,?)',
                (now, username, fid, sheet, col, row, json.dumps(value, ensure_ascii=False)))
            db.execute('DELETE FROM summaries WHERE facility_id=?', (fid,))

    def overrides(self, fid):
        with self._db() as db:
            return [(r['sheet'], r['col'], r['row'], json.loads(r['value']))
                    for r in db.execute('SELECT * FROM overrides WHERE facility_id=?', (fid,))]

    # --- users ---
    def create_user(self, username, password, role, name=None, facility_ids=()):
        with self._db() as db:
            cur = db.execute(
                'INSERT INTO users (username, pw_hash, role, name, created) VALUES (?,?,?,?,?)',
                (username, _hash_pw(password), role, name, time.time()))
            uid = cur.lastrowid
            for fid in facility_ids:
                db.execute('INSERT OR IGNORE INTO user_facilities VALUES (?,?)', (uid, fid))
            return uid

    def verify_user(self, username, password):
        with self._db() as db:
            r = db.execute('SELECT * FROM users WHERE username=?', (username,)).fetchone()
            if r and _verify_pw(password, r['pw_hash']):
                return dict(r)
            return None

    def users(self):
        with self._db() as db:
            out = []
            for r in db.execute('SELECT * FROM users ORDER BY role, username'):
                u = dict(r)
                u['facility_ids'] = [x['facility_id'] for x in db.execute(
                    'SELECT facility_id FROM user_facilities WHERE user_id=?', (r['id'],))]
                out.append(u)
            return out

    def user_count(self):
        with self._db() as db:
            return db.execute('SELECT COUNT(*) c FROM users').fetchone()['c']

    def delete_user(self, uid):
        with self._db() as db:
            db.execute('DELETE FROM user_facilities WHERE user_id=?', (uid,))
            db.execute('DELETE FROM users WHERE id=?', (uid,))

    def user_facility_ids(self, uid):
        with self._db() as db:
            return {r['facility_id'] for r in db.execute(
                'SELECT facility_id FROM user_facilities WHERE user_id=?', (uid,))}

    def set_user_facilities(self, uid, facility_ids):
        with self._db() as db:
            db.execute('DELETE FROM user_facilities WHERE user_id=?', (uid,))
            for fid in facility_ids:
                db.execute('INSERT OR IGNORE INTO user_facilities VALUES (?,?)', (uid, fid))

    # --- CSV(勤怠/タイミー) ---
    def add_csv(self, fid, kind, filename, content):
        with self._db() as db:
            cur = db.execute('INSERT INTO csv_files (facility_id, kind, filename, content, uploaded) VALUES (?,?,?,?,?)',
                             (fid, kind, filename, content, time.time()))
            return cur.lastrowid

    def csv_list(self, fid, kind=None):
        q = 'SELECT id, facility_id, kind, filename, uploaded, LENGTH(content) size FROM csv_files WHERE facility_id=?'
        args = [fid]
        if kind:
            q += ' AND kind=?'
            args.append(kind)
        with self._db() as db:
            return [dict(r) for r in db.execute(q + ' ORDER BY uploaded', args)]

    def csv_content(self, csv_id):
        with self._db() as db:
            r = db.execute('SELECT content FROM csv_files WHERE id=?', (csv_id,)).fetchone()
            return r['content'] if r else None

    def delete_csv(self, csv_id):
        with self._db() as db:
            db.execute('DELETE FROM csv_files WHERE id=?', (csv_id,))

    # --- settings(氏名エイリアス等) ---
    def get_setting(self, key, default=''):
        with self._db() as db:
            r = db.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
            return r['value'] if r else default

    def set_setting(self, key, value):
        with self._db() as db:
            db.execute('INSERT INTO settings (key, value) VALUES (?,?) '
                       'ON CONFLICT(key) DO UPDATE SET value=excluded.value', (key, value))

    # --- snapshots(シフトの版) ---
    def add_snapshot(self, fid, data, username=None):
        with self._db() as db:
            db.execute('INSERT INTO snapshots (facility_id, ts, username, data) VALUES (?,?,?,?)',
                       (fid, time.time(), username, json.dumps(data, ensure_ascii=False)))

    def latest_snapshot(self, fid):
        with self._db() as db:
            r = db.execute('SELECT * FROM snapshots WHERE facility_id=? ORDER BY ts DESC LIMIT 1', (fid,)).fetchone()
            if not r:
                return None
            return {'ts': r['ts'], 'username': r['username'], 'data': json.loads(r['data'])}

    # --- summaries ---
    def save_summary(self, fid, data):
        with self._db() as db:
            db.execute('INSERT INTO summaries (facility_id, data, computed_at) VALUES (?,?,?) '
                       'ON CONFLICT(facility_id) DO UPDATE SET data=excluded.data, computed_at=excluded.computed_at',
                       (fid, json.dumps(data, ensure_ascii=False), time.time()))

    def summaries(self):
        with self._db() as db:
            return {r['facility_id']: {'data': json.loads(r['data']), 'computed_at': r['computed_at']}
                    for r in db.execute('SELECT * FROM summaries')}
