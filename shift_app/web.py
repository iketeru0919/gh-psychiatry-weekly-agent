"""シフト管理アプリ Web サーバ。

起動:  python -m shift_app.web  →  http://127.0.0.1:5000
データ: ./shift_app_data/ (SQLite + アップロードした Excel)
"""
import io
import os

from flask import (Flask, jsonify, redirect, render_template, request,
                   send_file, session, url_for)

from .export import export_bytes, next_month_bytes
from .masters import extract_masters
from .model_manager import (INPUT_SHEET, ModelManager, editable_daily_cell,
                            editable_user_cell)
from .store import Store

DATA_ROOT = os.environ.get('SHIFT_APP_DATA', os.path.join(os.getcwd(), 'shift_app_data'))

# ロール別に許可するエンドポイント(admin は全部)
_STAFF = {'index', 'facility', 'api_grid', 'api_grid_set',
          'api_daily', 'api_daily_set', 'login', 'logout', 'static'}
_FACILITY = _STAFF | {'api_dashboard', 'api_users', 'api_users_set',
                      'api_masters', 'export_xlsx', 'next_month'}
_AREA = _FACILITY | {'overview', 'api_overview'}
ROLE_ENDPOINTS = {'staff': _STAFF, 'facility': _FACILITY, 'area': _AREA}
ROLE_LABELS = {'admin': '本部管理者', 'area': 'エリアマネージャー',
               'facility': '施設管理者', 'staff': '現場'}


def create_app(data_root=None):
    app = Flask(__name__)
    store = Store(data_root or DATA_ROOT)
    manager = ModelManager(store)
    app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

    # セッション鍵はデータディレクトリに永続化
    secret_path = os.path.join(store.root, 'secret_key')
    if not os.path.exists(secret_path):
        with open(secret_path, 'wb') as fh:
            fh.write(os.urandom(32))
    app.secret_key = open(secret_path, 'rb').read()

    # 初回起動時: ユーザーが1人もいなければ環境変数から admin を作成
    if store.user_count() == 0 and os.environ.get('SHIFT_APP_ADMIN_PW'):
        store.create_user('admin', os.environ['SHIFT_APP_ADMIN_PW'], 'admin', name='本部管理者')

    def auth_enabled():
        return store.user_count() > 0

    def current_role():
        return session.get('role')

    def allowed_fids():
        """このユーザーが触れる施設ID集合(None=全施設)。"""
        if not auth_enabled() or current_role() == 'admin':
            return None
        return store.user_facility_ids(session.get('uid', -1))

    def fid_allowed(fid):
        fids = allowed_fids()
        return fids is None or fid in fids

    @app.before_request
    def _guard():
        if not auth_enabled() or request.endpoint in (None, 'login', 'logout', 'static'):
            return None
        role = current_role()
        if role is None:
            if request.path.startswith('/api/'):
                return jsonify({'ok': False, 'error': 'ログインが必要です'}), 401
            return redirect(url_for('login', next=request.path))
        if role != 'admin' and request.endpoint not in ROLE_ENDPOINTS.get(role, set()):
            return jsonify({'ok': False, 'error': 'この操作の権限がありません'}), 403
        fid = (request.view_args or {}).get('fid')
        if fid is not None and not fid_allowed(fid):
            return jsonify({'ok': False, 'error': 'この施設へのアクセス権がありません'}), 403
        return None

    @app.route('/login', methods=['GET', 'POST'])
    def login():
        error = None
        if request.method == 'POST':
            u = store.verify_user(request.form.get('username', ''), request.form.get('password', ''))
            if u:
                session['uid'] = u['id']
                session['role'] = u['role']
                session['username'] = u['username']
                return redirect(request.args.get('next') or
                                (url_for('overview') if u['role'] in ('admin', 'area') else url_for('index')))
            error = 'ユーザー名またはパスワードが違います'
        return render_template('login.html', error=error)

    @app.get('/logout')
    def logout():
        session.clear()
        return redirect(url_for('login') if auth_enabled() else url_for('index'))

    @app.get('/')
    def index():
        facs = store.facilities()
        fids = allowed_fids()
        if fids is not None:
            facs = [f for f in facs if f['id'] in fids]
        return render_template('index.html', facilities=facs,
                               role=current_role() or 'admin',
                               username=session.get('username'))

    # ---- 横断ダッシュボード ----
    @app.get('/overview')
    def overview():
        return render_template('overview.html', username=session.get('username'),
                               role=current_role() or 'admin')

    @app.get('/api/overview')
    def api_overview():
        facs = store.facilities()
        fids = allowed_fids()
        if fids is not None:
            facs = [f for f in facs if f['id'] in fids]
        cached = store.summaries()
        force = request.args.get('recompute') == '1'
        out = []
        for f in facs:
            if force or f['id'] not in cached:
                data, ts = manager.summary(f['id'], force=force)
            else:
                data, ts = cached[f['id']]['data'], cached[f['id']]['computed_at']
            out.append({'id': f['id'], 'name': f['name'], 'summary': data, 'computed_at': ts})
        return jsonify({'facilities': out})

    # ---- ユーザー管理(admin) ----
    @app.route('/admin/users', methods=['GET', 'POST'])
    def admin_users():
        error = None
        if request.method == 'POST':
            act = request.form.get('action')
            if act == 'create':
                try:
                    fids = [int(x) for x in request.form.getlist('facilities')]
                    store.create_user(request.form['username'].strip(),
                                      request.form['password'],
                                      request.form['role'],
                                      name=request.form.get('name') or None,
                                      facility_ids=fids)
                except Exception as e:
                    error = f'作成できません: {e}'
            elif act == 'delete':
                uid = int(request.form['uid'])
                if uid != session.get('uid'):
                    store.delete_user(uid)
                else:
                    error = '自分自身は削除できません'
        return render_template('users_admin.html', users=store.users(),
                               facilities=store.facilities(), error=error,
                               role_labels=ROLE_LABELS, username=session.get('username'))

    @app.post('/import')
    def import_xlsx():
        f = request.files.get('file')
        if not f or not f.filename.endswith('.xlsx'):
            return redirect(url_for('index'))
        data = f.read()
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        try:
            m = extract_masters(tmp_path)
        finally:
            os.unlink(tmp_path)
        fid = store.add_facility(m.facility_name or f.filename, m.year, m.month, data)
        return redirect(url_for('facility', fid=fid))

    @app.post('/f/<int:fid>/delete')
    def delete_facility(fid):
        store.delete_facility(fid)
        manager.drop(fid)
        return redirect(url_for('index'))

    @app.get('/f/<int:fid>')
    def facility(fid):
        f = store.facility(fid)
        if f is None:
            return redirect(url_for('index'))
        return render_template('facility.html', f=f)

    @app.get('/api/f/<int:fid>/grid')
    def api_grid(fid):
        model = manager.get(fid)
        return jsonify(model.grid())

    @app.post('/api/f/<int:fid>/grid')
    def api_grid_set(fid):
        payload = request.get_json(force=True)
        for edit in payload.get('edits', []):
            row = int(edit['row'])
            day = int(edit['day'])
            value = (edit.get('value') or '').strip() or None
            manager.set_cell(fid, INPUT_SHEET, 5 + day, row, value,
                             username=session.get('username'))
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/users')
    def api_users(fid):
        model = manager.get(fid)
        return jsonify(model.users())

    @app.post('/api/f/<int:fid>/users')
    def api_users_set(fid):
        payload = request.get_json(force=True)
        for edit in payload.get('edits', []):
            sheet = edit['sheet']
            col, row = int(edit['col']), int(edit['row'])
            if not editable_user_cell(sheet, col, row):
                return jsonify({'ok': False, 'error': f'編集不可セル {sheet}!{col},{row}'}), 400
            value = (str(edit.get('value') or '')).strip() or None
            if value is not None:
                try:
                    value = float(value)
                except ValueError:
                    pass
            manager.set_cell(fid, sheet, col, row, value,
                             username=session.get('username'))
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/daily/<int:day>')
    def api_daily(fid, day):
        model = manager.get(fid)
        return jsonify(model.daily(day))

    @app.post('/api/f/<int:fid>/daily')
    def api_daily_set(fid):
        payload = request.get_json(force=True)
        for edit in payload.get('edits', []):
            sheet = edit['sheet']
            col, row = int(edit['col']), int(edit['row'])
            if not editable_daily_cell(sheet, col, row):
                return jsonify({'ok': False, 'error': f'編集不可セル {sheet}!{col},{row}'}), 400
            value = (str(edit.get('value') or '')).strip() or None
            manager.set_cell(fid, sheet, col, row, value,
                             username=session.get('username'))
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/dashboard')
    def api_dashboard(fid):
        model = manager.get(fid)
        return jsonify(model.dashboard())

    @app.get('/f/<int:fid>/export.xlsx')
    def export_xlsx(fid):
        f = store.facility(fid)
        data = export_bytes(store.upload_path(f['filename']), store.overrides(fid))
        name = f'{f["name"]}_{f["year"]}年{f["month"]}月.xlsx'
        return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    @app.post('/f/<int:fid>/next-month')
    def next_month(fid):
        f = store.facility(fid)
        data, y, m = next_month_bytes(store.upload_path(f['filename']), store.overrides(fid))
        new_id = store.add_facility(f['name'], y, m, data)
        return redirect(url_for('facility', fid=new_id))

    @app.get('/api/f/<int:fid>/masters')
    def api_masters(fid):
        from dataclasses import asdict
        f = store.facility(fid)
        m = extract_masters(store.upload_path(f['filename']))
        return jsonify(asdict(m))

    return app


if __name__ == '__main__':
    create_app().run(debug=False, port=int(os.environ.get('PORT', 5000)))
