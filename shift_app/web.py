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
                      'api_masters', 'export_xlsx', 'next_month',
                      'csv_upload', 'csv_delete', 'api_csv_list', 'api_rodo',
                      'api_night', 'api_timee', 'api_aliases',
                      'api_snapshot', 'api_snapshot_diff'}
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

    # ---- 労務チェック(勤怠CSV照合・タイミー・夜勤体制) ----
    from . import rodo

    def _aliases():
        return rodo.parse_aliases(store.get_setting('name_aliases'))

    @app.post('/f/<int:fid>/csv')
    def csv_upload(fid):
        kind = request.form.get('kind', 'kintai')
        for f in request.files.getlist('file'):
            if f and f.filename.lower().endswith('.csv'):
                store.add_csv(fid, kind, f.filename, f.read())
        return redirect(url_for('facility', fid=fid) + '#rodo')

    @app.post('/f/<int:fid>/csv/<int:csv_id>/delete')
    def csv_delete(fid, csv_id):
        store.delete_csv(csv_id)
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/csv')
    def api_csv_list(fid):
        return jsonify({'files': store.csv_list(fid)})

    @app.get('/api/f/<int:fid>/rodo')
    def api_rodo(fid):
        tol = int(request.args.get('tol', 15))
        hour_tol = float(request.args.get('hourtol', 0.25))
        model = manager.get(fid)
        aliases = _aliases()
        maps = model.schedule_maps(aliases)
        files = store.csv_list(fid, 'kintai')
        if not files:
            return jsonify({'error': '勤怠CSVが未アップロードです', 'daily': [], 'monthly': []})
        merged = {'by_day': {}, 'by_name': {}, 'names': {}}
        for f in files:
            k = rodo.parse_kintai(store.csv_content(f['id']), aliases, maps['excluded'])
            for key, byd in k['by_day'].items():
                for d, lst in byd.items():
                    merged['by_day'].setdefault(key, {}).setdefault(d, []).extend(lst)
            for key, lst in k['by_name'].items():
                merged['by_name'].setdefault(key, []).extend(lst)
            merged['names'].update(k['names'])
        daily = rodo.compare_daily(maps['schedule'], merged, tol)
        planned, meta = model.planned_hours(aliases)
        monthly = rodo.compare_monthly(planned, meta, merged, hour_tol)
        return jsonify({'daily': daily, 'monthly': monthly,
                        'csv_names': sorted(merged['names'].values()),
                        'files': [f['filename'] for f in files]})

    @app.get('/api/f/<int:fid>/night')
    def api_night(fid):
        model = manager.get(fid)
        maps = model.schedule_maps(_aliases())
        ncodes = rodo.night_codes_from_master(maps['code_times'])
        cal = rodo.night_calendar(maps['by_day'], ncodes)
        y, m = maps['year'], maps['month']
        import calendar as _cal
        days = []
        counts = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        for d in range(1, model.days_in_month() + 1):
            key = f'{y}/{m}/{d}'
            entries = cal.get(key, [])
            n = len(entries)
            counts[min(n, 4)] += 1
            days.append({'day': d,
                         'weekday': ['月', '火', '水', '木', '金', '土', '日'][_cal.weekday(y, m, d)],
                         'count': n,
                         'staff': [{'name': e['name'], 'code': e['code']} for e in entries]})
        return jsonify({'days': days, 'counts': counts,
                        'night_codes': sorted(ncodes.keys()), 'year': y, 'month': m})

    @app.get('/api/f/<int:fid>/timee')
    def api_timee(fid):
        model = manager.get(fid)
        maps = model.schedule_maps(_aliases())
        files = store.csv_list(fid, 'timee')
        details = []
        for f in files:
            details.extend(rodo.parse_timee(store.csv_content(f['id'])))
        summary = rodo.timee_summary(details)
        # シフト上のタイミー枠と明細件数の日別照合
        by_date = {}
        for x in details:
            by_date[x['求人日']] = by_date.get(x['求人日'], 0) + 1
        recon = []
        for d in sorted(set(maps['timee_slots']) | set(by_date), key=rodo._dord):
            slots = maps['timee_slots'].get(d, 0)
            actual = by_date.get(d, 0)
            if slots != actual:
                recon.append({'日付': d, 'シフト上の枠': slots, '明細件数': actual,
                              '判定': '要確認',
                              '内容': '明細なし' if actual == 0 else 'シフト枠なし' if slots == 0 else '件数不一致'})
        return jsonify({'summary': summary, 'recon': recon,
                        'files': [f['filename'] for f in files]})

    @app.route('/api/aliases', methods=['GET', 'POST'])
    def api_aliases():
        if request.method == 'POST':
            store.set_setting('name_aliases', request.get_json(force=True).get('text', ''))
            return jsonify({'ok': True})
        return jsonify({'text': store.get_setting('name_aliases')})

    # ---- スナップショット(週次チェック) ----
    @app.post('/api/f/<int:fid>/snapshot')
    def api_snapshot(fid):
        model = manager.get(fid)
        grid = model.grid()
        data = {str(r['row']): {'name': r['name'], 'cells': r['cells']} for r in grid['rows']}
        store.add_snapshot(fid, data, username=session.get('username'))
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/snapshot/diff')
    def api_snapshot_diff(fid):
        snap = store.latest_snapshot(fid)
        if snap is None:
            return jsonify({'error': 'スナップショットがありません'})
        model = manager.get(fid)
        grid = model.grid()
        cur = {str(r['row']): {'name': r['name'], 'cells': r['cells']} for r in grid['rows']}
        changes = []
        for rk in sorted(set(snap['data']) | set(cur), key=int):
            before = snap['data'].get(rk)
            after = cur.get(rk)
            if before is None or after is None:
                changes.append({'職員': (after or before)['name'], '日': '',
                                '前': '' if before is None else '(行あり)',
                                '後': '' if after is None else '(行あり)',
                                '種別': '職員追加' if before is None else '職員削除'})
                continue
            if str(before['name']) != str(after['name']):
                changes.append({'職員': f"{before['name']} → {after['name']}", '日': '',
                                '前': before['name'], '後': after['name'], '種別': '氏名変更'})
            for i in range(max(len(before['cells']), len(after['cells']))):
                b = str(before['cells'][i]) if i < len(before['cells']) else ''
                a = str(after['cells'][i]) if i < len(after['cells']) else ''
                if b != a:
                    changes.append({'職員': after['name'] or before['name'], '日': i + 1,
                                    '前': b or '(空欄)', '後': a or '(空欄)', '種別': '勤務変更'})
        import datetime
        return jsonify({'taken_at': datetime.datetime.fromtimestamp(snap['ts']).strftime('%Y/%m/%d %H:%M'),
                        'taken_by': snap['username'], 'changes': changes})

    @app.get('/api/f/<int:fid>/masters')
    def api_masters(fid):
        from dataclasses import asdict
        f = store.facility(fid)
        m = extract_masters(store.upload_path(f['filename']))
        return jsonify(asdict(m))

    return app


if __name__ == '__main__':
    create_app().run(debug=False, port=int(os.environ.get('PORT', 5000)))
