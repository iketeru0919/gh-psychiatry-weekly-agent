"""シフト管理アプリ Web サーバ。

起動:  python -m shift_app.web  →  http://127.0.0.1:5000
データ: ./shift_app_data/ (SQLite + アップロードした Excel)
"""
import os

from flask import Flask, jsonify, redirect, render_template, request, url_for

from .masters import extract_masters
from .model_manager import (INPUT_SHEET, ModelManager, editable_daily_cell,
                            editable_user_cell)
from .store import Store

DATA_ROOT = os.environ.get('SHIFT_APP_DATA', os.path.join(os.getcwd(), 'shift_app_data'))


def create_app(data_root=None):
    app = Flask(__name__)
    store = Store(data_root or DATA_ROOT)
    manager = ModelManager(store)
    app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

    @app.get('/')
    def index():
        return render_template('index.html', facilities=store.facilities())

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
            manager.set_cell(fid, INPUT_SHEET, 5 + day, row, value)
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
            manager.set_cell(fid, sheet, col, row, value)
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
            manager.set_cell(fid, sheet, col, row, value)
        return jsonify({'ok': True})

    @app.get('/api/f/<int:fid>/dashboard')
    def api_dashboard(fid):
        model = manager.get(fid)
        return jsonify(model.dashboard())

    @app.get('/api/f/<int:fid>/masters')
    def api_masters(fid):
        from dataclasses import asdict
        f = store.facility(fid)
        m = extract_masters(store.upload_path(f['filename']))
        return jsonify(asdict(m))

    return app


if __name__ == '__main__':
    create_app().run(debug=False, port=int(os.environ.get('PORT', 5000)))
