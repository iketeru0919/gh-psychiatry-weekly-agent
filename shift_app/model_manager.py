"""施設ブックの計算モデル管理。

取り込んだ Excel の数式モデル(xlcalc.Workbook)に UI からの入力(上書き)を
適用し、シフトグリッド・管理ダッシュボードのデータを提供する。
計算は常にブック内の数式そのもので行うため、Excel との完全一致が保たれる。
"""
import calendar
import re

from .xlcalc.evaluator import RangeVal, Workbook, XLErr, serial_to_date
from .xlcalc.parser import col_to_num, num_to_col, parse_formula

INPUT_SHEET = '入力シート【現場配布用】'
ADMIN_SHEET = '管理用シート'
MASTER_SHEET = 'マスタ設定'
PLACEMENT_SHEET = '人員配置【入力用】'
MONTHLY_SHEET = '★毎月利用日数入力'

USER_ROWS = range(13, 48)          # 利用者行(両シート共通)
MONTH_BASE_COLS = [5 + 5 * k for k in range(12)]  # E,J,O,... 12か月分の先頭列

GRID_FIRST_ROW = 6
GRID_LAST_ROW = 55
GRID_DAY_COL0 = 6          # F列 = 1日目
SUMMARY_ROWS = range(57, 74)   # 下部の日勤/夜勤合計など
ADMIN_STAFF_ROWS = range(9, 65)

_C = col_to_num


def _disp(v):
    """評価値を表示用に変換。"""
    if v is None:
        return ''
    if isinstance(v, XLErr):
        return v.code
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return int(v)
        return round(v, 3)
    return v


class FacilityModel:
    def __init__(self, path, overrides, cache_dir=None):
        self.wb = Workbook(path, cache_dir=cache_dir)
        for sheet, col, row, value in overrides:
            self._apply(sheet, col, row, value)

    def _apply(self, sheet, col, row, value):
        cells = self.wb.cells.setdefault(sheet, {})
        if value is None or value == '':
            cells.pop((col, row), None)
        else:
            cells[(col, row)] = ('v', float(value) if isinstance(value, (int, float)) else value)
        self.wb.formulas.get(sheet, {}).pop((col, row), None)

    def set_cell(self, sheet, col, row, value):
        self._apply(sheet, col, row, value)
        self.wb._memo.clear()
        self.wb.volatile_cells.clear()

    def v(self, sheet, addr_or_col, row=None):
        if row is None:
            import re
            m = re.match(r'([A-Z]+)(\d+)', addr_or_col)
            col, row = _C(m.group(1)), int(m.group(2))
        else:
            col = addr_or_col
        try:
            return _disp(self.wb.value(sheet, col, row))
        except (RuntimeError, NotImplementedError, ValueError) as e:
            return f'<{e}>'

    # ---- メタ ----
    def year_month(self):
        y = self.v(MASTER_SHEET, 'A2')
        m = self.v(MASTER_SHEET, 'J2')
        try:
            return int(y), int(m)
        except (TypeError, ValueError):
            return None, None

    def days_in_month(self):
        y, m = self.year_month()
        if not y:
            return 31
        return calendar.monthrange(y, m)[1]

    def shift_codes(self):
        """マスタ設定の略記号一覧(グリッド入力の選択肢)。"""
        codes = []
        for r in range(11, 116):
            c = self.wb.value(MASTER_SHEET, 1, r)
            if isinstance(c, str) and c.strip():
                codes.append(c.strip())
        seen = set()
        return [c for c in codes if not (c in seen or seen.add(c))]

    # ---- シフトグリッド ----
    def grid(self):
        days = self.days_in_month()
        y, m = self.year_month()
        cols = []
        for d in range(1, days + 1):
            wd = ['月', '火', '水', '木', '金', '土', '日'][calendar.weekday(y, m, d)] if y else ''
            cols.append({'day': d, 'weekday': wd})
        rows = []
        for r in range(GRID_FIRST_ROW, GRID_LAST_ROW + 1):
            name = self.v(INPUT_SHEET, 'C%d' % r)
            role = self.v(INPUT_SHEET, 'B%d' % r)
            if not name and not role:
                continue
            cells = []
            for d in range(days):
                col = GRID_DAY_COL0 + d
                cell = self.wb.cells.get(INPUT_SHEET, {}).get((col, r))
                raw = cell[1] if cell and cell[0] == 'v' else (self.v(INPUT_SHEET, col, r) if cell else '')
                cells.append(_disp(raw) if not isinstance(raw, str) else raw)
            rows.append({'row': r, 'name': name, 'role': role,
                         'qual': self.v(INPUT_SHEET, 'D%d' % r), 'cells': cells})
        summary = []
        for r in SUMMARY_ROWS:
            label = self.v(INPUT_SHEET, 'C%d' % r)
            if not label:
                continue
            vals = [self.v(INPUT_SHEET, GRID_DAY_COL0 + d, r) for d in range(days)]
            summary.append({'label': label, 'cells': vals})
        return {'days': cols, 'rows': rows, 'summary': summary,
                'codes': self.shift_codes(), 'year': y, 'month': m}

    # ---- 管理ダッシュボード ----
    def dashboard(self):
        A = ADMIN_SHEET
        staff = []
        for r in ADMIN_STAFF_ROWS:
            name = self.v(A, 'M%d' % r)
            if not name:
                continue
            staff.append({
                'row': r,
                'name': name,
                'role': self.v(A, 'B%d' % r),
                'qual': self.v(A, 'V%d' % r),
                'pay': self.v(A, 'CT%d' % r),
                'total_hours': self.v(A, 'CU%d' % r),      # 総労働勤務時間
                'day_hours': self.v(A, 'DE%d' % r),        # 日中勤務時間数
                'fte': self.v(A, 'DL%d' % r),              # 日中配置換算(人工)
                'holidays': self.v(A, 'ES%d' % r),         # 休日数
                'legal_holiday': self.v(A, 'EW%d' % r),    # 法定休日判定
            })
        kpi = {
            '月の常勤基準時間': self.v(A, 'DL3'),
            '生活支援員 合計時間': self.v(A, 'CU25'),
            '生活支援員 日中時間': self.v(A, 'DE25'),
            '生活支援員 人工(換算)': self.v(A, 'DL25'),
            '世話人 合計時間': self.v(A, 'CU65'),
            '世話人 日中時間': self.v(A, 'DE65'),
            '世話人 人工(換算)': self.v(A, 'DL65'),
            '割増金発生 件数': self.v(A, 'EW65'),
            '月給制 合計時間': self.v(A, 'EF94'),
            '時給制 合計時間': self.v(A, 'EP94'),
        }
        haichi = {
            '現在のラシエル基準値': self.v(A, 'BU73'),
            '採用基準(法/ラ)': self.v(A, 'CE73'),
            '基準判定': self.v(A, 'BB75'),
            '人員超過・過不足数': self.v(A, 'BU75'),
            '③法令基準値(前年度実績)': self.v(A, 'DU75'),
            '④現利用者 生活支援数': self.v(A, 'CS73'),
            '④現利用者 世話人数': self.v(A, 'CS74'),
            '④判定': self.v(A, 'EV78'),
            '④過不足': self.v(A, 'EV81'),
        }
        inputs = {}
        for label, (col, row) in ADMIN_INPUT_CELLS.items():
            inputs[label] = {'sheet': A, 'col': col, 'row': row, 'value': self._raw(A, col, row)}
        return {'staff': staff, 'kpi': kpi, 'haichi': haichi, 'inputs': inputs}

    # ---- 利用者稼働実態 ----
    def _raw(self, sheet, col, row):
        cell = self.wb.cells.get(sheet, {}).get((col, row))
        if cell and cell[0] == 'v':
            return _disp(cell[1])
        return ''

    def users(self):
        P, M = PLACEMENT_SHEET, MONTHLY_SHEET
        placement = {
            'month_days': self._raw(P, 5, 5),
            'rows': [],
            'need': {
                '生活支援必要数': self.v(P, 'E7'),
                '区分3': self.v(P, 'F7'), '区分4': self.v(P, 'G7'),
                '区分5': self.v(P, 'H7'), '区分6': self.v(P, 'I7'),
            },
            'totals': {
                '区分3': self.v(P, 'F48'), '区分4': self.v(P, 'G48'),
                '区分5': self.v(P, 'H48'), '区分6': self.v(P, 'I48'),
                '延べ利用回数': self.v(P, 'F49'),
            },
        }
        for r in USER_ROWS:
            name = self._raw(P, 3, r)
            kubun = self._raw(P, 4, r)
            days = self._raw(P, 5, r)
            placement['rows'].append({'row': r, 'name': name, 'kubun': kubun, 'days': days})

        months = []
        for i, base in enumerate(MONTH_BASE_COLS):
            months.append({
                'index': i + 1,
                'col': base,
                'label': self.v(M, base, 4) or f'{i + 1}か月目',
                'month_days': self._raw(M, base, 5),
                'need': self.v(M, base, 7),
            })
        monthly_rows = []
        for r in USER_ROWS:
            name = self._raw(M, 3, r)
            kubun = self._raw(M, 4, r)
            cells = [self._raw(M, base, r) for base in MONTH_BASE_COLS]
            monthly_rows.append({'row': r, 'name': name, 'kubun': kubun, 'cells': cells})
        return {'placement': placement, 'months': months, 'monthly_rows': monthly_rows,
                'sheets': {'placement': P, 'monthly': M}}

    # ---- 日次(勤怠実績+チェックリスト) ----
    def daily_sheet_names(self):
        """日番号 → (勤怠シート名, 日次チェックシート名)。全角数字も正規化する。"""
        kintai, checklist = {}, {}
        for name in self.wb.cells:
            n = name.translate(_FW)
            m = re.fullmatch(r'\((\d+)\)', n.strip())
            if m:
                kintai[int(m.group(1))] = name
                continue
            m = re.fullmatch(r'日次\s*\((\d+)\)', n.strip())
            if m:
                checklist[int(m.group(1))] = name
        return kintai, checklist

    @staticmethod
    def _fmt_serial_time(v):
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            mins = round((v % 1 if v >= 1 else v) * 24 * 60)
            if v >= 1:  # 日またぎ(翌日)
                mins = round(v * 24 * 60) if v < 2 else mins
            return f'{mins // 60}:{mins % 60:02d}'
        return _disp(v)

    def daily(self, day):
        kintai_map, check_map = self.daily_sheet_names()
        K = kintai_map.get(day)
        C = check_map.get(day)
        out = {'day': day, 'kintai_sheet': K, 'checklist_sheet': C,
               'days_in_month': self.days_in_month()}
        if K:
            date_serial = self.wb.value(K, 10, 1)  # J1
            if isinstance(date_serial, (int, float)):
                d = serial_to_date(date_serial)
                out['date_label'] = f'{d.year}年{d.month}月{d.day}日 {self.v(K, "N1")}'
            # 出勤者リスト(W5 の UNIQUE/FILTER をライブ評価)
            workers = []
            w5 = self.wb.formulas.get(K, {}).get((23, 5))
            if w5:
                from .xlcalc.evaluator import _Ctx
                try:
                    res = self.wb.eval_node(parse_formula(w5), _Ctx(K, 23, 5))
                    if isinstance(res, RangeVal):
                        workers = [r[0] for r in res.rows if r[0]]
                except Exception:
                    workers = []
            out['workers'] = workers
            rows = []
            for i, r in enumerate(KINTAI_ROWS):
                worker = workers[i] if i < len(workers) else ''
                rows.append({
                    'row': r, 'no': i + 1, 'worker': worker,
                    'plan_start': self._fmt_serial_time(self.wb.value(K, 3, r)) if worker else '',
                    'plan_end': self._fmt_serial_time(self.wb.value(K, 5, r)) if worker else '',
                    'actual_start': self._raw(K, 6, r), 'actual_end': self._raw(K, 8, r),
                    'break_start': self._raw(K, 9, r), 'break_end': self._raw(K, 11, r),
                    'no_change': self._raw(K, 13, r), 'changed': self._raw(K, 14, r),
                    'sign': self._raw(K, 15, r),
                })
            out['kintai'] = rows
            out['note'] = self._raw(K, *KINTAI_NOTE_CELL)
        if C:
            items = []
            for r in CHECKLIST_ROWS:
                label = self._raw(C, 2, r) or self.v(C, 2, r)
                cells = [self._raw(C, col, r) for col in CHECKLIST_COLS]
                items.append({'row': r, 'label': label, 'cells': cells})
            out['checklist'] = {
                'headers': [self._raw(C, col, 4) or self.v(C, col, 4) for col in CHECKLIST_COLS],
                'items': items,
            }
        return out


    # ---- 労務チェック用のシフト側データ ----
    def code_times(self):
        """略記号 → (開始分, 終了分, 日またぎか)。マスタ設定の時刻欄から。"""
        from .rodo import norm_shift
        out = {}
        for r in range(11, 116):
            code = self._raw(MASTER_SHEET, 1, r)
            if not isinstance(code, str) or not code.strip():
                continue
            start = self._raw(MASTER_SHEET, 5, r)    # E列
            end = self._raw(MASTER_SHEET, 11, r)     # K列
            sm = round(start * 1440) if isinstance(start, (int, float)) else None
            overnight = False
            em = None
            if isinstance(end, (int, float)):
                if end >= 1:
                    overnight = True
                    em = round((end - 1) * 1440) or 1440
                else:
                    em = round(end * 1440)
                    if sm is not None and em < sm:
                        overnight = True
            key = norm_shift(code)
            if key and key not in out:
                out[key] = (sm, em, overnight)
        return out

    def _lookup_times(self, code, ctimes):
        """照合用: 記号から(開始分,終了分)。前方一致のフォールバック付き。"""
        from .rodo import norm_shift
        s = norm_shift(code)
        hit = ctimes.get(s)
        if hit is None:
            for k, v in ctimes.items():
                if k.startswith(s) or s.startswith(k):
                    hit = v
                    break
        if hit is None:
            return None
        sm, em, overnight = hit
        if sm is None and em is None:
            return None
        if overnight and (em is None or em == 0):
            em = 1440
        return (sm, em)

    def schedule_maps(self, aliases=None):
        """グリッドから照合・夜勤チェック用の構造を作る。"""
        from .rodo import is_dispatch, norm_name
        grid = self.grid()
        ctimes = self.code_times()
        y, m = grid['year'], grid['month']
        schedule, by_day, timee_slots, excluded = {}, {}, {}, set()
        for row in grid['rows']:
            name = str(row['name'] or '').strip()
            if not name:
                continue
            key = norm_name(name, aliases)
            if is_dispatch(name, row.get('role')):
                excluded.add(key)
                continue
            is_timee = 'タイミー' in name
            for i, code in enumerate(row['cells']):
                code = str(code or '').strip()
                if not code or code == '0':
                    continue
                d = f'{y}/{m}/{i + 1}'
                if is_timee:
                    timee_slots[d] = timee_slots.get(d, 0) + 1
                    continue
                times = self._lookup_times(code, ctimes)
                schedule.setdefault(key, {}).setdefault(d, []).append(
                    {'code': code, 'name': name, 'times': times})
                by_day.setdefault(d, []).append({'name': name, 'code': code})
        return {'schedule': schedule, 'by_day': by_day, 'timee_slots': timee_slots,
                'excluded': excluded, 'year': y, 'month': m, 'code_times': ctimes}

    def planned_hours(self, aliases=None):
        """職員別の月間予定時間(管理用シートCU列)。重複ブロック(15-16行)は除外。"""
        from .rodo import is_dispatch, norm_name
        planned, meta = {}, {}
        for r in ADMIN_STAFF_ROWS:
            if r in (15, 16):
                continue
            name = self.v(ADMIN_SHEET, 'M%d' % r)
            if not name:
                continue
            role = self.v(ADMIN_SHEET, 'B%d' % r)
            if is_dispatch(name, role) or 'タイミー' in str(name):
                continue
            key = norm_name(name, aliases)
            cu = self.v(ADMIN_SHEET, 'CU%d' % r)
            if isinstance(cu, (int, float)):
                planned[key] = planned.get(key, 0) + cu
            meta.setdefault(key, {'name': name, 'role': role,
                                  'pay': self.v(ADMIN_SHEET, 'CT%d' % r)})
        return planned, meta


# 日次シート: 勤怠実績(N) の編集可能列と、日次チェックリストの範囲
KINTAI_ROWS = range(5, 20)
KINTAI_EDIT_COLS = {6, 8, 9, 11, 13, 14, 15}   # F,H,I,K(実績/休憩) M,N(変更) O(署名)
KINTAI_NOTE_CELL = (2, 22)                      # 特記事項欄
CHECKLIST_ROWS = range(5, 24)
CHECKLIST_COLS = range(3, 18)                   # 出勤者①〜⑮

_FW = str.maketrans('０１２３４５６７８９', '0123456789')


# 管理用シート上の手入力セル(④区分別利用者数・他施設兼務)
ADMIN_INPUT_CELLS = {
    '区分1': (_C('CK'), 70), '区分2': (_C('CO'), 70), '区分3': (_C('CS'), 70),
    '区分4': (_C('CW'), 70), '区分5': (_C('DA'), 70), '区分6': (_C('DE'), 70),
    'サビ管 他施設兼務': (_C('DO'), 71), '管理者 他施設兼務': (_C('DO'), 75),
}


def editable_daily_cell(sheet, col, row):
    """日次タブから編集を許可するセルか判定する。シート名は日別シートに限る。"""
    n = sheet.translate(_FW).strip()
    if re.fullmatch(r'\(\d+\)', n):
        if row in KINTAI_ROWS and col in KINTAI_EDIT_COLS:
            return True
        return (col, row) == KINTAI_NOTE_CELL
    if re.fullmatch(r'日次\s*\(\d+\)', n):
        return row in CHECKLIST_ROWS and col in CHECKLIST_COLS
    return False


def editable_user_cell(sheet, col, row):
    """利用者稼働タブから編集を許可するセルか判定する。"""
    if sheet == ADMIN_SHEET:
        return (col, row) in ADMIN_INPUT_CELLS.values()
    if sheet == PLACEMENT_SHEET:
        if row in USER_ROWS and col in (3, 4, 5):        # 氏名/区分/延べ日数
            return True
        return (col, row) == (5, 5)                       # 月の日数
    if sheet == MONTHLY_SHEET:
        if row in USER_ROWS and (col in (3, 4) or col in MONTH_BASE_COLS):
            return True
        return row == 5 and col in MONTH_BASE_COLS        # 各月の日数
    return False


def _facility_summary(model):
    """横断ダッシュボード用のサマリ(1施設分)。"""
    import datetime
    d = model.dashboard()
    y, m = model.year_month()
    today = datetime.date.today()
    days = model.days_in_month()
    # 入力状況
    grid = model.grid()
    named_rows = [r for r in grid['rows'] if r['name']]
    filled = sum(1 for r in named_rows for c in r['cells'] if c)
    # 経過日数(対象月が当月なら今日まで、過去月なら全日、未来月なら0)
    if y and (y, m) < (today.year, today.month):
        elapsed = days
    elif y and (y, m) == (today.year, today.month):
        elapsed = today.day
    else:
        elapsed = 0
    kintai_map, check_map = model.daily_sheet_names()
    kintai_missing = check_missing = 0
    for day in range(1, elapsed + 1):
        K, C = kintai_map.get(day), check_map.get(day)
        if K and not any(model._raw(K, 15, r) for r in KINTAI_ROWS):
            kintai_missing += 1
        if C and not any(model._raw(C, col, 5) for col in CHECKLIST_COLS):
            check_missing += 1
    legal_ng = sum(1 for s in d['staff'] if '割増' in str(s['legal_holiday']))
    return {
        'year': y, 'month': m,
        '基準判定': d['haichi']['基準判定'],
        '過不足': d['haichi']['人員超過・過不足数'],
        'ラシエル基準値': d['haichi']['現在のラシエル基準値'],
        '④判定': d['haichi']['④判定'],
        '④過不足': d['haichi']['④過不足'],
        '割増金件数': d['kpi']['割増金発生 件数'],
        '法定休日NG職員数': legal_ng,
        '生支人工': d['kpi']['生活支援員 人工(換算)'],
        '世話人人工': d['kpi']['世話人 人工(換算)'],
        'シフト入力率': round(filled / max(1, len(named_rows) * days) * 100),
        '勤怠署名漏れ日数': kintai_missing,
        '日次未記入日数': check_missing,
        '経過日数': elapsed,
    }


class ModelManager:
    """facility_id → FacilityModel のプロセス内キャッシュ。"""

    def __init__(self, store):
        self.store = store
        self._models = {}

    def get(self, fid):
        if fid not in self._models:
            f = self.store.facility(fid)
            if f is None:
                return None
            self._models[fid] = FacilityModel(
                self.store.upload_path(f['filename']),
                self.store.overrides(fid),
                cache_dir=self.store.cache_dir)
        return self._models[fid]

    def set_cell(self, fid, sheet, col, row, value, username=None):
        self.store.set_override(fid, sheet, col, row, value, username=username)
        model = self.get(fid)
        if model:
            model.set_cell(sheet, col, row, value)

    def drop(self, fid):
        self._models.pop(fid, None)

    def summary(self, fid, force=False):
        """サマリを取得(キャッシュ優先。編集があるとキャッシュは自動失効)。"""
        if not force:
            cached = self.store.summaries().get(fid)
            if cached:
                return cached['data'], cached['computed_at']
        model = self.get(fid)
        if model is None:
            return None, None
        data = _facility_summary(model)
        self.store.save_summary(fid, data)
        import time
        return data, time.time()
