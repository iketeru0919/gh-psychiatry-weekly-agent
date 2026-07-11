"""施設ブックの計算モデル管理。

取り込んだ Excel の数式モデル(xlcalc.Workbook)に UI からの入力(上書き)を
適用し、シフトグリッド・管理ダッシュボードのデータを提供する。
計算は常にブック内の数式そのもので行うため、Excel との完全一致が保たれる。
"""
import calendar

from .xlcalc.evaluator import Workbook, XLErr, serial_to_date
from .xlcalc.parser import col_to_num, num_to_col

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


# 管理用シート上の手入力セル(④区分別利用者数・他施設兼務)
ADMIN_INPUT_CELLS = {
    '区分1': (_C('CK'), 70), '区分2': (_C('CO'), 70), '区分3': (_C('CS'), 70),
    '区分4': (_C('CW'), 70), '区分5': (_C('DA'), 70), '区分6': (_C('DE'), 70),
    'サビ管 他施設兼務': (_C('DO'), 71), '管理者 他施設兼務': (_C('DO'), 75),
}


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

    def set_cell(self, fid, sheet, col, row, value):
        self.store.set_override(fid, sheet, col, row, value)
        model = self.get(fid)
        if model:
            model.set_cell(sheet, col, row, value)

    def drop(self, fid):
        self._models.pop(fid, None)
