"""取り込み時チェック: ブック健全性+業務チェック+前月比較のレポート生成。"""
import re

from . import rodo
from .model_manager import INPUT_SHEET, MASTER_SHEET


def staff_monthly(model, aliases=None):
    """職員別月次集計(確認ブックの施設シート相当)。

    合計時間は管理用シートCU列(検証済み数式)を正とし、
    就業日数・夜勤回数・有給時間はシフト記号から数える。
    """
    maps = model.schedule_maps(aliases)
    ncodes = rodo.night_codes_from_master(maps['code_times'])
    hours = model.code_hours()
    planned, meta = model.planned_hours(aliases)

    per = {}
    for key, byd in maps['schedule'].items():
        rec = per.setdefault(key, {'days': 0, 'night': 0, 'paid_hours': 0.0, 'codes': {}})
        for d, shifts in byd.items():
            worked = False
            for s in shifts:
                code = rodo.norm_shift(s['code'])
                rec['codes'][code] = rec['codes'].get(code, 0) + 1
                if code in ncodes:
                    rec['night'] += 1
                if _is_paid_leave(code):
                    rec['paid_hours'] += (hours.get(code) or (0, 0))[1] or 0
                elif not rodo.is_off_shift(code):
                    worked = True
            if worked:
                rec['days'] += 1

    rows = []
    for key in sorted(set(per) | set(planned), key=lambda k: str(meta.get(k, {}).get('name', k))):
        m = meta.get(key, {})
        p = per.get(key, {'days': 0, 'night': 0, 'paid_hours': 0.0})
        total = planned.get(key)
        rows.append({
            '職員名': m.get('name') or key,
            '支給': m.get('pay', ''),
            '就業日数': p['days'],
            '合計時間': round(total, 2) if isinstance(total, (int, float)) else '',
            '夜勤回数': p['night'],
            '有給時間': round(p['paid_hours'], 1),
        })
    return rows


def _is_paid_leave(code):
    s = rodo.norm_shift(code)
    return bool(re.match(r'^(有|常有|非有)', s)) and not s.startswith('有a')


def build_report(model, store, fid, masters_warnings=None):
    aliases = rodo.parse_aliases(store.get_setting('name_aliases'))
    maps = model.schedule_maps(aliases)
    hours = model.code_hours()
    y, m = maps['year'], maps['month']

    # --- A. ブック健全性 ---
    known = set(maps['code_times']) | set(hours)
    unknown = {}
    for key, byd in maps['schedule'].items():
        for d, shifts in byd.items():
            for s in shifts:
                code = rodo.norm_shift(s['code'])
                if code not in known and not rodo.is_off_shift(code):
                    unknown.setdefault(code, []).append(d)
    broken = 0
    for sheet in (INPUT_SHEET, '管理用シート', MASTER_SHEET):
        for src in model.wb.formulas.get(sheet, {}).values():
            if '#REF!' in src:
                broken += 1
    dup = [f for f in store.facilities()
           if f['id'] != fid and f['year'] == y and f['month'] == m
           and f['name'] == store.facility(fid)['name']]

    # --- B. 業務チェック ---
    d = model.dashboard()
    ncodes = rodo.night_codes_from_master(maps['code_times'])
    cal = rodo.night_calendar(maps['by_day'], ncodes)
    zero_days, one_days = [], []
    for day in range(1, model.days_in_month() + 1):
        n = len(cal.get(f'{y}/{m}/{day}', []))
        if n == 0:
            zero_days.append(day)
        elif n == 1:
            one_days.append(day)
    grid = model.grid()
    named = [r for r in grid['rows'] if r['name']]
    filled = sum(1 for r in named for c in r['cells'] if c)
    fill_rate = round(filled / max(1, len(named) * model.days_in_month()) * 100)
    legal_ng = [s['name'] for s in d['staff'] if '割増' in str(s['legal_holiday'])]

    sm = staff_monthly(model, aliases)
    anomalies = []
    for r in sm:
        t = r['合計時間']
        if isinstance(t, (int, float)):
            if t > 200:
                anomalies.append(f"{r['職員名']}: 月{t}時間(200h超)")
            if t == 0 and r['就業日数'] > 0:
                anomalies.append(f"{r['職員名']}: 勤務日ありで合計0時間(記号の時間未設定?)")

    totals = {
        '総労働時間': round(sum(r['合計時間'] for r in sm if isinstance(r['合計時間'], (int, float))), 1),
        '夜勤回数': sum(r['夜勤回数'] for r in sm),
        '有給時間': round(sum(r['有給時間'] for r in sm), 1),
        '職員数': len([r for r in sm if r['就業日数'] > 0 or r['夜勤回数'] > 0]),
    }

    # --- C. 前月比較 ---
    prev = None
    py, pm = (y - 1, 12) if m == 1 else (y, m - 1)
    name = store.facility(fid)['name']
    prev_fac = next((f for f in store.facilities()
                     if f['name'] == name and f['year'] == py and f['month'] == pm), None)
    if prev_fac:
        prev_rep = store.get_report(prev_fac['id'])
        if prev_rep and 'totals' in prev_rep:
            prev = {'年月': f'{py}/{pm}'}
            for k, v in totals.items():
                pv = prev_rep['totals'].get(k)
                prev[k] = {'前月': pv, '今月': v,
                           '増減': round(v - pv, 1) if isinstance(pv, (int, float)) else ''}

    warn_count = (len(masters_warnings or []) + len(unknown) + (1 if broken else 0)
                  + (1 if dup else 0) + len(zero_days) + len(anomalies) + len(legal_ng)
                  + (1 if '未達' in str(d['haichi'].get('基準判定', '')) else 0))
    return {
        'year': y, 'month': m,
        'summary': {'警告件数': warn_count, 'シフト入力率': fill_rate, **totals},
        'totals': totals,
        'format': {
            'masters_warnings': masters_warnings or [],
            'unknown_codes': [{'記号': c, '日付': ' '.join(ds[:8])} for c, ds in sorted(unknown.items())],
            'broken_formulas': broken,
            'duplicate_import': bool(dup),
        },
        'business': {
            'haichi': d['haichi'],
            '割増金発生件数': d['kpi'].get('割増金発生 件数'),
            '法定休日NG': legal_ng,
            '夜勤0名の日': zero_days,
            '夜勤1名の日': one_days,
            'シフト入力率': fill_rate,
            '異常値': anomalies,
        },
        'staff_monthly': sm,
        'prev': prev,
    }
