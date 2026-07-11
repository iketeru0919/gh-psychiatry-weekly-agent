"""労務チェック: 勤怠CSV照合・タイミー明細集計・夜勤体制チェック。

単体HTMLツール(照合ツールv2.14 / タイミー集計v2.6 / 夜勤チェッカー)の
ロジックを移植したもの。シフト側のデータは検証済み数式エンジンから取るため、
Excelの固定位置パースは行わない。
"""
import csv
import io
import re
import unicodedata

# ---------------------------------------------------------------- 共通

OFF_SHIFT_WORDS = ['休', '希休', '希望休', '有', '有休', '常有', '欠', '公休', '振休', '特休', '代休', '年休']


def norm_name(v, aliases=None):
    """氏名の正規化(NFKC・空白除去・末尾の丸数字/連番除去)+エイリアス解決。"""
    s = unicodedata.normalize('NFKC', str(v or '')).upper()
    s = re.sub(r'[\s ​-‍﻿]', '', s)
    s = re.sub(r'[①②③④⑤]$', '', s)
    s = re.sub(r'\(?[0-9]+\)?$', '', s).strip()
    if aliases:
        return aliases.get(s, s)
    return s


def parse_aliases(text):
    """「基準名=別名1,別名2」形式のテキスト → {正規化別名: 正規化基準名}"""
    out = {}
    for line in (text or '').splitlines():
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        parts = [norm_name(p) for p in re.split(r'[=,、]', line) if p.strip()]
        if len(parts) >= 2:
            for p in parts:
                out[p] = parts[0]
    return out


def norm_shift(v):
    return re.sub(r'（.*?）|\(.*?\)', '', str(v or '').replace(' ', '').replace('　', '')).strip()


def is_off_shift(code):
    s = norm_shift(code)
    if not s:
        return True
    return any(s == k or s.startswith(k) for k in OFF_SHIFT_WORDS)


def is_dispatch(*vals):
    return any('派遣' in str(v or '') for v in vals)


def _decode(content: bytes) -> str:
    if content[:3] == b'\xef\xbb\xbf':
        return content.decode('utf-8-sig')
    try:
        t = content.decode('utf-8')
        if '�' not in t:
            return t
    except UnicodeDecodeError:
        pass
    try:
        return content.decode('cp932')
    except UnicodeDecodeError:
        return content.decode('utf-8', 'replace')


def _rows(content: bytes):
    text = _decode(content)
    return list(csv.DictReader(io.StringIO(text)))


def time_to_min(v):
    s = str(v or '').strip()
    if not s:
        return None
    if s == '24:00':
        return 1440
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def min_to_time(m):
    if m is None:
        return ''
    m = round(m)
    return f'{(m // 60) % 24:02d}:{m % 60:02d}'


def date_key(v):
    m = re.match(r'(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})', str(v or '').strip())
    if m:
        return f'{int(m.group(1))}/{int(m.group(2))}/{int(m.group(3))}'
    return ''


# ---------------------------------------------------------------- 勤怠CSV

def parse_kintai(content: bytes, aliases=None, excluded_keys=frozenset()):
    """タイムカードCSV → {key: {date: [rec]}} と {key: [rec]}"""
    by_day, by_name, names = {}, {}, {}
    for o in _rows(content):
        name = str(o.get('従業員氏名') or '').strip()
        if not name or name == '従業員氏名':
            continue
        key = norm_name(name, aliases)
        if is_dispatch(name) or key in excluded_keys:
            continue
        d = date_key(o.get('年/月/日'))
        if not d:
            continue
        rec = {
            'name': name, 'date': d,
            'in': str(o.get('出勤打刻') or '').strip(), 'out': str(o.get('退勤打刻') or '').strip(),
            'in_min': time_to_min(o.get('出勤打刻')), 'out_min': time_to_min(o.get('退勤打刻')),
            'comment': str(o.get('コメント') or '').strip(),
        }
        by_day.setdefault(key, {}).setdefault(d, []).append(rec)
        by_name.setdefault(key, []).append(rec)
        names.setdefault(key, name)
    return {'by_day': by_day, 'by_name': by_name, 'names': names}


def _break_hours(comment):
    s = unicodedata.normalize('NFKC', str(comment or '')).replace(' ', '')
    m = re.search(r'(?:休憩深夜|深夜休憩|休憩)(\d+(?:\.\d+)?)時間', s)
    return float(m.group(1)) if m else 0.0


def _comment_out_min(comment):
    s = unicodedata.normalize('NFKC', str(comment or '')).replace(' ', '')
    m = re.search(r'退勤(\d{1,2})[:：](\d{2})', s)
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None


def _dord(d):
    y, m, dd = (int(x) for x in d.split('/'))
    return y * 372 + m * 31 + dd


def actual_monthly_hours(by_name):
    """打刻をペアリングして職員別の月間実働時間(概算)を出す。"""
    out = {}
    for key, recs in by_name.items():
        sorted_r = sorted(recs, key=lambda r: _dord(r['date']))
        used = set()
        total = 0.0
        for i, r in enumerate(sorted_r):
            if i in used:
                continue
            if r['in_min'] is not None and r['out_min'] is not None:
                end = r['out_min'] + (1440 if r['out_min'] < r['in_min'] else 0)
                total += max(0.0, (end - r['in_min']) / 60 - _break_hours(r['comment']))
                used.add(i)
                continue
            if r['in_min'] is not None and r['out_min'] is None:
                if not (15 * 60 <= r['in_min'] < 1440):   # 夜勤入りの時間帯のみペア探索
                    continue
                for j in range(i + 1, len(sorted_r)):
                    if j in used:
                        continue
                    n = sorted_r[j]
                    diff = _dord(n['date']) - _dord(r['date'])
                    if diff > 1:
                        break
                    if diff > 0 and n['in_min'] is None and n['out_min'] is not None and 5 * 60 <= n['out_min'] < 11 * 60:
                        br = _break_hours(n['comment']) or _break_hours(r['comment'])
                        total += max(0.0, (1440 + n['out_min'] - r['in_min']) / 60 - br)
                        used.add(i)
                        used.add(j)
                        break
                if i not in used:
                    co = _comment_out_min(r['comment'])
                    if co is not None and 5 * 60 <= co < 11 * 60:
                        total += max(0.0, (1440 + co - r['in_min']) / 60 - _break_hours(r['comment']))
                        used.add(i)
        out[key] = round(total, 2)
    return out


def classify_shift(code, times):
    """day / nightIn / nightOut / off。times=(start_min, end_min or None)"""
    s = norm_shift(code)
    if times is None or (times[0] is None and times[1] is None):
        return 'off'
    start, end = times
    if '明' in s or (start is not None and start <= 30 and end is not None and end <= 12 * 60):
        return 'nightOut'
    if '入' in s or end == 1440 or (start is not None and start >= 12 * 60 and end is not None and end >= 23 * 60):
        return 'nightIn'
    return 'day'


def compare_daily(schedule, kintai, tol=15):
    """日別のズレ一覧。schedule: {key: {date: [{'code','name','times'}]}}"""
    rows = []
    keys = set(schedule) | set(kintai['by_day'])
    for key in keys:
        dates = set(schedule.get(key, {})) | set(kintai['by_day'].get(key, {}))
        for d in sorted(dates, key=_dord):
            shifts = schedule.get(key, {}).get(d, [])
            clocks = kintai['by_day'].get(key, {}).get(d, [])
            ins = [c['in_min'] for c in clocks if c['in_min'] is not None]
            outs = [c['out_min'] for c in clocks if c['out_min'] is not None]
            csv_in = ' / '.join(c['in'] for c in clocks if c['in'])
            csv_out = ' / '.join(c['out'] for c in clocks if c['out'])
            disp_name = (shifts[0]['name'] if shifts else clocks[0]['name'] if clocks else key)
            if not shifts:
                if ins or outs:
                    rows.append({'職員名': disp_name, '日付': d, 'シフト': '空欄', '予定開始': '', '予定終了': '',
                                 'CSV出勤': csv_in, 'CSV退勤': csv_out, '判定': '要確認',
                                 '内容': '空欄=休み扱いですが、CSVに打刻があります'})
                continue
            for s in shifts:
                code, times = s['code'], s['times']
                cls = classify_shift(code, times)
                status, msg = 'OK', []
                if is_off_shift(code) or cls == 'off':
                    if ins or outs:
                        status, msg = '要確認', ['休み系シフトですが打刻があります']
                elif times is None:
                    status, msg = '未登録', ['シフトがマスタ未登録']
                elif cls == 'nightIn':
                    if not ins:
                        status, msg = '要確認', ['夜勤入り・出勤打刻なし']
                    else:
                        best = min(abs(v - times[0]) for v in ins)
                        if best > tol:
                            status, msg = 'ズレ', [f'出勤 {best}分ズレ']
                elif cls == 'nightOut':
                    if not outs:
                        status, msg = '要確認', ['夜勤明け・退勤打刻なし']
                    else:
                        best = min(abs(v - times[1]) for v in outs)
                        if best > tol:
                            status, msg = 'ズレ', [f'退勤 {best}分ズレ']
                else:
                    if not ins and not outs:
                        status, msg = '要確認', ['勤務予定ですが打刻がありません']
                    else:
                        if times[0] is not None:
                            if not ins:
                                status = '要確認'
                                msg.append('出勤打刻なし')
                            else:
                                best = min(abs(v - times[0]) for v in ins)
                                if best > tol:
                                    status = 'ズレ' if status == 'OK' else status
                                    msg.append(f'出勤 {best}分ズレ')
                        if times[1] is not None:
                            if not outs:
                                status = '要確認'
                                msg.append('退勤打刻なし')
                            else:
                                best = min(abs(v - times[1]) for v in outs)
                                if best > tol:
                                    status = 'ズレ' if status == 'OK' else status
                                    msg.append(f'退勤 {best}分ズレ')
                if status != 'OK':
                    rows.append({'職員名': s['name'], '日付': d, 'シフト': code,
                                 '予定開始': min_to_time(times[0]) if times else '',
                                 '予定終了': min_to_time(times[1]) if times else '',
                                 'CSV出勤': csv_in, 'CSV退勤': csv_out,
                                 '判定': status, '内容': '、'.join(msg)})
    return rows


def compare_monthly(planned_by_key, meta_by_key, kintai, hour_tol=0.25):
    actual = actual_monthly_hours(kintai['by_name'])
    keys = set(planned_by_key) | set(actual)
    out = []
    for key in sorted(keys):
        meta = meta_by_key.get(key, {})
        planned = round(planned_by_key.get(key, 0) or 0, 2)
        act = actual.get(key, 0)
        diff = round(act - planned, 2)
        out.append({
            '職員名': meta.get('name') or kintai['names'].get(key, key),
            '職種': meta.get('role', ''),
            '雇用形態': meta.get('pay', ''),
            '勤務表CU時間': planned, 'CSV実績概算時間': act, '差分': diff,
            '判定': 'OK' if abs(diff) <= hour_tol else '要確認',
        })
    return out


# ---------------------------------------------------------------- タイミー明細

def parse_timee(content: bytes):
    out = []
    for o in _rows(content):
        d = date_key(o.get('求人日'))
        if not d:
            continue
        if is_dispatch(o.get('姓'), o.get('名'), o.get('求人タイトル')):
            continue

        def num(k):
            try:
                return float(str(o.get(k) or '0').replace(',', '') or 0)
            except ValueError:
                return 0.0
        wid = (o.get('ワーカーID（ワーカーによって一意な値）') or o.get('ワーカーID')
               or f"{o.get('姓', '')}_{o.get('名', '')}")
        midnight = num('深夜労働時間')
        midnight_pay = num('深夜割増手当')
        bs, be = time_to_min(o.get('休憩開始時間')), time_to_min(o.get('休憩終了時間'))
        br = None
        if bs is not None and be is not None:
            br = round(((be + (1440 if be < bs else 0)) - bs) / 60, 2)
        ym = '/'.join(d.split('/')[:2])
        out.append({
            '年月': ym, '求人日': d,
            '氏名': f"{o.get('姓', '')} {o.get('名', '')}".strip(), 'ワーカーID': wid,
            '勤務区分': '夜勤' if (midnight > 0 or midnight_pay > 0) else '日勤',
            'チェックイン': str(o.get('QRチェックイン時間') or ''), 'チェックアウト': str(o.get('QRチェックアウト時間') or ''),
            '休憩時間': br, '休憩1時間': br is not None and abs(br - 1) < 0.01,
            '実働時間': num('実働時間'), '深夜労働時間': midnight,
            '合計報酬額': num('合計報酬額'), 'タイミー利用料': num('利用料'),
        })
    return out


def timee_summary(details):
    """施設内の月別集計+日別+両方勤務者。"""
    months = {}
    for x in details:
        months.setdefault(x['年月'], []).append(x)
    monthly = []
    both_rows = []
    for ym in sorted(months):
        arr = months[ym]
        day = [x for x in arr if x['勤務区分'] == '日勤']
        night = [x for x in arr if x['勤務区分'] == '夜勤']
        one_hour = [x for x in night if x['休憩1時間']]
        # 日勤夜勤両方
        byw = {}
        for x in arr:
            byw.setdefault((x['ワーカーID'], x['氏名']), []).append(x)
        both = 0
        for (wid, name), items in byw.items():
            dd = [x for x in items if x['勤務区分'] == '日勤']
            nn = [x for x in items if x['勤務区分'] == '夜勤']
            if dd and nn:
                both += 1
                both_rows.append({'年月': ym, '氏名': name, '日勤回数': len(dd), '夜勤回数': len(nn),
                                  '実働時間': round(sum(x['実働時間'] for x in items), 1),
                                  '総コスト': round(sum(x['合計報酬額'] + x['タイミー利用料'] for x in items))})
        monthly.append({
            '年月': ym, '明細件数': len(arr),
            '日勤回数': len(day), '夜勤回数': len(night),
            '日勤者数': len({x['ワーカーID'] for x in day}), '夜勤者数': len({x['ワーカーID'] for x in night}),
            '夜勤1h休憩回数': len(one_hour),
            '夜勤1h休憩率': round(len(one_hour) / len(night) * 100, 1) if night else 0,
            '実働時間': round(sum(x['実働時間'] for x in arr), 1),
            '日勤勤務時間': round(sum(x['実働時間'] for x in day), 1),
            '夜勤勤務時間': round(sum(x['実働時間'] for x in night), 1),
            '合計報酬額': round(sum(x['合計報酬額'] for x in arr)),
            'タイミー利用料': round(sum(x['タイミー利用料'] for x in arr)),
            '総コスト': round(sum(x['合計報酬額'] + x['タイミー利用料'] for x in arr)),
            '日勤夜勤両方者数': both,
        })
    daily = {}
    for x in details:
        daily.setdefault(x['求人日'], []).append(x)
    daily_rows = [{
        '求人日': d,
        '日勤回数': len([x for x in arr if x['勤務区分'] == '日勤']),
        '夜勤回数': len([x for x in arr if x['勤務区分'] == '夜勤']),
        '実働時間': round(sum(x['実働時間'] for x in arr), 1),
        '総コスト': round(sum(x['合計報酬額'] + x['タイミー利用料'] for x in arr)),
    } for d, arr in sorted(daily.items(), key=lambda kv: _dord(kv[0]))]
    return {'monthly': monthly, 'daily': daily_rows, 'both': both_rows}


# ---------------------------------------------------------------- 夜勤体制

def night_codes_from_master(code_times):
    """マスタの時刻から夜勤入りコードを判定(夜勤チェッカー移植)。"""
    out = {}
    for code, (start, end, overnight) in code_times.items():
        if start is None:
            continue
        is_night = overnight or end == 1440 or (
            start >= 16 * 60 and end is not None and end < 12 * 60)
        # 明けコードは除外
        if '明' in norm_shift(code):
            is_night = False
        if is_night:
            out[code] = (start, end)
    return out


def night_calendar(schedule_by_day, night_codes):
    """{date: [{'name','code'}]} → 日別夜勤者リスト(夜勤コードのみ)。"""
    out = {}
    for d, entries in schedule_by_day.items():
        out[d] = [e for e in entries if norm_shift(e['code']) in night_codes]
    return out
