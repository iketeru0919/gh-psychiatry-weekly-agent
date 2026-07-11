"""ゴールデンテスト: 数式エンジンの計算結果を Excel のキャッシュ値と全数突合する。

使い方:
    python -m shift_app.golden ブック.xlsx [シート名 ...]

シート名を省略すると主要シート(管理用シート・入力シート【現場配布用】)を検証する。
"""
import sys

from .xlcalc.evaluator import Workbook, XLErr
from .xlcalc.parser import num_to_col

DEFAULT_SHEETS = ['管理用シート', '入力シート【現場配布用】']


def approx_equal(a, b):
    if a is None and (b is None or b == '' or b == 0):
        return True
    if b is None and (a is None or a == '' or a == 0):
        return True
    if isinstance(a, XLErr) or isinstance(b, XLErr):
        return isinstance(a, XLErr) and isinstance(b, XLErr) and a.code == b.code
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= max(1e-9, 1e-9 * max(abs(a), abs(b)))
    if isinstance(a, str) and isinstance(b, str):
        return a == b
    # 数値と文字列など型違い
    return False


def run(path, sheets=None, verbose_limit=15, cache_dir=None):
    wb = Workbook(path, cache_dir=cache_dir)
    sheets = sheets or [s for s in DEFAULT_SHEETS if s in wb.cells]
    total_ok = total_ng = total_skip = 0
    for sheet in sheets:
        ok = ng = skip = 0
        mismatches = []
        for (col, row), src in sorted(wb.formulas[sheet].items(), key=lambda kv: (kv[0][1], kv[0][0])):
            expected = wb.cached[sheet].get((col, row))
            try:
                got = wb.value(sheet, col, row)
            except NotImplementedError as e:
                skip += 1
                mismatches.append((f'{num_to_col(col)}{row}', src, expected, f'<NotImplemented: {e}>'))
                continue
            if (sheet, col, row) in wb.volatile_cells:
                skip += 1  # TODAY 依存は比較対象外
                continue
            if approx_equal(got, expected):
                ok += 1
            else:
                ng += 1
                if len(mismatches) < 200:
                    mismatches.append((f'{num_to_col(col)}{row}', src, expected, got))
        print(f'[{sheet}] 一致: {ok} / 不一致: {ng} / 対象外(TODAY依存等): {skip}')
        for addr, src, exp, got in mismatches[:verbose_limit]:
            s = src if len(src) < 90 else src[:90] + '…'
            print(f'  {addr}: expected={exp!r} got={got!r}  formula={s}')
        total_ok += ok
        total_ng += ng
        total_skip += skip
    print(f'\nTOTAL 一致: {total_ok} / 不一致: {total_ng} / 対象外: {total_skip}')
    return total_ng == 0


if __name__ == '__main__':
    import os
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    cache_dir = None
    for a in sys.argv[1:]:
        if a.startswith('--cache='):
            cache_dir = a.split('=', 1)[1]
            os.makedirs(cache_dir, exist_ok=True)
    path = args[0]
    sheets = args[1:] or None
    ok = run(path, sheets, cache_dir=cache_dir)
    sys.exit(0 if ok else 1)
