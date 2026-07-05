# -*- coding: utf-8 -*-
"""hyotei_check.py を埋め込んだ Colab ノートブックを生成する。"""
import json
from pathlib import Path

HERE = Path(__file__).parent
module_src = (HERE / "hyotei_check.py").read_text(encoding="utf-8")

intro_md = """\
# 職務評定表 定期チェック(Google Colab)

Google Drive の **「職務評定」フォルダ** 内の Excel(施設別)から、指定した月のデータを読み取り、

- **施設別に1つの Excel に集約**(H列:氏名 / I列:順位 / J列:評価 / E列:雇用形態 / F列:入社日 / C列:職種)
- **入力エラーを検出**(必須項目の空欄・順位の重複/欠番・数式エラーや値の異常・対象月シートの不在)

を行い、結果を Drive の `職務評定/チェック結果` フォルダに保存します。

## 使い方
1. 下のセルの **対象月** に確認したい月を入力(カンマ区切りで複数指定可。例: `2026-06` / `2026-04, 2026-06`)
2. メニューの「ランタイム」→「すべてのセルを実行」(または各セルを上から実行)
3. 初回は Google Drive へのアクセス許可を求められるので許可する
4. 実行結果がこの画面に表示され、集約 Excel が `チェック結果` フォルダに保存されます

## 月の判定について
シート名の表記が施設によりバラバラ(`2606` / `2026.6` / `R8.6` / `202606` / `宮下6月` / `Sheet1` など)なため、
**各シート内の「評価月」セルの日付を優先**して月を判定し、シート名は補助的に使います。
どちらからも判定できないシートは「月判定不可」としてエラー一覧に報告します。
"""

params_src = """\
#@title ① 設定(対象月などを入力して実行) { display-mode: "form" }

#@markdown **対象月**(カンマ区切りで複数指定可。例: 2026-06 または 2026-04, 2026-06)
対象月 = "2026-06"  #@param {type:"string"}

#@markdown **対象フォルダ**(マイドライブ内のパス)
対象フォルダ = "職務評定"  #@param {type:"string"}

#@markdown **結果の保存先**(対象フォルダ内に作られるサブフォルダ名)
結果フォルダ = "チェック結果"  #@param {type:"string"}

from google.colab import drive
drive.mount('/content/drive')

from pathlib import Path
FOLDER = Path('/content/drive/MyDrive') / 対象フォルダ
OUT_DIR = FOLDER / 結果フォルダ
assert FOLDER.is_dir(), f"フォルダが見つかりません: {FOLDER}"
print("対象フォルダ:", FOLDER)
"""

runner_src = """\
#@title ③ チェック実行(結果の表示と保存)
import time

months = [parse_month_token(t) for t in 対象月.replace('、', ',').split(',') if t.strip()]
assert months, "対象月が指定されていません"

paths = [p for p in sorted(FOLDER.glob('*.xlsx')) if not p.name.startswith('~$')]
print(f"対象月: {', '.join(months)} / 対象ファイル: {len(paths)} 件\\n")

records, issues = check_folder(paths, months)

OUT_DIR.mkdir(exist_ok=True)
stamp = time.strftime('%Y%m%d_%H%M')
out_path = OUT_DIR / f"職務評定チェック_{'_'.join(months)}_{stamp}.xlsx"
write_output(records, issues, months, out_path)

print(f"読み取った職員データ: {len(records)} 行")
print(f"検出したエラー: {len(issues)} 件")
print(f"出力ファイル: {out_path}\\n")

from collections import Counter

if issues:
    print('=' * 70)
    print('【エラー件数(種別ごと)】 ※明細は出力Excelの「エラー詳細」シート参照')
    print('=' * 70)
    for kind, n in Counter(i.kind for i in issues).most_common():
        print(f"  {kind}: {n} 件")

    # 件数の多い「空欄」「値の異常」以外は、重要なので明細も画面に出す
    important = [i for i in issues if i.kind not in ('空欄', '値の異常')]
    if important:
        print()
        print('=' * 70)
        print(f'【要確認のエラー明細】({len(important)} 件)')
        print('=' * 70)
        for i in sorted(important, key=lambda x: (x.kind, x.facility, str(x.month))):
            row = f" 行{i.row}" if i.row != '-' else ''
            print(f"[{i.kind}] {i.facility} / {i.month}{row}: {i.detail}")
else:
    print('エラーは見つかりませんでした。')

# 施設ごとの件数サマリー
c = Counter((r.facility, r.month) for r in records)
print()
print('=' * 70)
print('【施設別サマリー(読み取り件数)】')
print('=' * 70)
for (fac, m), n in sorted(c.items()):
    print(f"{fac} / {m}: {n} 名")
"""


def code_cell(src):
    return {"cell_type": "code", "metadata": {}, "execution_count": None,
            "outputs": [], "source": src.splitlines(keepends=True)}


def md_cell(src):
    return {"cell_type": "markdown", "metadata": {},
            "source": src.splitlines(keepends=True)}


nb = {
    "nbformat": 4,
    "nbformat_minor": 0,
    "metadata": {
        "colab": {"provenance": [], "name": "職務評定チェック.ipynb"},
        "kernelspec": {"name": "python3", "display_name": "Python 3"},
        "language_info": {"name": "python"},
    },
    "cells": [
        md_cell(intro_md),
        code_cell(params_src),
        code_cell("#@title ② チェックロジック(変更不要)\n" + module_src),
        code_cell(runner_src),
    ],
}

out = HERE / "職務評定チェック.ipynb"
out.write_text(json.dumps(nb, ensure_ascii=False, indent=1), encoding="utf-8")
print("wrote", out)
