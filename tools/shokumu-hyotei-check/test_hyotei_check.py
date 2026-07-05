# -*- coding: utf-8 -*-
"""実ファイルで観測したレイアウトを再現してロジックを検証するテスト。"""
import datetime as dt
from pathlib import Path

import openpyxl
from hyotei_check import (check_folder, parse_month_token, month_from_sheet_title,
                          write_output)

TMP = Path(__file__).parent / "testdata"
TMP.mkdir(exist_ok=True)


def build_sheet(ws, eval_date, people):
    """観測した様式を再現: B2=評価月ラベル, B3=日付, 6行目左ヘッダー,
    7-8行目は例(8行目のH/I/Jにラベル), 9行目からデータ, 残骸FALSE行。
    people: (氏名, 職種, 雇用形態, 入社日, H氏名, 順位, 評価)"""
    ws["B1"] = "ラシエル 従業員勤務評定表"
    ws["B2"] = "評価月"; ws["C2"] = "事業所名"; ws["E2"] = "評価者"
    ws["B3"] = eval_date; ws["C3"] = "ラシエルテスト"
    ws["B6"] = "氏　　名"; ws["C6"] = "職　　種"; ws["D6"] = "総　　評"
    ws["E6"] = "雇用形態"; ws["F6"] = "入社日"; ws["H6"] = "↓↓入力変更禁止↓↓"
    ws["B7"] = "○○　○○（例）"; ws["C7"] = "生活支援員"; ws["E7"] = "契社"
    ws["H7"] = "評定計算式"
    ws["B8"] = "○○　○○（例）"; ws["C8"] = "世話人"; ws["E8"] = "派遣"
    ws["H8"] = "氏名"; ws["I8"] = "順位"; ws["J8"] = "評価"
    r = 9
    for p in people:
        ws.cell(r, 2, p[0]); ws.cell(r, 3, p[1]); ws.cell(r, 4, "総評テキスト")
        ws.cell(r, 5, p[2]); ws.cell(r, 6, p[3])
        ws.cell(r, 8, p[4]); ws.cell(r, 9, p[5]); ws.cell(r, 10, p[6])
        r += 1
    for _ in range(5):  # 数式残骸
        ws.cell(r, 10, "FALSE")
        r += 1


def make_files():
    d = dt.date
    # 浄水型: シート名 "2602" だが評価月セルは 2026/6 → 評価月セルを優先すべき
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "2602"
    build_sheet(ws, d(2026, 6, 20), [
        ("瀬口　公子", "管理者", "契社", d(2026, 6, 1), "瀬口　公子", 1, "A"),
        ("井上美加", "生活支援員", "時給", d(2025, 12, 1), "井上美加", 2, "A"),
    ])
    wb.save(TMP / "【浄水】職務評定表（植松）2606.xlsx")

    # 名取型: シート名 "Sheet1"、順位の重複(4)と欠番(3)、空欄・値異常あり
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    build_sheet(ws, d(2026, 6, 22), [
        ("浜塚　淳", "生活支援員", "正社", d(2022, 6, 1), "浜塚　淳", 1, "A"),
        ("今野　千佳", "管理者候補", "時給", d(2022, 5, 8), "今野　千佳", 2, "A"),
        ("曽我　ゆみ", "世話人", "", d(2022, 3, 1), "曽我　ゆみ", 4, "B"),   # 雇用形態空欄
        ("瀬野　美智代", "世話人", "時給", "不明", "瀬野　美智代", 4, "B"),  # 入社日異常+順位重複
        ("南舘　恵理子", "世話人", "時給", d(2022, 10, 25), "南舘　恵理子", 5, "G"),  # 評価異常
        ("追加　太郎", "世話人", "時給", d(2023, 1, 1), "", "", ""),  # H未反映
    ])
    wb.save(TMP / "【名取】職務評定表（吉川）2026・6.xlsx")

    # 石巻型: 複数月シート "2025.4"〜"2026.6"
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for (y, m) in [(2026, 4), (2026, 6)]:
        ws = wb.create_sheet(f"{y}.{m}")
        build_sheet(ws, d(y, m, 15), [
            ("菅原　一郎", "施設長", "正社", d(2020, 4, 1), "菅原　一郎", 1, "A"),
            ("鈴木　二郎", "世話人", "時給", d(2021, 5, 1), "鈴木　二郎", 2, "#REF!"),  # 数式エラー
        ])
    wb.save(TMP / "【石巻】勤務評定表（菅原）2026.4.xlsx")

    # 下手野型: 6月ファイルのはずがシートは4月のみ → 対象月シート不在になるはず
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "202604"
    build_sheet(ws, d(2026, 4, 10), [
        ("米田　花子", "世話人", "時給", d(2024, 4, 1), "米田　花子", 1, "B"),
    ])
    wb.save(TMP / "【下手野】職務評定表（米田）202606.xlsx")

    # 宮下型: シート名 "宮下6月"、評価月セルなし(年はシート名から取れない)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "宮下6月"
    build_sheet(ws, "", [
        ("堀切　三郎", "世話人", "時給", d(2023, 6, 1), "堀切　三郎", 1, "C"),
    ])
    ws["B3"] = dt.date(2026, 6, 1)  # 評価月セルは日付あり
    wb.save(TMP / "【宮下】職務評定（堀切）202606.xlsx")

    # 一時ファイルはスキップされること
    (TMP / "~$【浄水】職務評定表（植松）2606.xlsx").write_bytes(b"junk")


def main():
    make_files()

    # 月指定パーサ
    assert parse_month_token("2026-06") == "2026-06"
    assert parse_month_token("2026.6") == "2026-06"
    assert parse_month_token("202606") == "2026-06"
    assert parse_month_token("2026/6") == "2026-06"
    assert parse_month_token("2026年6月") == "2026-06"

    # シート名パーサ
    assert month_from_sheet_title("2025.4") == (2025, 4)
    assert month_from_sheet_title("R8.6") == (2026, 6)
    assert month_from_sheet_title("202604") == (2026, 4)
    assert month_from_sheet_title("2602") == (2026, 2)
    assert month_from_sheet_title("宮下6月") == (None, 6)
    assert month_from_sheet_title("Sheet1") == (None, None)

    months = ["2026-06"]
    paths = list(TMP.glob("*.xlsx"))
    records, issues = check_folder(paths, months)

    facs = {r.facility for r in records}
    assert facs == {"浄水", "名取", "石巻", "宮下"}, facs

    # 浄水: シート名2602でも評価月セルで2026-06と判定され読めている
    josui = [r for r in records if r.facility == "浄水"]
    assert len(josui) == 2 and all(r.month == "2026-06" for r in josui)

    # 石巻: 2026-06 のシートだけ読まれる(2026-04は対象外)
    ishi = [r for r in records if r.facility == "石巻"]
    assert len(ishi) == 2 and all(r.sheet_name == "2026.6" for r in ishi)

    kinds = {}
    for i in issues:
        kinds.setdefault(i.kind, []).append(i)

    assert any("雇用形態" in i.detail for i in kinds.get("空欄", [])), kinds.get("空欄")
    assert any("順位 [4]" in i.detail for i in kinds.get("順位の重複", []))
    assert any("[3" in i.detail for i in kinds.get("順位の欠番", []))
    assert any("評価 'G'" in i.detail for i in kinds.get("値の異常", []))
    assert any("入社日 '不明'" in i.detail for i in kinds.get("値の異常", []))
    assert any("#REF!" in i.detail for i in kinds.get("数式エラー", []))
    assert any("追加　太郎" in i.detail for i in kinds.get("不整合", []))
    assert any(i.facility == "下手野" for i in kinds.get("対象月シート不在", []))
    # 下手野以外に不在エラーが出ないこと
    assert all(i.facility == "下手野" for i in kinds.get("対象月シート不在", []))

    out = TMP / "output.xlsx"
    write_output(records, issues, months, out)
    wb = openpyxl.load_workbook(out)
    assert wb.sheetnames[0] == "エラー一覧"
    assert set(wb.sheetnames) == {"エラー一覧", "浄水", "名取", "石巻", "宮下"}
    ws = wb["浄水"]
    assert ws.cell(2, 2).value == "瀬口　公子" and ws.cell(2, 3).value == 1

    # 複数月指定: 石巻は4月と6月の両方が読まれる
    records2, _ = check_folder(paths, ["2026-04", "2026-06"])
    ishi2 = [r for r in records2 if r.facility == "石巻"]
    assert {r.month for r in ishi2} == {"2026-04", "2026-06"}

    print("全テスト成功:", len(records), "件のレコード,", len(issues), "件のエラー検出")
    for i in issues:
        print(f"  [{i.kind}] {i.facility} {i.month} 行{i.row}: {i.detail}")


if __name__ == "__main__":
    main()
