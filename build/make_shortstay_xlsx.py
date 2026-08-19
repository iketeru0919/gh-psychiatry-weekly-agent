# -*- coding: utf-8 -*-
"""ショートステイ・体験利用 管理台帳（Excel）を生成する。

設計の要:
  - 入力は「1滞在 = 1行」(01_予約入力) のみ。他シートは全て数式で導出する。
  - 部屋数は M_部屋 の行数で決まる。数式は最大4室ぶん用意し、
    未登録の室は空欄を返すので、1室施設・2室施設で同じファイルが使える。
  - 空室・稼働は「その日の夜に泊まるか」で数える(退所日の夜は空き)。
  - 重い配列数式は「表示中の1か月ぶん(31行)」に限定し、期間集計は
    滞在1件あたり1回のSUMPRODUCTで済ませる。日付を全部展開しない。
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment
import datetime

OUT = "/home/user/gh-psychiatry-weekly-agent/build/ショートステイ管理台帳.xlsx"

FONT = "Meiryo"
MAX_ROOMS = 4
BK_LAST = 204          # 01_予約入力 のデータ最終行(200件)
USER_ROWS = 30
CAT_ROWS = 10
MONTH0 = datetime.date(2026, 4, 1)

F_INPUT  = PatternFill("solid", fgColor="FFF7DD")
F_AUTO   = PatternFill("solid", fgColor="F2F2F2")
F_HEAD   = PatternFill("solid", fgColor="2F6690")
F_SUB    = PatternFill("solid", fgColor="DCE6EF")
F_WARN   = PatternFill("solid", fgColor="FBD3D3")
F_WARN2  = PatternFill("solid", fgColor="FCE4C8")
F_OK     = PatternFill("solid", fgColor="D9EDDC")
F_TENT   = PatternFill("solid", fgColor="EDE3F5")
F_WKND   = PatternFill("solid", fgColor="E8EEF4")

C_HEAD  = Font(name=FONT, size=10, bold=True, color="FFFFFF")
C_SUB   = Font(name=FONT, size=10, bold=True, color="1F3B52")
C_BODY  = Font(name=FONT, size=10)
C_AUTO  = Font(name=FONT, size=10, color="666666", italic=True)
C_TITLE = Font(name=FONT, size=13, bold=True, color="1F3B52")
C_NOTE  = Font(name=FONT, size=9, color="666666")
C_BOLD  = Font(name=FONT, size=10, bold=True)

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()
wb.remove(wb.active)

def sheet(name):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    return ws

def head_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = C_HEAD; c.fill = F_HEAD; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30

def title(ws, text, sub=None):
    ws["A1"] = text; ws["A1"].font = C_TITLE
    if sub:
        ws["A2"] = sub; ws["A2"].font = C_NOTE

def label_input(ws, row, label, value, note="", fmt=None, col=1):
    L = get_column_letter(col); V = get_column_letter(col + 1); N = get_column_letter(col + 2)
    ws[f"{L}{row}"] = label; ws[f"{L}{row}"].font = C_SUB
    ws[f"{L}{row}"].fill = F_SUB; ws[f"{L}{row}"].border = BORDER
    c = ws[f"{V}{row}"]; c.value = value; c.fill = F_INPUT; c.font = C_BODY; c.border = BORDER
    if fmt: c.number_format = fmt
    if note:
        ws[f"{N}{row}"] = note; ws[f"{N}{row}"].font = C_NOTE

def label_auto(ws, row, label, formula, note="", fmt=None, col=1):
    L = get_column_letter(col); V = get_column_letter(col + 1); N = get_column_letter(col + 2)
    ws[f"{L}{row}"] = label; ws[f"{L}{row}"].font = C_SUB
    ws[f"{L}{row}"].fill = F_SUB; ws[f"{L}{row}"].border = BORDER
    c = ws[f"{V}{row}"]; c.value = formula; c.fill = F_AUTO; c.font = C_AUTO; c.border = BORDER
    if fmt: c.number_format = fmt
    if note:
        ws[f"{N}{row}"] = note; ws[f"{N}{row}"].font = C_NOTE

# ── シート参照（全てクォート） ────────────────────────────────────
BK   = "'01_予約入力'"
MON  = "'02_月間表'"
MSET = "'M_施設設定'"
MROOM= "'M_部屋'"
MUSER= "'M_利用者'"
MCAT = "'M_区分'"

R = lambda col: f"{BK}!${col}$5:${col}${BK_LAST}"
NAME_, CAT_, ROOM_, IN_, OUT_, ST_, ALW_ = (R("B"), R("C"), R("D"), R("E"), R("G"), R("K"), R("U"))
VALID = f'({IN_}<>"")*({ST_}<>"キャンセル")'

CAP   = f"{MSET}!$B$6"     # 部屋数（定員）
UROW  = f"{MUSER}!$A$4:$A$60"
UALW  = f"{MUSER}!$B$4:$B$60"
CROW  = f"{MCAT}!$A$4:$A$20"
CCNT  = f"{MCAT}!$B$4:$B$20"
RROW  = f"{MROOM}!$A$4:$A$20"


# ══════════════════════════════════════════════════════════════════
# 00_使い方
# ══════════════════════════════════════════════════════════════════
ws = sheet("00_使い方")
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 100
title(ws, "ショートステイ・体験利用 管理台帳",
      "入力するのは 01_予約入力 と、M_ で始まるマスタだけです。02〜07 は全て自動計算です。")

rows = [
    ("SEC", "はじめに設定すること", ""),
    ("① M_施設設定", "施設名・棟・標準の入所/退所時刻・住所・TEL/FAX を入れます。"),
    ("② M_部屋", "部屋名を1行1室で登録します。★ここの行数が、そのまま定員（部屋数）になります。"),
    ("③ M_利用者", "氏名と支給量（受給者証の月あたり上限日数）。支給量が空欄の方は残日数を管理しません。"),
    ("④ M_区分・M_送迎", "区分（SS・無料・契約など）と、その区分を支給量に数えるかどうか。"),
    ("SEC", "★ 部屋数が施設ごとに違う場合（1室の施設と2室の施設）", ""),
    ("同じファイルで使えます", "1室の施設は M_部屋 に1行だけ、2室の施設は2行だけ登録してください。"
                        "定員・空室数・空き記号（○△×）は全て M_部屋 の行数から計算しているので、"
                        "シートも数式も一切変えずに両方に対応します。最大4室まで。"),
    ("部屋を減らしたとき", "2室から1室に変更したのに、古い予約が「部屋2」のまま残っていると、"
                        "01_予約入力 のS列に「⚠部屋が未登録」と赤で出ます。その予約の部屋を選び直してください。"),
    ("1室のときの表示", "1室の施設では「△（残りわずか）」は構造上発生せず、必ず ○ か × になります。"
                        "使わない「部屋2〜4」の列は、右クリック→非表示 にしてください（削除は不要です）。"),
    ("★ファイルは施設ごとに分ける", "このファイルを施設の数だけコピーして「施設A_管理台帳.xlsx」「施設B_管理台帳.xlsx」"
                        "のように分けてください。1つのファイルに複数施設を混ぜると、定員が施設ごとに違うため"
                        "空き判定が成立せず、全ての集計に施設の絞り込みが必要になって必ず崩れます。"),
    ("SEC", "数え方の定義（全シート共通のルール）", ""),
    ("利用日数", "入所日から退所日までの暦日数。両端を含みます（4/1入所・4/3退所 = 3日）。"),
    ("宿泊数", "退所日 − 入所日（4/1入所・4/3退所 = 2泊）。"),
    ("日帰り（体験利用）", "入所日 = 退所日。1利用日・0泊。食事は昼のみが既定です。"),
    ("空室の判定", "「その日の夜に泊まるか」で見ます。★退所日の当日の夜は空き扱いです（次の方が入れます）。"),
    ("集計に入る予約", "確定・利用済み・仮予約・調整中は集計に入り、キャンセルだけが除外されます。"
                    "キャンセルは行を削除せず、状態を「キャンセル」に変えて記録として残してください。"),
    ("食事の既定", "日帰り=昼のみ ／ 初日=夕 ／ 中日=朝昼夕 ／ 最終日=朝。"
                 "例外だけ 04_食数表 の「調整」列に実数を入れると、そちらが優先されます。"),
    ("SEC", "セルの色の意味", ""),
    ("うすい黄色", "手で入力するセルです。"),
    ("うすい灰色（斜体）", "自動計算です。上書きすると壊れます。触らないでください。"),
    ("赤・オレンジ", "警告です。ダブルブッキング（⚠重複）と支給量オーバー（⚠支給量超過）を知らせます。"),
    ("SEC", "毎日の使い方", ""),
    ("予約が入ったら", "01_予約入力 の空いている行に、1滞在=1行で入れます。"
                  "S列・T列が赤くならないことだけ確認してください。"),
    ("「空いてますか」と聞かれたら", "03_空室照会 に日付を2つ入れるだけです。"),
    ("月の予定を見る", "02_月間表 の「対象月」を変えます。★この1か所を変えると、04_食数表 と 07_FAX空き表 の月も一緒に変わります。"),
    ("厨房へ食数を渡す", "04_食数表 の「確定」列（I〜K）を印刷して渡します。"),
    ("月末の報告・請求突合", "05_月次集計 に期間を入れると、区分別・利用者別が出ます。"),
    ("支給量の確認", "06_支給量管理 で残日数を確認してから予約を受けます。"),
    ("FAXで空き状況を送る", "07_FAX空き表 を印刷します。利用者名は載りません。"),
    ("SEC", "バックアップ", ""),
    ("月に一度はコピーを", "このファイル自体が台帳です。月初に「YYYYMM_管理台帳.xlsx」の名前でコピーを別フォルダに残してください。"),
]
r = 4
for item in rows:
    if item[0] == "SEC":
        c = ws.cell(row=r, column=1, value=item[1]); c.font = C_SUB; c.fill = F_SUB
        ws.cell(row=r, column=2).fill = F_SUB
        r += 1; continue
    a, b = item
    ca = ws.cell(row=r, column=1, value=a); ca.font = C_BOLD
    ca.alignment = Alignment(vertical="top", wrap_text=True)
    cb = ws.cell(row=r, column=2, value=b); cb.font = C_BODY
    cb.alignment = Alignment(vertical="top", wrap_text=True)
    ws.row_dimensions[r].height = 32
    r += 1


# ══════════════════════════════════════════════════════════════════
# M_施設設定
# ══════════════════════════════════════════════════════════════════
ws = sheet("M_施設設定")
title(ws, "施設設定", "黄色のセルを施設に合わせて書き換えてください。")
ws.column_dimensions["A"].width = 22
ws.column_dimensions["B"].width = 32
ws.column_dimensions["C"].width = 62
label_input(ws, 4, "施設名", "○○事業所", "帳票の見出しに使います")
label_input(ws, 5, "棟・フロア", "本棟 1F")
label_auto (ws, 6, "部屋数（定員）", f"=COUNTA({RROW})", "自動。M_部屋 の行数がそのまま定員になります（1部屋につき1名）")
label_input(ws, 7, "標準の入所時刻", "16:00", "予約入力の目安。実際の時刻は予約ごとに入れます")
label_input(ws, 8, "標準の退所時刻", "09:00")
label_input(ws, 9, "入力者名", "", "更新履歴に残す名前")
label_input(ws, 10, "郵便番号", "")
label_input(ws, 11, "住所", "", "FAX帳票の送信元として印字されます")
label_input(ws, 12, "電話番号", "")
label_input(ws, 13, "FAX番号", "")
label_input(ws, 14, "担当者名", "")
ADDR, TEL, FAXNO = f"{MSET}!$B$11", f"{MSET}!$B$12", f"{MSET}!$B$13"


# ══════════════════════════════════════════════════════════════════
# マスタ各種
# ══════════════════════════════════════════════════════════════════
def master_grid(ws, header_row, ncols, first, last, input_cols):
    for i in range(first, last + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=i, column=c)
            cell.border = BORDER; cell.font = C_BODY
            if c in input_cols: cell.fill = F_INPUT

ws = sheet("M_部屋")
title(ws, "部屋マスタ", "1行 = 1室。★ここの行数が定員になります。1室の施設は1行だけ、2室の施設は2行だけ残してください。")
head_row(ws, 3, ["部屋名", "備考"], [18, 60])
master_grid(ws, 3, 2, 4, 20, {1})
ws["A4"] = "部屋1"; ws["A5"] = "部屋2"
ws["B5"] = "1室のみの施設では、この行を削除してください"; ws["B5"].font = C_NOTE

ws = sheet("M_利用者")
title(ws, "利用者マスタ", "支給量＝受給者証の月あたり上限日数。空欄ならその方の残日数は管理しません。")
head_row(ws, 3, ["氏名", "支給量（日／月）", "備考"], [24, 16, 52])
master_grid(ws, 3, 3, 4, 60, {1, 2})
for i, (nm, alw, note) in enumerate([("利用者A", 7, "サンプルです。実際の氏名に書き換えてください"),
                                     ("利用者B", 5, ""), ("利用者C", None, "支給量を管理しない例")], start=4):
    ws.cell(row=i, column=1, value=nm)
    if alw is not None: ws.cell(row=i, column=2, value=alw)
    ws.cell(row=i, column=3, value=note).font = C_NOTE
ws["A3"].comment = Comment("氏名を書き換えると、01_予約入力 のプルダウンにも反映されます。\n"
                           "ただし既に入力済みの予約の氏名は自動では変わりません。", "設計メモ", height=90, width=320)

ws = sheet("M_区分")
title(ws, "区分マスタ", "「支給量に含める」に ○ を入れた区分だけ、06_支給量管理 の日数に数えます。")
head_row(ws, 3, ["区分", "支給量に含める（○／空欄）", "備考"], [16, 24, 52])
master_grid(ws, 3, 3, 4, 20, {1, 2})
for i, (nm, cnt) in enumerate([("SS", "○"), ("無料", "○"), ("契約", "○")], start=4):
    ws.cell(row=i, column=1, value=nm); ws.cell(row=i, column=2, value=cnt)

ws = sheet("M_送迎")
title(ws, "送迎ステータス マスタ", "")
head_row(ws, 3, ["送迎区分", "備考"], [18, 52])
master_grid(ws, 3, 2, 4, 15, {1})
for i, nm in enumerate(["あり", "なし", "調整中"], start=4):
    ws.cell(row=i, column=1, value=nm)

ws = sheet("M_状態")
title(ws, "予約状態マスタ", "★「キャンセル」の文字だけは集計が参照しているので、変更しないでください。")
head_row(ws, 3, ["予約状態", "意味"], [16, 68])
for i, (nm, note) in enumerate([("確定", "集計・稼働に入ります"),
                                ("仮予約", "集計に入りますが、月間表とFAX表では「仮」と表示します"),
                                ("調整中", "同上"),
                                ("利用済み", "実績。集計に入ります"),
                                ("キャンセル", "記録として残しますが、全ての集計と部屋の重複判定から外れます")], start=4):
    ws.cell(row=i, column=1, value=nm).font = C_BODY
    ws.cell(row=i, column=2, value=note).font = C_NOTE
    for c in (1, 2): ws.cell(row=i, column=c).border = BORDER

ws = sheet("M_FAX送付先")
title(ws, "FAX送付先マスタ", "番号を都度手入力せず一覧から選ぶことで、誤送信を防ぎます。")
head_row(ws, 3, ["事業所名", "担当者", "FAX番号", "電話番号", "備考"], [30, 16, 18, 18, 34])
master_grid(ws, 3, 5, 4, 24, {1, 2, 3, 4, 5})


# ── 定義名（プルダウン用。実データ件数ぶんに自動で伸縮） ──────────
def defname(name, ref): wb.defined_names.add(DefinedName(name, attr_text=ref))
defname("氏名一覧", f"OFFSET({MUSER}!$A$4,0,0,MAX(1,COUNTA({UROW})),1)")
defname("部屋一覧", f"OFFSET({MROOM}!$A$4,0,0,MAX(1,COUNTA({RROW})),1)")
defname("区分一覧", f"OFFSET({MCAT}!$A$4,0,0,MAX(1,COUNTA({CROW})),1)")
defname("送迎一覧", "OFFSET('M_送迎'!$A$4,0,0,MAX(1,COUNTA('M_送迎'!$A$4:$A$15)),1)")
defname("状態一覧", "'M_状態'!$A$4:$A$8")


# ══════════════════════════════════════════════════════════════════
# 01_予約入力
# ══════════════════════════════════════════════════════════════════
ws = sheet("01_予約入力")
title(ws, "予約入力（1滞在 = 1行）",
      "黄色の列だけ入力します。灰色は自動計算。S列・T列が赤くなったら、その行を直してください。")
headers = ["予約ID", "氏名", "区分", "部屋", "入所日", "入所時刻", "退所予定日", "退所時刻",
           "利用日数", "宿泊数", "予約状態", "送迎", "送迎方法", "送迎場所・時間帯",
           "外部サービス", "備考", "更新日", "更新者", "⚠重複・部屋チェック", "⚠支給量チェック", "支給量対象日数"]
widths  = [10, 16, 8, 10, 12, 10, 12, 10, 9, 8, 11, 9, 16, 20, 16, 30, 12, 12, 13, 15, 13]
head_row(ws, 4, headers, widths)
ws.freeze_panes = "B5"
ws.auto_filter.ref = f"A4:U{BK_LAST}"
AUTO_COLS = {"A", "I", "J", "S", "T", "U"}

for d in range(5, BK_LAST + 1):
    ws[f"A{d}"] = f'=IF($E{d}="","","SS-"&TEXT(ROW()-4,"0000"))'
    ws[f"I{d}"] = f'=IF(OR($E{d}="",$G{d}=""),"",$G{d}-$E{d}+1)'
    ws[f"J{d}"] = f'=IF(OR($E{d}="",$G{d}=""),"",$G{d}-$E{d})'
    ws[f"S{d}"] = (
        f'=IF($D{d}="","",IF(COUNTIF({RROW},$D{d})=0,"⚠部屋が未登録",'
        f'IF(OR($E{d}="",$G{d}="",$K{d}="キャンセル"),"",'
        f'IF(SUMPRODUCT(($D$5:$D${BK_LAST}=$D{d})*($E$5:$E${BK_LAST}<>"")*($G$5:$G${BK_LAST}<>"")'
        f'*($K$5:$K${BK_LAST}<>"キャンセル")*($E$5:$E${BK_LAST}<$G{d})*($G$5:$G${BK_LAST}>$E{d}))>1,'
        f'"⚠重複",""))))')
    ws[f"U{d}"] = (
        f'=IF(OR($E{d}="",$G{d}="",$K{d}="キャンセル"),0,'
        f'IF(IFERROR(INDEX({CCNT},MATCH($C{d},{CROW},0)),"○")="○",$G{d}-$E{d}+1,0))')
    ws[f"T{d}"] = (
        f'=IF(OR($B{d}="",$E{d}=""),"",'
        f'IF(SUMPRODUCT(({UROW}=$B{d})*({UALW}<>""))=0,"",'
        f'IF(SUMIFS($U$5:$U${BK_LAST},$B$5:$B${BK_LAST},$B{d},'
        f'$E$5:$E${BK_LAST},">="&DATE(YEAR($E{d}),MONTH($E{d}),1),'
        f'$E$5:$E${BK_LAST},"<="&EOMONTH($E{d},0))'
        f'>SUMIFS({UALW},{UROW},$B{d}),"⚠支給量超過","")))')
    for col in range(1, 22):
        L = get_column_letter(col); c = ws[f"{L}{d}"]
        c.border = BORDER
        c.font, c.fill = (C_AUTO, F_AUTO) if L in AUTO_COLS else (C_BODY, F_INPUT)
        if L in ("E", "G", "Q"): c.number_format = "yyyy/mm/dd"
        if L in ("F", "H"): c.number_format = "hh:mm"
        c.alignment = Alignment(vertical="center")

samples = [
    ("利用者A", "SS",  "部屋1", datetime.date(2026,4,6),  "16:00", datetime.date(2026,4,8),  "09:00", "確定",   "あり",  "施設車",  "自宅前 16:00", "", "サンプル行です。削除してお使いください"),
    ("利用者B", "契約", "部屋2", datetime.date(2026,4,6),  "16:30", datetime.date(2026,4,7),  "09:00", "確定",   "なし",  "",        "",             "", ""),
    ("利用者C", "無料", "部屋1", datetime.date(2026,4,11), "10:00", datetime.date(2026,4,11), "16:00", "確定",   "あり",  "家族送迎", "",            "", "日帰り（体験利用）の例：入所日と退所日が同じ"),
    ("利用者A", "SS",  "部屋1", datetime.date(2026,4,20), "16:00", datetime.date(2026,4,24), "09:00", "仮予約", "調整中", "",       "",             "", "仮予約の例"),
]
for i, s in enumerate(samples):
    d = 5 + i
    for col, val in zip("BCDEFGHKLMNOP", s):
        ws[f"{col}{d}"] = val

ws["S4"].comment = Comment("同じ部屋・同じ夜に2件以上入っていると「⚠重複」が出ます。\n"
                           "退所日の夜は空き扱いなので、前の方の退所日＝次の方の入所日は重複になりません。\n"
                           "M_部屋 に無い部屋名が入っていると「⚠部屋が未登録」が出ます。\n"
                           "（2室→1室に変更したのに、古い予約が部屋2のまま残っている場合など）",
                           "設計メモ", height=100, width=340)
ws["T4"].comment = Comment("M_利用者 に支給量が入っている方だけ判定します。\n"
                           "入所日の属する月で合計し、M_区分 で「○」の区分だけを数えます。",
                           "設計メモ", height=100, width=340)

for rng, name in [(f"B5:B{BK_LAST}", "氏名一覧"), (f"C5:C{BK_LAST}", "区分一覧"),
                  (f"D5:D{BK_LAST}", "部屋一覧"), (f"K5:K{BK_LAST}", "状態一覧"),
                  (f"L5:L{BK_LAST}", "送迎一覧")]:
    dv = DataValidation(type="list", formula1=f"={name}", allow_blank=True, showDropDown=False)
    dv.error = "一覧にない値です。マスタに登録してから選んでください。"
    dv.errorTitle = "入力できません"
    ws.add_data_validation(dv); dv.add(rng)

rng_all = f"A5:U{BK_LAST}"
ws.conditional_formatting.add(rng_all, FormulaRule(formula=['$K5="キャンセル"'],
    font=Font(name=FONT, size=10, color="999999", strike=True),
    fill=PatternFill("solid", fgColor="EFEFEF"), stopIfTrue=True))
ws.conditional_formatting.add(rng_all, FormulaRule(formula=['$S5<>""'], fill=F_WARN, stopIfTrue=True))
ws.conditional_formatting.add(rng_all, FormulaRule(formula=['$T5<>""'], fill=F_WARN2, stopIfTrue=True))
ws.conditional_formatting.add(rng_all, FormulaRule(formula=['OR($K5="仮予約",$K5="調整中")'], fill=F_TENT))


# ══════════════════════════════════════════════════════════════════
# 02_月間表  ―  「対象月」の唯一の入力元。04・07 もここを見る
# ══════════════════════════════════════════════════════════════════
ws = sheet("02_月間表")
title(ws, "月間表（誰が・どの部屋に泊まるか）",
      "★対象月はこのシートで指定します。04_食数表 と 07_FAX空き表 も同じ月を表示します。")
label_input(ws, 3, "対象月", MONTH0, "その月の日付ならどれでも可（1日でなくてもかまいません）", "yyyy/mm/dd")
label_auto (ws, 4, "月初", '=DATE(YEAR($B$3),MONTH($B$3),1)', "自動", "yyyy/mm/dd")
M1 = f"{MON}!$B$4"

hdr = ["日付", "曜日"] + [f"部屋{i}" for i in range(1, MAX_ROOMS + 1)] + \
      ["使用室数", "空室数", "仮予約", "空き記号", "日帰り"]
head_row(ws, 6, hdr, [12, 6] + [18] * MAX_ROOMS + [9, 9, 9, 11, 9])
for i in range(MAX_ROOMS):
    L = get_column_letter(3 + i)
    ws[f"{L}6"] = f'=IF(COUNTA({RROW})>={i+1},INDEX({RROW},{i+1}),"")'
USE_C = get_column_letter(3 + MAX_ROOMS)
VAC_C = get_column_letter(4 + MAX_ROOMS)
TEN_C = get_column_letter(5 + MAX_ROOMS)
SYM_C = get_column_letter(6 + MAX_ROOMS)
DAY_C = get_column_letter(7 + MAX_ROOMS)
FIRST_R, LAST_R = get_column_letter(3), get_column_letter(2 + MAX_ROOMS)

for i in range(31):
    d = 7 + i
    ws[f"A{d}"] = f'=IF(MONTH($B$4+{i})<>MONTH($B$4),"",$B$4+{i})'
    ws[f"B{d}"] = f'=IF($A{d}="","",MID("日月火水木金土",WEEKDAY($A{d}),1))'
    for j in range(MAX_ROOMS):
        L = get_column_letter(3 + j)
        idx = (f'SUMPRODUCT(MAX(({ROOM_}={L}$6)*({ST_}<>"キャンセル")*({IN_}<>"")'
               f'*({IN_}<=$A{d})*({OUT_}>$A{d})*(ROW({IN_})-4)))')
        ws[f"{L}{d}"] = (f'=IF(OR($A{d}="",{L}$6=""),"",'
                         f'IF({idx}=0,"",INDEX({NAME_},{idx})))')
    ws[f"{USE_C}{d}"] = (f'=IF($A{d}="","",SUMPRODUCT({VALID}'
                         f'*({IN_}<=$A{d})*({OUT_}>$A{d})))')
    ws[f"{VAC_C}{d}"] = f'=IF($A{d}="","",{CAP}-${USE_C}{d})'
    ws[f"{TEN_C}{d}"] = (f'=IF($A{d}="","",SUMPRODUCT((({ST_}="仮予約")+({ST_}="調整中"))'
                         f'*({IN_}<>"")*({IN_}<=$A{d})*({OUT_}>$A{d})))')
    ws[f"{SYM_C}{d}"] = (f'=IF($A{d}="","",IF(${TEN_C}{d}>0,"仮",'
                         f'IF(${VAC_C}{d}<=0,"×",IF(${VAC_C}{d}<{CAP},"△","○"))))')
    ws[f"{DAY_C}{d}"] = f'=IF($A{d}="","",SUMPRODUCT({VALID}*({IN_}=$A{d})*({OUT_}=$A{d})))'
    for col in range(1, 8 + MAX_ROOMS):
        c = ws.cell(row=d, column=col)
        c.border = BORDER; c.font = C_AUTO; c.fill = F_AUTO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{d}"].number_format = "yyyy/mm/dd"
    for j in range(MAX_ROOMS):
        ws.cell(row=d, column=3 + j).alignment = Alignment(horizontal="left", vertical="center")

def mark_cf(worksheet, rng):
    top = rng.split(":")[0]
    for mark, fill in [("×", F_WARN), ("△", F_WARN2), ("仮", F_TENT), ("○", F_OK)]:
        worksheet.conditional_formatting.add(rng, FormulaRule(formula=[f'{top}="{mark}"'], fill=fill))

mark_cf(ws, f"{SYM_C}7:{SYM_C}37")
ws.conditional_formatting.add("A7:B37", FormulaRule(formula=['AND($A7<>"",WEEKDAY($A7,2)>=6)'], fill=F_WKND))
ws[f"A39"] = "※ 「仮」は仮予約・調整中がその夜に含まれることを示します。※ 日帰り（0泊）は夜をふさがないため、使用室数には入りません。"
ws["A39"].font = C_NOTE
ws.print_area = f"A1:{DAY_C}37"
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True


# ══════════════════════════════════════════════════════════════════
# 03_空室照会
# ══════════════════════════════════════════════════════════════════
ws = sheet("03_空室照会")
title(ws, "空室照会 ―「この期間空いていますか」に即答する", "日付を2つ入れるだけです。月の指定は要りません。")
label_input(ws, 3, "開始日（入所日）", datetime.date(2026, 4, 6), "", "yyyy/mm/dd")
label_input(ws, 4, "終了日（退所予定日）", datetime.date(2026, 4, 9),
            "★退所日の夜は数えません（その夜は次の方が入れるため）", "yyyy/mm/dd")
label_auto (ws, 5, "泊数", '=IF(OR($B$3="",$B$4=""),"",MAX(0,$B$4-$B$3))',
            "0泊（同日）なら日帰り・体験利用です")
S_, E_ = "$B$3", "$B$4"
BOOKED = (f'((({OUT_}<={E_})*{OUT_}+({OUT_}>{E_})*{E_})'
          f'-(({IN_}>={S_})*{IN_}+({IN_}<{S_})*{S_}))')
head_row(ws, 7, ["部屋", "空いている夜数", "ふさがっている夜数", "判定"], [18, 16, 20, 30])
for i in range(MAX_ROOMS):
    d = 8 + i
    ws[f"A{d}"] = f'=IF(COUNTA({RROW})>={i+1},INDEX({RROW},{i+1}),"")'
    ws[f"C{d}"] = (f'=IF(OR($A{d}="",{E_}<={S_}),"",'
                   f'SUMPRODUCT({VALID}*({ROOM_}=$A{d})*({BOOKED}>0)*{BOOKED}))')
    ws[f"B{d}"] = f'=IF($C{d}="","",MAX(0,{E_}-{S_}-$C{d}))'
    ws[f"D{d}"] = (f'=IF($A{d}="","",IF({E_}<={S_},"日付を確認してください",'
                   f'IF($B{d}={E_}-{S_},"○ 全日空き",IF($B{d}=0,"× 空きなし","△ 一部空き"))))')
    for col in range(1, 5):
        c = ws.cell(row=d, column=col); c.border = BORDER; c.font = C_AUTO; c.fill = F_AUTO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{d}"].alignment = Alignment(horizontal="left", vertical="center")
    ws[f"D{d}"].alignment = Alignment(horizontal="left", vertical="center")
for mark, fill in [("○", F_OK), ("△", F_WARN2), ("×", F_WARN)]:
    ws.conditional_formatting.add("D8:D11", FormulaRule(formula=[f'LEFT($D8,1)="{mark}"'], fill=fill))
ws["A13"] = "※ 仮予約・調整中も「ふさがっている」として数えます。仮のまま押さえている部屋は 02_月間表 で「仮」と表示されます。"
ws["A14"] = "※ 日帰り（0泊）は夜をふさがないため、この照会には出ません。日中の重なりは 02_月間表 の「日帰り」列で確認してください。"
ws["A15"] = "※ 「一部空き」のときは、02_月間表 でどの日が埋まっているかを確認してください。"
for r_ in (13, 14, 15): ws[f"A{r_}"].font = C_NOTE


# ══════════════════════════════════════════════════════════════════
# 04_食数表（厨房用）
# ══════════════════════════════════════════════════════════════════
ws = sheet("04_食数表")
title(ws, "食数表（厨房用）",
      "既定は 日帰り=昼のみ／初日=夕／中日=朝昼夕／最終日=朝。例外だけ「調整」列に実数を入れると、そちらが優先されます。")
label_auto(ws, 3, "対象月", f"={MON}!$B$4", "★月の変更は 02_月間表 で行ってください", "yyyy/mm/dd")
hdr = ["日付", "曜日", "朝(自動)", "昼(自動)", "夕(自動)", "朝 調整", "昼 調整", "夕 調整",
       "朝 確定", "昼 確定", "夕 確定"] + [f"部屋{i} 宿泊者" for i in range(1, MAX_ROOMS + 1)]
head_row(ws, 5, hdr, [12, 6, 9, 9, 9, 9, 9, 9, 9, 9, 9] + [16] * MAX_ROOMS)
for i in range(MAX_ROOMS):
    ws[f"{get_column_letter(12+i)}5"] = f'={MON}!{get_column_letter(3+i)}$6&" 宿泊者"'
for i in range(31):
    d = 6 + i
    src = 7 + i
    ws[f"A{d}"] = f'=IF({MON}!$A{src}="","",{MON}!$A{src})'
    ws[f"B{d}"] = f'=IF($A{d}="","",MID("日月火水木金土",WEEKDAY($A{d}),1))'
    ws[f"C{d}"] = f'=IF($A{d}="","",SUMPRODUCT({VALID}*({IN_}<$A{d})*({OUT_}>=$A{d})))'
    ws[f"D{d}"] = (f'=IF($A{d}="","",SUMPRODUCT({VALID}*((({IN_}<$A{d})*({OUT_}>$A{d}))'
                   f'+(({IN_}=$A{d})*({OUT_}=$A{d})))))')
    ws[f"E{d}"] = f'=IF($A{d}="","",SUMPRODUCT({VALID}*({IN_}<=$A{d})*({OUT_}>$A{d})))'
    for k, (auto, adj) in enumerate([("C", "F"), ("D", "G"), ("E", "H")]):
        conf = get_column_letter(9 + k)
        ws[f"{conf}{d}"] = f'=IF($A{d}="","",IF(${adj}{d}="",${auto}{d},${adj}{d}))'
    for j in range(MAX_ROOMS):
        ws[f"{get_column_letter(12+j)}{d}"] = f'=IF($A{d}="","",{MON}!{get_column_letter(3+j)}{src})'
    for col in range(1, 12 + MAX_ROOMS):
        c = ws.cell(row=d, column=col); c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        if get_column_letter(col) in ("F", "G", "H"): c.font, c.fill = C_BODY, F_INPUT
        else: c.font, c.fill = C_AUTO, F_AUTO
    ws[f"A{d}"].number_format = "yyyy/mm/dd"
    for j in range(MAX_ROOMS):
        ws.cell(row=d, column=12 + j).alignment = Alignment(horizontal="left", vertical="center")
ws["A38"] = "月合計"; ws["A38"].font = C_SUB; ws["A38"].fill = F_SUB; ws["A38"].border = BORDER
for col in ("I", "J", "K"):
    c = ws[f"{col}38"]; c.value = f"=SUM({col}6:{col}36)"
    c.font = C_BOLD; c.fill = F_SUB; c.border = BORDER
    c.alignment = Alignment(horizontal="center")
ws["L38"] = "厨房へはこの「確定」列（I〜K）を渡してください"; ws["L38"].font = C_NOTE
ws.conditional_formatting.add("A6:B36", FormulaRule(formula=['AND($A6<>"",WEEKDAY($A6,2)>=6)'], fill=F_WKND))
ws.print_area = "A1:K38"
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True


# ══════════════════════════════════════════════════════════════════
# 05_月次集計
# ══════════════════════════════════════════════════════════════════
ws = sheet("05_月次集計")
title(ws, "期間集計（上司報告・請求突合用）",
      "期間を入れると区分別・利用者別が出ます。キャンセルは除外。期間にかかった滞在は、期間内に重なる日数だけ数えます。")
label_input(ws, 3, "開始日", datetime.date(2026, 4, 1), "", "yyyy/mm/dd")
label_input(ws, 4, "終了日", datetime.date(2026, 4, 30), "", "yyyy/mm/dd")
S_, E_ = "$B$3", "$B$4"
DAYS_TERM = (f'((({OUT_}<={E_})*{OUT_}+({OUT_}>{E_})*{E_})'
             f'-(({IN_}>={S_})*{IN_}+({IN_}<{S_})*{S_})+1)')
NIGHTS_TERM = (f'((({OUT_}<={E_}+1)*{OUT_}+({OUT_}>{E_}+1)*({E_}+1))'
               f'-(({IN_}>={S_})*{IN_}+({IN_}<{S_})*{S_}))')
OVL_D = f'({IN_}<={E_})*({OUT_}>={S_})'
OVL_N = f'({IN_}<={E_})*({OUT_}>{S_})'

ws["A6"] = "■ 区分別"; ws["A6"].font = C_SUB
head_row(ws, 7, ["区分", "利用者数", "利用回数", "延べ利用日数", "延べ宿泊数"], [16, 12, 12, 14, 14])
for i in range(CAT_ROWS):
    d = 8 + i
    ws[f"A{d}"] = f'=IF(COUNTA({CROW})>={i+1},INDEX({CROW},{i+1}),"")'
    ws[f"B{d}"] = (f'=IF($A{d}="","",SUMPRODUCT(--(COUNTIFS({NAME_},{UROW},{CAT_},$A{d},'
                   f'{IN_},"<="&{E_},{OUT_},">="&{S_},{ST_},"<>キャンセル")>0)*({UROW}<>"")))')
    ws[f"C{d}"] = (f'=IF($A{d}="","",COUNTIFS({CAT_},$A{d},{IN_},"<="&{E_},'
                   f'{OUT_},">="&{S_},{ST_},"<>キャンセル"))')
    ws[f"D{d}"] = f'=IF($A{d}="","",SUMPRODUCT({VALID}*({CAT_}=$A{d})*{OVL_D}*{DAYS_TERM}))'
    ws[f"E{d}"] = (f'=IF($A{d}="","",SUMPRODUCT({VALID}*({CAT_}=$A{d})*{OVL_N}'
                   f'*({NIGHTS_TERM}>0)*{NIGHTS_TERM}))')
    for col in range(1, 6):
        c = ws.cell(row=d, column=col); c.border = BORDER; c.font = C_AUTO; c.fill = F_AUTO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{d}"].alignment = Alignment(horizontal="left", vertical="center")
d = 8 + CAT_ROWS
ws[f"A{d}"] = "合計"
ws[f"B{d}"] = (f'=SUMPRODUCT(--(COUNTIFS({NAME_},{UROW},{IN_},"<="&{E_},'
               f'{OUT_},">="&{S_},{ST_},"<>キャンセル")>0)*({UROW}<>""))')
for col in ("C", "D", "E"):
    ws[f"{col}{d}"] = f"=SUM({col}8:{col}{d-1})"
for col in range(1, 6):
    c = ws.cell(row=d, column=col); c.border = BORDER; c.font = C_BOLD; c.fill = F_SUB
    c.alignment = Alignment(horizontal="center", vertical="center")
ws[f"A{d+1}"] = "※ 「利用者数」は実人数（同じ方が何回利用しても1人）。区分をまたぐ方がいる場合、区分別の合計と合計行は一致しません。"
ws[f"A{d+1}"].font = C_NOTE

ws["A21"] = "■ 利用者別"; ws["A21"].font = C_SUB
head_row(ws, 22, ["氏名", "利用回数", "延べ利用日数", "支給量対象日数（滞在全体）"], [22, 12, 14, 26])
for i in range(USER_ROWS):
    d = 23 + i
    ws[f"A{d}"] = f'=IF(COUNTA({UROW})>={i+1},INDEX({UROW},{i+1}),"")'
    ws[f"B{d}"] = (f'=IF($A{d}="","",COUNTIFS({NAME_},$A{d},{IN_},"<="&{E_},'
                   f'{OUT_},">="&{S_},{ST_},"<>キャンセル"))')
    ws[f"C{d}"] = f'=IF($A{d}="","",SUMPRODUCT({VALID}*({NAME_}=$A{d})*{OVL_D}*{DAYS_TERM}))'
    ws[f"D{d}"] = f'=IF($A{d}="","",SUMIFS({ALW_},{NAME_},$A{d},{IN_},">="&{S_},{IN_},"<="&{E_}))'
    for col in range(1, 5):
        c = ws.cell(row=d, column=col); c.border = BORDER; c.font = C_AUTO; c.fill = F_AUTO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{d}"].alignment = Alignment(horizontal="left", vertical="center")


# ══════════════════════════════════════════════════════════════════
# 06_支給量管理
# ══════════════════════════════════════════════════════════════════
ws = sheet("06_支給量管理")
title(ws, "支給量の残日数（受給者証の月あたり上限）",
      "入所日の属する月で集計します。M_区分 で「○」を付けた区分だけを数えます。")
label_input(ws, 3, "対象月", MONTH0, "その月の日付ならどれでも可", "yyyy/mm/dd")
head_row(ws, 5, ["氏名", "支給量（日／月）", "当月の対象日数", "残り", "状態"], [22, 16, 16, 12, 18])
for i in range(USER_ROWS):
    d = 6 + i
    ws[f"A{d}"] = f'=IF(COUNTA({UROW})>={i+1},INDEX({UROW},{i+1}),"")'
    ws[f"B{d}"] = (f'=IF($A{d}="","",IF(SUMPRODUCT(({UROW}=$A{d})*({UALW}<>""))=0,"",'
                   f'SUMIFS({UALW},{UROW},$A{d})))')
    ws[f"C{d}"] = (f'=IF($A{d}="","",SUMIFS({ALW_},{NAME_},$A{d},'
                   f'{IN_},">="&DATE(YEAR($B$3),MONTH($B$3),1),{IN_},"<="&EOMONTH($B$3,0)))')
    ws[f"D{d}"] = f'=IF(OR($A{d}="",$B{d}=""),"",$B{d}-$C{d})'
    ws[f"E{d}"] = (f'=IF($A{d}="","",IF($B{d}="","（管理しない）",'
                   f'IF($D{d}<0,"⚠ 超過",IF($D{d}=0,"上限ちょうど",IF($D{d}<=1,"残りわずか","OK")))))')
    for col in range(1, 6):
        c = ws.cell(row=d, column=col); c.border = BORDER; c.font = C_AUTO; c.fill = F_AUTO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{d}"].alignment = Alignment(horizontal="left", vertical="center")
last_u = 6 + USER_ROWS - 1
ws.conditional_formatting.add(f"A6:E{last_u}", FormulaRule(formula=['$E6="⚠ 超過"'], fill=F_WARN))
ws.conditional_formatting.add(f"A6:E{last_u}", FormulaRule(formula=['$E6="残りわずか"'], fill=F_WARN2))
ws[f"A{last_u+2}"] = "※ 月をまたぐ滞在は、入所日の属する月にまとめて計上します。月割りが必要な場合は 05_月次集計 を使ってください。"
ws[f"A{last_u+2}"].font = C_NOTE


# ══════════════════════════════════════════════════════════════════
# 07_FAX空き表（印刷用・氏名は載せない）
# ══════════════════════════════════════════════════════════════════
ws = sheet("07_FAX空き表")
ws["A1"] = f'={MSET}!$B$4&"　"&{MSET}!$B$5'; ws["A1"].font = C_TITLE
ws["A2"] = f'=TEXT({M1},"yyyy年m月")&"　短期入所 空き状況"'
ws["A2"].font = Font(name=FONT, size=11, bold=True)
ws["A3"] = f'={ADDR}&"　TEL "&{TEL}&"　FAX "&{FAXNO}'; ws["A3"].font = C_NOTE
ws["A4"] = f'="作成日："&TEXT(TODAY(),"yyyy年m月d日")&"　／　対象月は 02_月間表 で変更します"'
ws["A4"].font = C_NOTE
ws.column_dimensions["A"].width = 14
ws["A6"] = "日"; ws["A7"] = "曜日"; ws["A8"] = "空き状況"
for r_ in (6, 7, 8):
    c = ws[f"A{r_}"]; c.font = C_SUB; c.fill = F_SUB; c.border = BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
for i in range(31):
    L = get_column_letter(2 + i); src = 7 + i
    ws.column_dimensions[L].width = 3.8
    ws[f"{L}6"] = f'=IF({MON}!$A{src}="","",DAY({MON}!$A{src}))'
    ws[f"{L}7"] = f'=IF({MON}!$A{src}="","",{MON}!$B{src})'
    ws[f"{L}8"] = f'=IF({MON}!$A{src}="","",{MON}!${SYM_C}{src})'
    for r_ in (6, 7, 8):
        c = ws[f"{L}{r_}"]; c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.font = Font(name=FONT, size=11, bold=True) if r_ == 8 else C_BODY
LASTC = get_column_letter(32)
mark_cf(ws, f"B8:{LASTC}8")
ws.conditional_formatting.add(f"B7:{LASTC}7", FormulaRule(formula=['OR(B$7="土",B$7="日")'], fill=F_WKND))

ws["A10"] = "凡例"; ws["A10"].font = C_SUB
for i, (m, t) in enumerate([("○", "空室あり"), ("△", "残りわずか"), ("仮", "仮予約で調整中"), ("×", "満室")]):
    c = ws.cell(row=11, column=2 + i * 4, value=m)
    c.alignment = Alignment(horizontal="center"); c.font = Font(name=FONT, size=11, bold=True); c.border = BORDER
    ws.cell(row=11, column=3 + i * 4, value=t).font = C_NOTE
ws["A13"] = "・各日の欄は「その日の夜（宿泊）」の空き状況です。退所日の当日は、次の方がご利用いただけます。"
ws["A14"] = "・空き状況は日々変動します。ご予約前に必ずお電話でご確認ください。"
ws["A15"] = f'="・全"&{CAP}&"室。この用紙に利用者様のお名前は記載しておりません。"'
for r_ in (13, 14, 15): ws[f"A{r_}"].font = C_NOTE
ws["A17"] = "月内集計"; ws["A17"].font = C_SUB
for i, (m, lbl) in enumerate([("○", "空室あり"), ("△", "残りわずか"), ("仮", "仮予約"), ("×", "満室")]):
    col = 2 + i * 4
    c = ws.cell(row=18, column=col, value=f'=COUNTIF($B$8:${LASTC}$8,"{m}")&"日"')
    c.font = C_BODY; c.alignment = Alignment(horizontal="center")
    ws.cell(row=18, column=col + 1, value=lbl).font = C_NOTE
ws.print_area = f"A1:{LASTC}18"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

order = ["00_使い方", "01_予約入力", "02_月間表", "03_空室照会", "04_食数表",
         "05_月次集計", "06_支給量管理", "07_FAX空き表",
         "M_施設設定", "M_部屋", "M_利用者", "M_区分", "M_送迎", "M_状態", "M_FAX送付先"]
wb._sheets = [wb[n] for n in order]
wb.active = wb.index(wb["00_使い方"])
wb.save(OUT)
print("saved:", OUT)
