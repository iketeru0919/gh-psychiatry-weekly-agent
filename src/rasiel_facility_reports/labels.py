from __future__ import annotations

# LINEの営業活動報告チャットは、管理者ごとに項目の順番や言い回しが微妙に異なる
# (例:「営業予定時間」「短期稼働日数」「短期入所利用日数」はいずれも⑤番目に置かれるが
# 意味も表記も違う)。番号の位置ではなく、ラベル文字列そのものをキーに正規化する。
#
# キー: チャット本文に現れるラベル文字列
# 値: 集計で使う正規化済みフィールド名(canonical field)
LABEL_TO_FIELD: dict[str, str] = {
    "名前": "reporter_name",
    "拠点名": "facility_name",
    "当月目標入居件数": "target_move_ins",
    "目標入居件数": "target_move_ins",
    "当月短期稼働日数": "work_days",
    "短期稼働日数": "work_days",
    "短期入所利用日数": "work_days",
    "営業予定時間": "hours_planned",
    "営業実績時間": "hours_actual",
    "営業累計時間": "hours_cumulative",
    "営業実績件数累計": "cumulative_sales_count",
    "営業実施件数": "cumulative_sales_count",
    "営業件数": "cumulative_sales_count",
    "入居総数": "occupancy",
    "入居者数": "occupancy",
    "入居数": "occupancy",
    "営業内容": "sales_content",
    "預り金管理": "deposit_management",
    "カメラ録画確認": "camera_check",
    "カメラ作動確認": "camera_check",
    "カメラ動作確認": "camera_check",
    "カメラ認": "camera_check",
    "問合せ件数": "inquiries",
    "案件発掘": "leads_found",
    "見学": "tours",
    "体験": "trial_stays",
    "仮申込": "provisional_applications",
    "入居": "move_ins",
}

# startswithマッチで長い文字列を優先させるため、ラベル文字列を長い順に並べておく。
LABELS_BY_LENGTH_DESC: list[str] = sorted(LABEL_TO_FIELD, key=len, reverse=True)

# 出力CSVの列順、および人が読むための日本語見出し。
FIELD_DISPLAY_ORDER: list[tuple[str, str]] = [
    ("date", "日付"),
    ("time", "時刻"),
    ("facility_key", "拠点名(正規化)"),
    ("facility_raw", "拠点名(原文)"),
    ("reporter_name", "報告者"),
    ("am_name", "報告先(AM)"),
    ("occupancy", "入居者数"),
    ("target_move_ins", "目標入居件数"),
    ("work_days", "短期稼働日数"),
    ("hours_planned", "営業予定時間"),
    ("hours_actual", "営業実績時間"),
    ("hours_cumulative", "営業累計時間"),
    ("sales_content", "営業内容"),
    ("inquiries", "問合せ件数"),
    ("leads_found", "案件発掘"),
    ("tours", "見学"),
    ("trial_stays", "体験"),
    ("provisional_applications", "仮申込"),
    ("move_ins", "入居"),
    ("cumulative_sales_count", "営業実績件数累計"),
    ("deposit_management", "預り金管理"),
    ("camera_check", "カメラ録画確認"),
    ("next_day_plan", "翌日の予定"),
]

ALL_FIELD_KEYS: list[str] = [key for key, _ in FIELD_DISPLAY_ORDER]

# 拠点名の表記ゆれを吸収するために取り除く語句(会社名・敷地表現などのノイズ)。
FACILITY_NOISE_TOKENS: list[str] = [
    "株式会社ラシエル",
    "RASIEL",
    "Rasiel",
    "ラシエル",
    "・管理者",
    "管理者",
    "兼サビ管",
    "兼サビ官",
    "サビ管",
    "サビ官",
    "拠点",
    " ",
    "　",
]
