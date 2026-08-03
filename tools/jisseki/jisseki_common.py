"""実績記録表ツールの共通処理。

jisseki_cell_extract.py と jisseki_file_check.py の両方から使う。
判定ロジックをここ1か所にまとめることで、
「確認ツールはファイルなしと言うのに抽出ツールは読める」といった食い違いを防ぐ。
"""

from pathlib import Path
from datetime import datetime
import calendar
import re

# ===== パス設定 =====

BASE_DIR = Path(
    r"C:\Users\k-takayama\Bihonest Corp. Dropbox\17.ラシエル\00.ラシエル_General\00 実績記録表\実績記録表(請求用)\実績記録表　保存先"
)

OUTPUT_DIR = Path(r"C:\Users\k-takayama\python_work")

EXCEL_EXTENSIONS = [".xlsx", ".xlsm", ".xls"]

# ===== 走査対象から外すフォルダ =====

# 施設名の一部にたまたま含まれても除外されないよう、語として十分に長いものだけを並べる。
# （"テスト" だけだと「RASIELテスト施設」のような実在フォルダまで落ちる）
EXCLUDE_FOLDER_KEYWORDS = [
    "テストフォルダ",
    "テスト用",
    "使用禁止",
    "アーカイブ",
    "archive",
    "過去分",
    "バックアップ",
    "backup",
    "旧様式",
]

# ===== 利用者シートではないシート =====

# 完全一致で除外するシート名
EXCLUDE_SHEET_NAMES = {
    "実績記録票",
    "実績記録表",
    "延べ日数",
    "算定項目リスト",
    "重度障害者支援加算要件",
}

# 部分一致で除外するキーワード
EXCLUDE_SHEET_KEYWORDS = [
    "原本",
    "記入例",
    "集計",
    "一覧",
    "マスタ",
    "設定",
    "コピー",
    "様式",
    "フォーマット",
    "テンプレート",
    "空き",
    "空室",
    "予備",
    "要件",
]

# ===== 施設名の正規化 =====

# (判定関数, 正規化後の名前) の順に評価し、最初に一致したものを採用する。
FACILITY_NAME_RULES = [
    (lambda n: any(k in n for k in ["白浜町Ⅰ", "白浜町I", "白浜町１", "白浜町1"]), "白浜町Ⅰ"),
    (lambda n: any(k in n for k in ["白浜町Ⅱ", "白浜町II", "白浜町２", "白浜町2"]), "白浜町Ⅱ"),
    (lambda n: "美田園" in n and "南" in n, "RASIEL美田園 南"),
    (lambda n: "美田園" in n and "北" in n, "RASIEL美田園 北"),
    (lambda n: "おきつ" in n or "興津" in n, "おきつアリーナ"),
]


def normalize_facility_name(folder_name: str) -> str:
    for matches, normalized in FACILITY_NAME_RULES:
        if matches(folder_name):
            return normalized

    return folder_name


def is_excluded_folder(folder_name: str) -> bool:
    lowered = folder_name.lower()

    for keyword in EXCLUDE_FOLDER_KEYWORDS:
        if keyword.lower() in lowered:
            return True

    return False


def find_facility_folders(base_dir: Path):
    """施設フォルダの一覧と、除外したフォルダ名の一覧を返す。"""
    folders = []
    excluded = []

    for path in sorted(base_dir.iterdir()):
        if not path.is_dir():
            continue

        if is_excluded_folder(path.name):
            excluded.append(path.name)
            continue

        folders.append(path)

    return folders, excluded


def is_target_sheet(sheet_name: str) -> bool:
    """シート名だけで判定できる範囲で、利用者シートかどうかを返す。"""
    if not sheet_name:
        return False

    name = sheet_name.strip()

    if name in EXCLUDE_SHEET_NAMES:
        return False

    for keyword in EXCLUDE_SHEET_KEYWORDS:
        if keyword in name:
            return False

    return True


# ===== ファイル種別（1F / 2F / 短期）=====


# 全角の英数字を半角へ寄せるための変換表（1F / １F / 1Ｆ / １Ｆ をすべて同じ形にする）
_WIDTH_TABLE = str.maketrans(
    {
        **{chr(0xFF10 + i): chr(0x30 + i) for i in range(10)},   # ０-９ -> 0-9
        **{chr(0xFF21 + i): chr(0x41 + i) for i in range(26)},   # Ａ-Ｚ -> A-Z
        **{chr(0xFF41 + i): chr(0x61 + i) for i in range(26)},   # ａ-ｚ -> a-z
    }
)


def normalize_width(text: str) -> str:
    return text.translate(_WIDTH_TABLE)


def get_file_type(file_name: str):
    """ファイル名から 1F / 2F / 短期 を判定する。

    全角・半角の組み合わせ（1F, １F, 1Ｆ, １Ｆ）をすべて拾う。
    """
    if "短期" in file_name:
        return "短期"

    normalized = normalize_width(file_name).lower()

    if "1f" in normalized:
        return "1F"

    if "2f" in normalized:
        return "2F"

    return None


FILE_TYPES = ["1F", "2F", "短期"]


# ===== 年月判定 =====

_ZEN_TO_HAN = str.maketrans("０１２３４５６７８９", "0123456789")


def normalize_text_for_month_match(text: str) -> str:
    """月判定用に、全角数字を半角数字へ変換する。"""
    return text.translate(_ZEN_TO_HAN)


def to_reiwa_year(year: int) -> int:
    return year - 2018


def _relative_text(file_path: Path, base_dir: Path = None) -> str:
    """BASE_DIR より下の部分だけを判定対象にする。

    BASE_DIR 自体に含まれる数字（"17.ラシエル" や "00 実績記録表" など）を
    年月として誤検出しないようにするため。
    """
    base_dir = base_dir or BASE_DIR

    try:
        text = str(file_path.relative_to(base_dir))
    except ValueError:
        text = str(file_path)

    return normalize_text_for_month_match(text)


def is_target_month_file(file_path: Path, target_year: int, target_month: int,
                         base_dir: Path = None) -> bool:
    """パス・ファイル名・フォルダ名から対象年月のファイルかを判定する。

    前方一致による誤検出（1月指定で 10/11/12月、2月指定で 12月を拾う）を
    防ぐため、単純な in ではなく数字境界つきの正規表現で判定する。
    """
    text = _relative_text(file_path, base_dir)

    reiwa = to_reiwa_year(target_year)
    month_pattern = f"0?{target_month}"
    month_2 = f"{target_month:02d}"

    patterns = [
        rf"{target_year}年{month_pattern}月",
        # 「R8.4」「令和8年4月」。年号と月の間に区切りか「年」が必要（R84 は拾わない）
        rf"(?:R|令和){reiwa}(?:年|[.\-_/／]){month_pattern}(?!\d)",
        rf"(?:R|令和){reiwa}(?:年|[.\-_/／])?{month_pattern}月",
        rf"(?<!\d){target_year}[.\-_/／]{month_pattern}(?!\d)",
        rf"(?<!\d){target_year}{month_2}(?!\d)",
    ]

    for pattern in patterns:
        if re.search(pattern, text):
            return True

    # 「年フォルダ」＋「月フォルダ」のように分かれている場合への対応
    year_exists = re.search(
        rf"(?<!\d){target_year}(?!\d)|(?:R|令和){reiwa}(?!\d)", text
    )

    month_exists = re.search(
        rf"(?<!\d){month_pattern}月|[\\/_](?:0?{target_month})[\\/_]", text
    )

    return bool(year_exists and month_exists)


_MONTH_PATTERNS = [
    (re.compile(r"(?<!\d)(\d{4})年(0?[1-9]|1[0-2])月"), "seireki"),
    (re.compile(r"(?:R|令和)(\d{1,2})(?:年|[.\-_/／])(0?[1-9]|1[0-2])月?(?!\d)"), "reiwa"),
    (re.compile(r"(?<!\d)(\d{4})[.\-_/／](0?[1-9]|1[0-2])(?!\d)"), "seireki"),
    (re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)"), "seireki"),
]


def extract_year_month_from_path(file_path: Path, base_dir: Path = None):
    """パスから (年, 月) を推定する。判定できない場合は None を返す。

    ファイル名の方がフォルダ名より確からしいため、
    一番右（パス末尾に近い）で見つかった表記を採用する。
    """
    text = _relative_text(file_path, base_dir)

    best = None
    best_pos = -1

    for pattern, era in _MONTH_PATTERNS:
        for match in pattern.finditer(text):
            if match.start() < best_pos:
                continue

            year = int(match.group(1))
            month = int(match.group(2))

            if era == "reiwa":
                year = year + 2018

            if not (1 <= month <= 12):
                continue

            if not (2000 <= year <= 2100):
                continue

            best = (year, month)
            best_pos = match.start()

    return best


def days_in_month(year: int, month: int) -> int:
    return calendar.monthrange(year, month)[1]


# ===== 対象ファイルの収集 =====


def collect_target_files(mode: str, target_year=None, target_month=None,
                         base_dir: Path = None):
    """施設フォルダごとに 1F / 2F / 短期 の対象ファイルを1つずつ選ぶ。

    戻り値: (target_files, errors, excluded_folders)

    target_files の各要素:
        施設名 / 元フォルダ名 / 区分 / ファイル / 最新更新日時 /
        検出年月 / 候補ファイル数 / 候補ファイル一覧
    """
    base_dir = base_dir or BASE_DIR

    target_files = []
    errors = []

    if not base_dir.exists():
        errors.append({
            "施設名": "",
            "区分": "",
            "利用者名": "",
            "内容": f"保存先フォルダが見つかりません: {base_dir}",
        })
        return target_files, errors, []

    facility_folders, excluded_folders = find_facility_folders(base_dir)

    print(f"施設フォルダ数: {len(facility_folders)}（除外: {len(excluded_folders)}）")

    normalized_to_folders = {}

    for facility_folder in facility_folders:
        facility_name = normalize_facility_name(facility_folder.name)
        normalized_to_folders.setdefault(facility_name, []).append(facility_folder.name)

        files_by_type = {file_type: [] for file_type in FILE_TYPES}

        for file in facility_folder.rglob("*"):
            if not file.is_file():
                continue

            if file.suffix.lower() not in EXCEL_EXTENSIONS:
                continue

            if file.name.startswith("~$"):
                continue

            if any(is_excluded_folder(part) for part in file.relative_to(facility_folder).parts[:-1]):
                continue

            file_type = get_file_type(file.name)

            if not file_type:
                continue

            if mode == "month":
                if not is_target_month_file(file, target_year, target_month, base_dir):
                    continue

            files_by_type[file_type].append(file)

        for file_type in FILE_TYPES:
            candidates = files_by_type[file_type]

            if not candidates:
                if mode == "latest":
                    message = "対象ファイルなし"
                else:
                    message = f"{target_year}年{target_month}月の対象ファイルなし"

                errors.append({
                    "施設名": facility_name,
                    "区分": file_type,
                    "利用者名": "",
                    "内容": message,
                })
                continue

            latest_file = _select_file(candidates, mode, target_year, target_month, base_dir)
            latest_time = datetime.fromtimestamp(latest_file.stat().st_mtime)

            if len(candidates) > 1:
                names = " / ".join(sorted(p.name for p in candidates))
                errors.append({
                    "施設名": facility_name,
                    "区分": file_type,
                    "利用者名": "",
                    "内容": (
                        f"候補ファイルが{len(candidates)}件あります。"
                        f"「{latest_file.name}」を採用しました（要確認）: {names}"
                    ),
                })

            target_files.append({
                "施設名": facility_name,
                "元フォルダ名": facility_folder.name,
                "区分": file_type,
                "ファイル": latest_file,
                "最新更新日時": latest_time.strftime("%Y-%m-%d %H:%M:%S"),
                "検出年月": extract_year_month_from_path(latest_file, base_dir),
                "候補ファイル数": len(candidates),
            })

    for normalized, folders in normalized_to_folders.items():
        if len(folders) > 1:
            errors.append({
                "施設名": normalized,
                "区分": "",
                "利用者名": "",
                "内容": f"複数フォルダが同じ施設名に正規化されています（合算されます）: {' / '.join(folders)}",
            })

    return target_files, errors, excluded_folders


def _select_file(candidates, mode, target_year, target_month, base_dir):
    """採用する1ファイルを決める。

    更新日時だけで選ぶと、Dropbox の同期で mtime が書き換わったときに
    古い月のファイルを掴む。月指定モードではファイル名の年月を優先する。
    """
    if mode == "month":
        exact = [
            p for p in candidates
            if extract_year_month_from_path(p, base_dir) == (target_year, target_month)
        ]

        if exact:
            candidates = exact

    return max(candidates, key=lambda p: p.stat().st_mtime)
