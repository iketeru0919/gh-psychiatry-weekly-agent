import json
import sys
import unicodedata
from pathlib import Path

import openpyxl


SHORT_STAY_MARKERS = ("短期", "短期入所")
LIFE_SUPPORT_MARKERS = ("生活援助日中",)
EXCLUDE_FILE_MARKERS = ("~$", "テンプレート", "旧", "コピー", "バックアップ", "test", "テスト", "短期")


def normalize_name(value):
    text = unicodedata.normalize("NFKC", str(value or ""))
    text = text.replace(" ", "").replace("　", "")
    replacements = {
        "髙": "高",
        "﨑": "崎",
        "邉": "邊",
        "辺": "邊",
        "渡邉": "渡邊",
        "栁": "柳",
        "澤": "沢",
        "菊地": "菊池",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return text


def as_number(value):
    if isinstance(value, (int, float)):
        return value
    return 0


def is_short_stay(service_name):
    return any(marker in str(service_name or "") for marker in SHORT_STAY_MARKERS)


def is_life_support(service_name):
    return any(marker in str(service_name or "") for marker in LIFE_SUPPORT_MARKERS)


def facility_folder_tokens(facility_name):
    normalized = normalize_name(facility_name)
    return [normalized, normalize_name(f"RASIEL{facility_name}"), normalize_name(f"ラシエル{facility_name}")]


def month_dir_candidates(year, month):
    return (
        f"{year}年{month}月提供分",
        f"{year}年{month:02d}月提供分",
        f"{month}月提供分",
        f"{month:02d}月提供分",
    )


def should_exclude_excel(path):
    name = path.name
    lower = name.lower()
    if not lower.endswith((".xlsx", ".xlsm")):
        return True
    return any(marker.lower() in lower for marker in EXCLUDE_FILE_MARKERS)


def find_provider_workbooks(provider_root, facility_name, year, month):
    root = Path(provider_root)
    records = []
    if not root.exists():
        return [], [
            {
                "facility_name": facility_name,
                "status": "未検出",
                "folder": str(root),
                "file": "",
                "note": "provider_root が存在しません。",
            }
        ]

    tokens = facility_folder_tokens(facility_name)
    facility_dirs = []
    for child in root.iterdir():
        if child.is_dir() and any(token in normalize_name(child.name) for token in tokens):
            facility_dirs.append(child)

    if not facility_dirs:
        return [], [
            {
                "facility_name": facility_name,
                "status": "未検出",
                "folder": str(root),
                "file": "",
                "note": "施設フォルダが見つかりません。",
            }
        ]

    year_text = f"{year}年"
    month_names = month_dir_candidates(year, month)
    excel_files = []
    for facility_dir in facility_dirs:
        month_dirs = []
        for subdir in facility_dir.rglob("*"):
            if not subdir.is_dir():
                continue
            normalized_name = normalize_name(subdir.name)
            if any(normalize_name(name) == normalized_name for name in month_names):
                month_dirs.append(subdir)
            elif year_text in str(subdir) and f"{month}月" in subdir.name and "提供分" in subdir.name:
                month_dirs.append(subdir)

        if not month_dirs:
            records.append(
                {
                    "facility_name": facility_name,
                    "status": "未検出",
                    "folder": str(facility_dir),
                    "file": "",
                    "note": f"{year}年{month}月提供分フォルダが見つかりません。",
                }
            )
            continue

        for month_dir in sorted(set(month_dirs)):
            candidates = [path for path in month_dir.rglob("*") if path.is_file() and not should_exclude_excel(path)]
            if not candidates:
                records.append(
                    {
                        "facility_name": facility_name,
                        "status": "未検出",
                        "folder": str(month_dir),
                        "file": "",
                        "note": "対象Excelが見つかりません。",
                    }
                )
            for candidate in sorted(candidates):
                excel_files.append(candidate)
                records.append(
                    {
                        "facility_name": facility_name,
                        "status": "候補",
                        "folder": str(month_dir),
                        "file": str(candidate),
                        "note": "読取候補として検出しました。",
                    }
                )

    return sorted(set(excel_files)), records


def find_header_map(row):
    return {value: idx for idx, value in enumerate(row) if value is not None}


def read_db_rows(workbook_path):
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    if "DB" not in wb.sheetnames:
        return {}
    result = {}
    ws = wb["DB"]
    for row_num, row in enumerate(ws.iter_rows(values_only=True), start=1):
        facility = row[2] if len(row) > 2 else None
        if not facility or facility in ("施設名", "担当AM"):
            continue
        result[str(facility)] = {
            "row": row_num,
            "facility_name": str(facility),
            "j_net_change": row[9] if len(row) > 9 else None,
            "k_net_change_reason": row[10] if len(row) > 10 else None,
            "m_billed_count": row[12] if len(row) > 12 else None,
            "n_provider_prev_count": row[13] if len(row) > 13 else None,
            "o_provider_count": row[14] if len(row) > 14 else None,
            "p_gap_label": row[15] if len(row) > 15 else None,
            "q_gap_reason": row[16] if len(row) > 16 else None,
            "t_billed_days": row[19] if len(row) > 19 else None,
            "aq_home_nursing_contracts": row[42] if len(row) > 42 else None,
        }
    return result


def read_billing(workbook_path, month_sheet, facility_name):
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    if month_sheet not in wb.sheetnames:
        return {"present": False, "users": {}, "resident_users": {}, "rows": []}
    ws = wb[month_sheet]
    header = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    header_map = find_header_map(header)
    users = {}
    rows = []
    for excel_row, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not row or row[0] != facility_name:
            continue
        record = {name: row[idx] if idx < len(row) else None for name, idx in header_map.items()}
        name = record.get("利用者名")
        service_name = record.get("サービス内容")
        count = as_number(record.get("回数"))
        normalized = normalize_name(name)
        user = users.setdefault(
            normalized,
            {
                "name": name,
                "normalized_name": normalized,
                "recipient_no": record.get("受給者証番号"),
                "life_support_days": 0,
                "short_stay_days": 0,
                "is_short_stay": False,
                "services": [],
                "excel_rows": [],
            },
        )
        user["excel_rows"].append(excel_row)
        user["services"].append(
            {
                "service_name": service_name,
                "unit": record.get("単位数"),
                "count": count,
                "service_units": record.get("サービス単位数"),
            }
        )
        if is_short_stay(service_name):
            user["is_short_stay"] = True
            user["short_stay_days"] += count
        if is_life_support(service_name):
            user["life_support_days"] += count
        rows.append(record)

    # 短期入所と生活援助日中を同月に併用する利用者がいるため、
    # 短期行の有無では除外せず、生活援助日中の回数があるかどうかで入居系と判定する。
    # 短期のみの利用者は life_support_days が 0 のためここで自然に除外される。
    resident_users = {
        key: value
        for key, value in users.items()
        if value["life_support_days"] > 0
    }
    return {
        "present": bool(rows),
        "all_user_count": len(users),
        "resident_user_count": len(resident_users),
        "life_support_days": sum(user["life_support_days"] for user in resident_users.values()),
        "users": users,
        "resident_users": resident_users,
        "rows": rows,
    }


def read_provider_file(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_vba=True)
    if "延べ日数" not in wb.sheetnames:
        return []
    ws = wb["延べ日数"]
    users = []
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        name = row[1] if len(row) > 1 else None
        if not isinstance(name, str) or not name.strip() or name.startswith("氏名"):
            continue
        days = row[3] if len(row) > 3 else None
        if not isinstance(days, (int, float)):
            continue
        users.append(
            {
                "name": name,
                "normalized_name": normalize_name(name),
                "category": row[2] if len(row) > 2 else None,
                "days": days,
                "single_days": row[5] if len(row) > 5 else None,
                "combined_days": row[6] if len(row) > 6 else None,
                "half_days": row[7] if len(row) > 7 else None,
                "direct_billing": row[8] if len(row) > 8 else None,
                "source_file": str(path),
                "source_row": row_num,
            }
        )
    return users


def read_provider(paths):
    merged = {}
    duplicates = []
    file_records = []
    for path in paths:
        try:
            users = read_provider_file(path)
        except Exception as exc:
            file_records.append(
                {
                    "status": "読取エラー",
                    "file": str(path),
                    "folder": str(Path(path).parent),
                    "note": str(exc),
                }
            )
            continue
        file_records.append(
            {
                "status": "読取OK" if users else "対象外",
                "file": str(path),
                "folder": str(Path(path).parent),
                "note": "延べ日数シートを読み取りました。" if users else "延べ日数シートまたは利用者行が見つかりません。",
            }
        )
        for user in users:
            key = user["normalized_name"]
            if key in merged:
                duplicates.append({"normalized_name": key, "name": user["name"], "source_file": user["source_file"]})
                merged[key]["days"] += as_number(user["days"])
                merged[key]["single_days"] = as_number(merged[key]["single_days"]) + as_number(user["single_days"])
                merged[key]["combined_days"] = as_number(merged[key]["combined_days"]) + as_number(user["combined_days"])
                merged[key]["half_days"] = as_number(merged[key]["half_days"]) + as_number(user["half_days"])
                merged[key]["source_file"] += "\n" + user["source_file"]
            else:
                merged[key] = user
    return {
        "users": merged,
        "sheet_user_count": len(merged),
        "active_user_count": sum(1 for user in merged.values() if as_number(user["days"]) > 0),
        "total_days": sum(as_number(user["days"]) for user in merged.values()),
        "duplicates": duplicates,
        "file_records": file_records,
    }


def build_reason(diff, summary, provider_user=None):
    item = diff["item"]
    billing = as_number(diff.get("billing_value"))
    provider = as_number(diff.get("provider_value"))
    user_name = diff.get("user_name") or "対象者"
    direct_billing = provider_user.get("direct_billing") if provider_user else None
    evidence_base = (
        f"請求延べ{summary['billing_life_support_days']}日、"
        f"提供延べ{summary['provider_total_days']}日。"
    )

    if item == "提供あり請求なし" and provider == 0:
        extra = " 直接請求フラグあり。" if direct_billing else ""
        return {
            "reason_confidence": "高",
            "reason_candidate": "延べ0日の利用者が提供ベース人数に含まれている可能性があります。",
            "evidence": f"{evidence_base}{user_name}さんは提供側にシートがありますが延べ0日です。{extra}",
            "check_point": "DB O列の提供人数に延べ0日利用者を含める定義か確認してください。",
            "db_q_draft": "理由：延べ0日の利用者を提供ベース人数に含めているため人数乖離。請求延べ日数と提供延べ日数は一致しており、請求漏れなし。対策：提供人数の計上定義を確認し、必要に応じて延べ0日利用者を除外。",
        }

    if item == "提供あり請求なし":
        return {
            "reason_confidence": "中",
            "reason_candidate": "未請求、月遅れ請求、受給者証更新待ちの可能性があります。",
            "evidence": f"{user_name}さんは提供側に{provider}日ありますが、請求側の生活援助日中系に見当たりません。",
            "check_point": "請求データの対象月、月遅れ予定、受給者証更新状況を確認してください。",
            "db_q_draft": f"理由：{user_name}さんの提供実績{provider}日が請求側に未反映のため乖離。対策：月遅れ・受給者証状況を確認し、必要に応じて翌月請求または請求修正。",
        }

    if item == "請求あり提供なし":
        return {
            "reason_confidence": "中",
            "reason_candidate": "過剰請求、提供記録漏れ、施設違いの可能性があります。",
            "evidence": f"{user_name}さんは請求側に{billing}日ありますが、提供側の延べ日数に見当たりません。",
            "check_point": "提供実績記録票の日別実績、対象施設、請求明細の施設名を確認してください。",
            "db_q_draft": f"理由：{user_name}さんの請求{billing}日が提供実績で確認できないため乖離。対策：提供記録と請求明細を確認し、過剰請求の場合は修正対応。",
        }

    if item == "利用者別延べ日数" and provider > billing:
        gap = provider - billing
        return {
            "reason_confidence": "中",
            "reason_candidate": "未請求、月遅れ、受給者証更新待ちの可能性があります。",
            "evidence": f"{user_name}さんは提供側が請求側より{gap}日多いです。",
            "check_point": "差異日数に該当する日別実績と、請求データの掲載月を確認してください。",
            "db_q_draft": f"理由：{user_name}さんの提供実績が請求より{gap}日多く、月遅れまたは未請求の可能性。対策：請求予定月を確認し、必要に応じて翌月請求。",
        }

    if item == "利用者別延べ日数" and billing > provider:
        gap = billing - provider
        return {
            "reason_confidence": "中",
            "reason_candidate": "過剰請求、提供記録漏れ、入院・外泊反映漏れの可能性があります。",
            "evidence": f"{user_name}さんは請求側が提供側より{gap}日多いです。",
            "check_point": "実績記録票の日別実績、入院・外泊、請求回数を確認してください。",
            "db_q_draft": f"理由：{user_name}さんの請求が提供実績より{gap}日多く、過剰請求または提供記録漏れの可能性。対策：実績記録票と請求明細を照合し、必要に応じて修正。",
        }

    return {
        "reason_confidence": "低",
        "reason_candidate": "差異の原因を追加確認してください。",
        "evidence": evidence_base,
        "check_point": "請求側・提供側の元データを確認してください。",
        "db_q_draft": "理由：差異原因を確認中。対策：請求データと提供実績を照合し、必要に応じて修正。",
    }


def attach_reason(diff, summary, provider_user=None):
    enriched = dict(diff)
    enriched.update(build_reason(diff, summary, provider_user))
    return enriched


def classify_facility(db_row, billing, provider, diffs):
    billing_days = as_number(billing.get("life_support_days"))
    provider_days = as_number(provider.get("total_days"))
    db_m = as_number(db_row.get("m_billed_count") if db_row else 0)
    db_o = as_number(db_row.get("o_provider_count") if db_row else 0)
    resident_count = as_number(billing.get("resident_user_count"))
    active_count = as_number(provider.get("active_user_count"))

    if any(item["severity"] == "重大" for item in diffs):
        return "重大"
    if any(item["severity"] == "要確認" for item in diffs) or billing_days != provider_days:
        return "要確認"
    if db_m != resident_count or db_o != active_count:
        return "注意"
    return "OK"


def summarize_main_reason(summary):
    if summary["judgement"] == "注意" and summary["billing_life_support_days"] == summary["provider_total_days"]:
        return "人数に乖離がありますが、請求延べ日数と提供延べ日数は一致しています。延べ0日利用者や人数定義を確認してください。"
    if summary["judgement"] == "OK":
        return "主要な人数・延べ日数は一致しています。"
    return "差異一覧の理由候補・根拠・確認ポイントを確認してください。"


def analyze_facility(config, db_rows):
    facility_name = config["facility_name"]
    db_row = db_rows.get(facility_name, {})
    billing = read_billing(CONFIG["billing_workbook"], CONFIG["target_month_label"], facility_name)
    discovery_records = []
    provider_paths = [Path(path) for path in config.get("provider_workbooks", [])]
    if not provider_paths and CONFIG.get("provider_root"):
        provider_paths, discovery_records = find_provider_workbooks(
            CONFIG["provider_root"],
            facility_name,
            int(CONFIG["target_year"]),
            int(CONFIG["target_month"]),
        )
    provider = read_provider(provider_paths)
    file_records = []
    for record in discovery_records:
        file_records.append(record)
    for record in provider.get("file_records", []):
        item = dict(record)
        item["facility_name"] = facility_name
        file_records.append(item)

    billing_users = billing.get("resident_users", {})
    provider_users = provider.get("users", {})
    summary = {
        "facility_name": facility_name,
        "db_row": db_row.get("row"),
        "db_m_billed_count": db_row.get("m_billed_count"),
        "db_o_provider_count": db_row.get("o_provider_count"),
        "db_p_gap_label": db_row.get("p_gap_label"),
        "db_q_gap_reason": db_row.get("q_gap_reason"),
        "db_t_billed_days": db_row.get("t_billed_days"),
        "billing_sheet_present": billing.get("present"),
        "billing_all_user_count": billing.get("all_user_count"),
        "billing_resident_user_count": billing.get("resident_user_count"),
        "billing_life_support_days": billing.get("life_support_days"),
        "provider_sheet_user_count": provider.get("sheet_user_count"),
        "provider_active_user_count": provider.get("active_user_count"),
        "provider_total_days": provider.get("total_days"),
    }

    diffs = []
    for key in sorted(set(provider_users) - set(billing_users)):
        user = provider_users[key]
        days = as_number(user.get("days"))
        diff = {
            "facility_name": facility_name,
            "user_name": user.get("name"),
            "item": "提供あり請求なし",
            "billing_value": 0,
            "provider_value": days,
            "difference": -days,
            "severity": "注意" if days == 0 else "重大",
        }
        diffs.append(attach_reason(diff, summary, user))

    for key in sorted(set(billing_users) - set(provider_users)):
        user = billing_users[key]
        days = as_number(user.get("life_support_days"))
        diff = {
            "facility_name": facility_name,
            "user_name": user.get("name"),
            "item": "請求あり提供なし",
            "billing_value": days,
            "provider_value": 0,
            "difference": days,
            "severity": "重大",
        }
        diffs.append(attach_reason(diff, summary))

    for key in sorted(set(billing_users) & set(provider_users)):
        billing_user = billing_users[key]
        provider_user = provider_users[key]
        billing_days = as_number(billing_user.get("life_support_days"))
        provider_days = as_number(provider_user.get("days"))
        if billing_days != provider_days:
            diff = {
                "facility_name": facility_name,
                "user_name": provider_user.get("name"),
                "item": "利用者別延べ日数",
                "billing_value": billing_days,
                "provider_value": provider_days,
                "difference": billing_days - provider_days,
                "severity": "要確認",
            }
            diffs.append(attach_reason(diff, summary, provider_user))

    summary["judgement"] = classify_facility(db_row, billing, provider, diffs)
    summary["main_reason"] = summarize_main_reason(summary)

    user_rows = []
    diff_by_name = {normalize_name(diff.get("user_name")): diff for diff in diffs}
    for key in sorted(set(billing_users) | set(provider_users)):
        billing_user = billing_users.get(key, {})
        provider_user = provider_users.get(key, {})
        billing_days = as_number(billing_user.get("life_support_days"))
        provider_days = as_number(provider_user.get("days"))
        related_diff = diff_by_name.get(key, {})
        user_rows.append(
            {
                "facility_name": facility_name,
                "user_name": provider_user.get("name") or billing_user.get("name"),
                "normalized_name": key,
                "billing_life_support_days": billing_days,
                "provider_days": provider_days,
                "difference": billing_days - provider_days,
                "direct_billing": provider_user.get("direct_billing"),
                "status": "OK" if billing_days == provider_days and billing_days > 0 else "確認",
                "reason_candidate": related_diff.get("reason_candidate"),
                "reason_confidence": related_diff.get("reason_confidence"),
                "evidence": related_diff.get("evidence"),
                "check_point": related_diff.get("check_point"),
                "db_q_draft": related_diff.get("db_q_draft"),
            }
        )

    return {"summary": summary, "diffs": diffs, "user_rows": user_rows, "file_records": file_records}


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: analyze_check.py CONFIG_JSON OUTPUT_JSON")
    global CONFIG
    config_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    with config_path.open("r", encoding="utf-8-sig") as f:
        CONFIG = json.load(f)
    if "target_month_label" not in CONFIG and "target_month" in CONFIG:
        CONFIG["target_month_label"] = f"{int(CONFIG['target_month'])}月"

    db_rows = read_db_rows(CONFIG["billing_workbook"])
    raw_facilities = CONFIG.get("facilities", [])
    facilities_config = [
        {"facility_name": item} if isinstance(item, str) else item
        for item in raw_facilities
    ]
    facilities = [analyze_facility(item, db_rows) for item in facilities_config]
    payload = {
        "report_title": CONFIG.get("report_title", "請求提供差異チェック"),
        "target_month_label": CONFIG.get("target_month_label"),
        "facilities": facilities,
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
