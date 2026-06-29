import streamlit as st
import pandas as pd
import json
import os
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="エリアディレクター 業務チェックツール",
    page_icon="🏢",
    layout="wide",
)

DATA_DIR = Path(__file__).parent / "data" / "facilities"
EXCEL_DIR = Path(__file__).parent / "data"

# ── 閾値設定 ──────────────────────────────────────────────
THRESHOLDS = {
    "occupancy_rate":        {"red": 85.0,  "yellow": 90.0,  "higher_is_better": True},
    "ss_rate":               {"red": 20.0,  "yellow": 30.0,  "higher_is_better": True},
    "night_shift_calc_rate": {"red": 95.0,  "yellow": 98.0,  "higher_is_better": True},
    "派遣利用率":             {"red": 25.0,  "yellow": 15.0,  "higher_is_better": False},
    "タイミー応募率":          {"red": 5.0,   "yellow": 10.0,  "higher_is_better": True},
    "重大事故件数":            {"red": 1,     "yellow": 0,     "higher_is_better": False},
    "苦情件数":               {"red": 1,     "yellow": 0,     "higher_is_better": False},
}


def load_all_facilities() -> list[dict]:
    """
    データソースの優先順位:
    1. data/ 直下に .xlsx ファイルがあればExcelから読み込む
    2. なければ data/facilities/ のJSONから読み込む（開発・デモ用）
    """
    excel_files = sorted(EXCEL_DIR.glob("*.xlsx"))
    if excel_files:
        from excel_reader import load_from_excel
        all_facilities = []
        for xf in excel_files:
            all_facilities.extend(load_from_excel(xf))
        return all_facilities

    # fallback: JSONファイル
    facilities = []
    for path in sorted(DATA_DIR.glob("*.json")):
        with open(path, encoding="utf-8") as f:
            facilities.append(json.load(f))
    return facilities


def get_latest_month(facility: dict) -> dict | None:
    monthly = facility.get("monthly", [])
    if not monthly:
        return None
    return sorted(monthly, key=lambda m: m["month"])[-1]


def traffic_light(value, key: str) -> str:
    if key not in THRESHOLDS:
        return ""
    t = THRESHOLDS[key]
    if t["higher_is_better"]:
        if value < t["red"]:
            return "🔴"
        elif value < t["yellow"]:
            return "🟡"
        return "🟢"
    else:
        if value >= t["red"]:
            return "🔴"
        elif value >= t["yellow"]:
            return "🟡"
        return "🟢"


def build_dashboard_df(facilities: list[dict]) -> pd.DataFrame:
    rows = []
    for f in facilities:
        m = get_latest_month(f)
        if not m:
            continue
        capacity = f.get("capacity", 1)
        派遣利用率 = round(m.get("派遣合計時間", 0) / max(m.get("実配置TC", 1) * 160, 1) * 100, 1)
        rows.append({
            "施設名":         f["facility_name"],
            "グループ":       f["group_company"],
            "定員":           capacity,
            "入居者数":       m.get("residents", 0),
            "稼働率(%)":      m.get("occupancy_rate", 0),
            "SS稼働率(%)":    m.get("ss_rate", 0),
            "夜勤加配率(%)":  m.get("night_shift_calc_rate", 0),
            "外部SV率(%)":    m.get("external_sv_rate", 0),
            "退去者数":       m.get("move_outs", 0),
            "算定単位数":     m.get("total_billing_units", 0),
            "派遣時間":       m.get("派遣合計時間", 0),
            "派遣利用率(%)":  派遣利用率,
            "タイミー日勤h":  m.get("タイミー日勤時間", 0),
            "タイミー夜勤h":  m.get("タイミー夜勤時間", 0),
            "タイミー応募率(%)": m.get("タイミー応募率", 0),
            "タイミー採用率(%)": m.get("タイミー採用率", 0),
            "苦情件数":       m.get("苦情件数", 0),
            "重大事故件数":   m.get("重大事故件数", 0),
            "ヒヤリハット件数": m.get("ヒヤリハット件数", 0),
            "応募者数":       m.get("応募者数", 0),
            "入職者数":       m.get("入職者数", 0),
            "退職者数":       m.get("退職者数", 0),
            "GH問合せ件数":   m.get("gh_inquiries", 0),
            "SS問合せ件数":   m.get("ss_inquiries", 0),
            "営業訪問件数":   m.get("sales_visits", 0),
            "見学体験者数":   m.get("tours", 0),
            "_month":         m.get("month", ""),
        })
    return pd.DataFrame(rows)


def color_cell(val, col: str) -> str:
    key_map = {
        "稼働率(%)":        "occupancy_rate",
        "SS稼働率(%)":      "ss_rate",
        "夜勤加配率(%)":    "night_shift_calc_rate",
        "派遣利用率(%)":    "派遣利用率",
        "タイミー応募率(%)": "タイミー応募率",
        "重大事故件数":     "重大事故件数",
        "苦情件数":         "苦情件数",
    }
    key = key_map.get(col)
    if not key or key not in THRESHOLDS:
        return ""
    t = THRESHOLDS[key]
    if t["higher_is_better"]:
        if val < t["red"]:
            return "background-color: #ffcccc"
        elif val < t["yellow"]:
            return "background-color: #fff3cc"
        return "background-color: #ccffcc"
    else:
        if val >= t["red"]:
            return "background-color: #ffcccc"
        elif val >= t["yellow"]:
            return "background-color: #fff3cc"
        return "background-color: #ccffcc"


# ── サイドバー ────────────────────────────────────────────
st.sidebar.title("🏢 ADチェックツール")

# Excelアップロード
st.sidebar.subheader("📂 データ読み込み")
uploaded = st.sidebar.file_uploader(
    "Excelファイルをアップロード",
    type=["xlsx"],
    help="施設ごとにシートが分かれているExcelファイルをアップロードしてください。",
)

if uploaded:
    import tempfile, shutil
    from excel_reader import load_from_excel
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        shutil.copyfileobj(uploaded, tmp)
        tmp_path = tmp.name
    try:
        facilities = load_from_excel(tmp_path)
        st.sidebar.success(f"✅ {len(facilities)}施設を読み込みました")
    except Exception as e:
        st.sidebar.error(f"読み込みエラー: {e}")
        facilities = load_all_facilities()
else:
    facilities = load_all_facilities()
    excel_files = sorted(EXCEL_DIR.glob("*.xlsx"))
    if excel_files:
        st.sidebar.info(f"📄 {excel_files[0].name} を自動読み込み中")
    else:
        st.sidebar.caption("デモデータ（JSONサンプル）を表示中")

page = st.sidebar.radio(
    "ページ選択",
    ["📊 横断ダッシュボード", "🔍 施設別詳細", "✅ 月次チェックリスト", "📈 トレンド分析"],
)

if not facilities:
    st.error("施設データが見つかりません。data/facilities/ にJSONファイルを配置してください。")
    st.stop()

latest_month = get_latest_month(facilities[0]).get("month", "") if facilities else ""
st.sidebar.markdown(f"**最新データ月:** {latest_month}")
st.sidebar.markdown(f"**施設数:** {len(facilities)}")

# ════════════════════════════════════════════════════════════
# PAGE 1: 横断ダッシュボード
# ════════════════════════════════════════════════════════════
if page == "📊 横断ダッシュボード":
    st.title("📊 横断KPIダッシュボード")
    st.caption(f"対象月: {latest_month}　　🔴=要対応　🟡=注意　🟢=良好")

    df = build_dashboard_df(facilities)

    # ── サマリーカード ──────────────────────────────────────
    col1, col2, col3, col4, col5 = st.columns(5)
    avg_occ   = df["稼働率(%)"].mean()
    avg_ss    = df["SS稼働率(%)"].mean()
    red_occ   = (df["稼働率(%)"] < 85).sum()
    total_acc = df["重大事故件数"].sum()
    total_dep = df["派遣時間"].sum()

    col1.metric("全施設平均稼働率",   f"{avg_occ:.1f}%",  delta=None)
    col2.metric("全施設平均SS稼働率", f"{avg_ss:.1f}%",   delta=None)
    col3.metric("🔴稼働率<85%施設",  f"{red_occ}施設",   delta=None)
    col4.metric("重大事故（当月計）", f"{total_acc}件",    delta=None)
    col5.metric("派遣合計時間",       f"{total_dep:,.0f}h", delta=None)

    st.divider()

    # ── ブロックA: 稼働・収益 ─────────────────────────────
    st.subheader("A｜稼働・収益系")
    cols_a = ["施設名", "グループ", "定員", "入居者数", "稼働率(%)", "SS稼働率(%)", "夜勤加配率(%)", "退去者数", "算定単位数"]
    style_cols_a = ["稼働率(%)", "SS稼働率(%)", "夜勤加配率(%)"]

    def apply_style(df_in, cols):
        return df_in.style.apply(
            lambda col: [color_cell(v, col.name) if col.name in cols else "" for v in col],
            axis=0
        )

    st.dataframe(
        apply_style(df[cols_a], style_cols_a),
        use_container_width=True,
        hide_index=True,
    )

    # ── ブロックB: 人員コスト ────────────────────────────
    st.subheader("B｜人員コスト系")
    cols_b = ["施設名", "派遣時間", "派遣利用率(%)", "タイミー日勤h", "タイミー夜勤h", "タイミー応募率(%)", "タイミー採用率(%)"]
    style_cols_b = ["派遣利用率(%)", "タイミー応募率(%)"]
    st.dataframe(
        apply_style(df[cols_b], style_cols_b),
        use_container_width=True,
        hide_index=True,
    )

    # ── ブロックC: 採用・定着 ────────────────────────────
    st.subheader("C｜採用・定着系")
    cols_c = ["施設名", "応募者数", "入職者数", "退職者数"]
    st.dataframe(df[cols_c], use_container_width=True, hide_index=True)

    # ── ブロックD: リスク・品質 ──────────────────────────
    st.subheader("D｜リスク・品質系")
    st.caption("⚠️ ヒヤリハット0件は『未報告』の可能性があります。管理者に確認してください。")
    cols_d = ["施設名", "苦情件数", "重大事故件数", "ヒヤリハット件数"]
    style_cols_d = ["苦情件数", "重大事故件数"]

    def highlight_zero_hiyari(row):
        styles = [""] * len(row)
        if "ヒヤリハット件数" in row.index:
            idx = row.index.get_loc("ヒヤリハット件数")
            if row["ヒヤリハット件数"] == 0:
                styles[idx] = "background-color: #fff3cc"
        return styles

    styled_d = df[cols_d].style.apply(
        lambda col: [color_cell(v, col.name) if col.name in style_cols_d else "" for v in col],
        axis=0
    ).apply(highlight_zero_hiyari, axis=1)
    st.dataframe(styled_d, use_container_width=True, hide_index=True)

    # ── ブロックE: ブランディング・営業 ─────────────────
    st.subheader("E｜ブランディング・営業系")
    st.caption("⚠️ 営業訪問0件の施設は管理者への確認が必要です。")
    cols_e = ["施設名", "GH問合せ件数", "SS問合せ件数", "営業訪問件数", "見学体験者数"]

    def highlight_zero_sales(row):
        styles = [""] * len(row)
        if "営業訪問件数" in row.index:
            idx = row.index.get_loc("営業訪問件数")
            if row["営業訪問件数"] == 0:
                styles[idx] = "background-color: #ffcccc"
        return styles

    st.dataframe(
        df[cols_e].style.apply(highlight_zero_sales, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    # ── 稼働率バーチャート ───────────────────────────────
    st.divider()
    st.subheader("稼働率・SS稼働率 比較")
    fig = go.Figure()
    fig.add_bar(name="稼働率(%)",    x=df["施設名"], y=df["稼働率(%)"],    marker_color="#4C72B0")
    fig.add_bar(name="SS稼働率(%)",  x=df["施設名"], y=df["SS稼働率(%)"],  marker_color="#55A868")
    fig.add_hline(y=85, line_dash="dash", line_color="red",    annotation_text="稼働率 警戒ライン 85%")
    fig.add_hline(y=20, line_dash="dot",  line_color="orange", annotation_text="SS 警戒ライン 20%")
    fig.update_layout(barmode="group", height=350, margin=dict(t=30, b=10))
    st.plotly_chart(fig, use_container_width=True)


# ════════════════════════════════════════════════════════════
# PAGE 2: 施設別詳細
# ════════════════════════════════════════════════════════════
elif page == "🔍 施設別詳細":
    st.title("🔍 施設別詳細")
    names = [f["facility_name"] for f in facilities]
    selected = st.selectbox("施設を選択", names)
    facility = next(f for f in facilities if f["facility_name"] == selected)
    monthly = sorted(facility.get("monthly", []), key=lambda m: m["month"])

    if not monthly:
        st.warning("月次データがありません。")
        st.stop()

    m = monthly[-1]
    st.subheader(f"{selected}　（{facility['group_company']}）　定員 {facility['capacity']}名")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("稼働率",    f"{m.get('occupancy_rate', 0):.1f}%")
    c2.metric("SS稼働率",  f"{m.get('ss_rate', 0):.1f}%")
    c3.metric("夜勤加配率", f"{m.get('night_shift_calc_rate', 0):.1f}%")
    c4.metric("算定単位数", f"{m.get('total_billing_units', 0):,}")

    st.divider()

    # 月次推移グラフ
    mdf = pd.DataFrame(monthly)
    st.subheader("月次推移")
    tab1, tab2, tab3 = st.tabs(["稼働・SS", "人員コスト", "リスク・採用"])

    with tab1:
        fig = px.line(mdf, x="month", y=["occupancy_rate", "ss_rate", "night_shift_calc_rate"],
                      labels={"value": "%", "month": "月", "variable": "指標"},
                      title="稼働率・SS稼働率・夜勤加配率 推移")
        fig.add_hline(y=85, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        fig2 = px.bar(mdf, x="month",
                      y=["派遣合計時間", "タイミー日勤時間", "タイミー夜勤時間"],
                      title="派遣・タイミー時間 推移",
                      labels={"value": "時間", "month": "月", "variable": "種別"},
                      barmode="stack")
        st.plotly_chart(fig2, use_container_width=True)

    with tab3:
        fig3 = go.Figure()
        fig3.add_bar(name="重大事故", x=mdf["month"], y=mdf["重大事故件数"], marker_color="red")
        fig3.add_bar(name="苦情",     x=mdf["month"], y=mdf["苦情件数"],     marker_color="orange")
        fig3.add_bar(name="ヒヤリハット", x=mdf["month"], y=mdf["ヒヤリハット件数"], marker_color="blue")
        fig3.add_scatter(name="応募者数", x=mdf["month"], y=mdf["応募者数"], mode="lines+markers")
        fig3.add_scatter(name="入職者数", x=mdf["month"], y=mdf["入職者数"], mode="lines+markers")
        fig3.update_layout(barmode="group", title="リスク・採用 推移", height=350)
        st.plotly_chart(fig3, use_container_width=True)

    # 週次データ
    weekly = facility.get("weekly_current", [])
    if weekly:
        st.subheader("当月 週次進捗")
        wdf = pd.DataFrame(weekly)
        st.dataframe(wdf, use_container_width=True, hide_index=True)

    # 全月次テーブル
    with st.expander("月次データ詳細（全項目）"):
        st.dataframe(mdf, use_container_width=True)


# ════════════════════════════════════════════════════════════
# PAGE 3: 月次チェックリスト
# ════════════════════════════════════════════════════════════
elif page == "✅ 月次チェックリスト":
    st.title("✅ ADルーティン 月次チェックリスト")
    st.caption("チェック完了状況はセッション内で保持されます（リロードでリセット）")

    df = build_dashboard_df(facilities)
    red_occ_list  = df[df["稼働率(%)"]    < 85]["施設名"].tolist()
    red_ss_list   = df[df["SS稼働率(%)"]  < 20]["施設名"].tolist()
    red_acc_list  = df[df["重大事故件数"] >= 1]["施設名"].tolist()
    zero_sales    = df[df["営業訪問件数"] == 0]["施設名"].tolist()
    zero_hiyari   = df[df["ヒヤリハット件数"] == 0]["施設名"].tolist()
    high_dispatch = df[df["派遣利用率(%)"] >= 25]["施設名"].tolist()

    sections = {
        "A｜稼働・収益の確認": [
            f"全施設の稼働率を横断で確認した（🔴要確認: {red_occ_list or 'なし'}）",
            f"SS稼働率が低い施設の原因を特定した（🔴<20%: {red_ss_list or 'なし'}）",
            "夜勤加配算定漏れがある施設に対応した",
            "算定単位数の前月比を確認し、増減理由を把握した",
        ],
        "B｜人員コストの確認": [
            f"派遣利用率が高い施設の構造原因を分類した（🔴>25%: {high_dispatch or 'なし'}）",
            "タイミー応募率・採用率が低い施設に改善策を打った",
            "施設間のタイミー・派遣コストのばらつきを比較した",
        ],
        "C｜採用・定着の確認": [
            "入職0・退職1以上の施設を要注意としてフラグ立てした",
            "AMに対して採用活動の状況をヒアリングした",
        ],
        "D｜リスク・品質の確認": [
            f"重大事故発生施設の再発防止策を確認した（🔴: {red_acc_list or 'なし'}）",
            f"ヒヤリハット0件施設に未報告でないか確認した（要確認: {zero_hiyari or 'なし'}）",
            "苦情件数増加施設の原因・対応状況を把握した",
        ],
        "E｜ブランディング・営業の確認": [
            f"営業訪問0件施設に管理者へ確認・指示した（🔴: {zero_sales or 'なし'}）",
            "SS問い合わせ件数が少ない施設の営業導線を確認した",
            "見学・体験者から契約に至っていない施設の原因を確認した",
        ],
        "F｜AMマネジメント": [
            "全AMの月次報告の質を確認した（KPI + 原因 + 対策の形式か）",
            "報告の質が低いAMへフィードバックした",
            "AM全員と1on1を実施した",
        ],
        "G｜SM向け報告作成": [
            "🔴施設のPL影響を金額で算出した",
            "「現状・原因・対策・PL影響・SM判断事項」の形で報告を作成した",
            "選択肢と推奨案を添えた（「どうしましょうか」ではなくA/B/C案を提示）",
        ],
    }

    total_items = sum(len(v) for v in sections.values())
    checked = 0

    for section, items in sections.items():
        st.subheader(section)
        for item in items:
            key = f"check_{section}_{item}"
            if st.checkbox(item, key=key):
                checked += 1
        st.write("")

    st.divider()
    progress = checked / total_items if total_items > 0 else 0
    st.progress(progress)
    st.markdown(f"**進捗: {checked} / {total_items} 完了　({progress*100:.0f}%)**")

    if progress == 1.0:
        st.success("✅ 全チェック完了！SM向け報告を作成してください。")


# ════════════════════════════════════════════════════════════
# PAGE 4: トレンド分析
# ════════════════════════════════════════════════════════════
elif page == "📈 トレンド分析":
    st.title("📈 施設間 トレンド比較")

    metric_options = {
        "稼働率(%)":         "occupancy_rate",
        "SS稼働率(%)":       "ss_rate",
        "夜勤加配率(%)":     "night_shift_calc_rate",
        "派遣合計時間":      "派遣合計時間",
        "タイミー日勤時間":  "タイミー日勤時間",
        "重大事故件数":      "重大事故件数",
        "入職者数":          "入職者数",
        "退職者数":          "退職者数",
    }

    selected_metric_label = st.selectbox("比較する指標", list(metric_options.keys()))
    selected_metric_key   = metric_options[selected_metric_label]

    all_rows = []
    for f in facilities:
        for m in f.get("monthly", []):
            val = m.get(selected_metric_key)
            if val is not None:
                all_rows.append({
                    "月":    m["month"],
                    "施設名": f["facility_name"],
                    "値":    val,
                })

    if not all_rows:
        st.warning("データがありません。")
    else:
        tdf = pd.DataFrame(all_rows)
        fig = px.line(
            tdf, x="月", y="値", color="施設名",
            title=f"{selected_metric_label} 施設別推移",
            markers=True,
        )
        st.plotly_chart(fig, use_container_width=True)

        # 最新月の施設間比較
        latest = tdf[tdf["月"] == tdf["月"].max()]
        fig2 = px.bar(
            latest.sort_values("値", ascending=False),
            x="施設名", y="値",
            title=f"{selected_metric_label}（最新月）施設間比較",
            color="値",
            color_continuous_scale="RdYlGn" if metric_options[selected_metric_label] in
                ["occupancy_rate", "ss_rate", "night_shift_calc_rate"] else "RdYlGn_r",
        )
        st.plotly_chart(fig2, use_container_width=True)
