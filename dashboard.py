import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ─────────────────────────────────────────
# ページ設定
# ─────────────────────────────────────────
st.set_page_config(
    page_title="ALPFA Access Log Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ───────────────────────── ────────────────
# 定数
# ─────────────────────────────────────────
# 入力ファイルのパス（ローカル環境に合わせて変更すること）
LOG_PATH  = Path(r"C:\Users\N206876\Documents\Tableau_aceess_log\CP-260413-1008.xlsx")
USER_PATH = Path(r"C:\Users\N206876\Documents\Tableau_aceess_log\【PROD】ALPFA Tableau users Information.xlsx")

# ヒートマップの曜日表示順（月曜始まり）
WEEKDAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ─────────────────────────────────────────
# データ読み込み
# ─────────────────────────────────────────

@st.cache_data
def load_log(src) -> pd.DataFrame:
    """
    アクセスログExcelを読み込み、分析に必要な派生列を付与して返す。

    src にはファイルパス（Path）またはアップロードされたファイルオブジェクト
    （UploadedFile / BytesIO）を受け取る。これにより、ローカルファイルと
    アップロードファイルの両方を同一関数で処理できる。

    派生列:
    - timestamp  : 日時型に変換。変換できない行は除外する。
    - year_month : 月次集計用の期間文字列（例: "2024-01"）
    - weekday    : 曜日名（英語）
    - hour       : アクセス時刻の時間部分（0〜23）
    """
    df = pd.read_excel(src, sheet_name="Sheet1", header=1)
    df.columns = ["id", "project", "event_type", "timestamp", "user", "workbook", "view"]
    df["timestamp"] = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["timestamp"])
    df["year_month"] = df["timestamp"].dt.to_period("M").astype(str)
    df["weekday"]    = df["timestamp"].dt.day_name()
    df["hour"]       = df["timestamp"].dt.hour
    return df


@st.cache_data
def load_users(src) -> pd.DataFrame:
    """
    ユーザー情報Excelを読み込み、必要列のみ抽出・整形して返す。

    src にはファイルパス（Path）またはアップロードされたファイルオブジェクト
    （UploadedFile / BytesIO）を受け取る。

    整形内容:
    - full_name : 姓名を結合した表示名（ログ上での氏名表示に使用）
    - user      : ログとの結合キー。前後の空白を除去して一致精度を高める。
    """
    df = pd.read_excel(src, sheet_name="Personal User list")
    df = df[["First Name", "Last Name", "User Name", "所属", "Dept", "Status"]].copy()
    df.columns = ["first_name", "last_name", "user", "company", "dept", "status"]
    df["full_name"] = (df["first_name"].fillna("") + " " + df["last_name"].fillna("")).str.strip()
    df["user"] = df["user"].astype(str).str.strip()
    return df


def build_df_raw(log_src, user_src) -> pd.DataFrame:
    """
    ログとユーザー情報を結合して分析用の基本データフレームを生成する。

    アップロードファイルとローカルファイルのどちらが渡されても同じ処理を行う。
    ユーザー情報が存在しない行は display_name にユーザーIDをフォールバックとして使用する。
    """
    df_log  = load_log(log_src)
    df_user = load_users(user_src)
    df = df_log.merge(df_user[["user", "full_name", "company", "dept"]], on="user", how="left")
    df["display_name"] = df["full_name"].where(
        df["full_name"].notna() & (df["full_name"] != ""), df["user"]
    )
    return df


# ─────────────────────────────────────────
# サイドバー – ファイルアップロード
# ─────────────────────────────────────────
# 両ファイルがアップロードされるまでダッシュボードを表示しない。
# アップロード前はサイドバーのアップロードUIのみを表示し、
# メインエリアには案内メッセージを表示してアプリを停止する。
st.sidebar.markdown("## 📂 Data Source")
st.sidebar.markdown('<p class="sidebar-section">📋 Upload Files</p>', unsafe_allow_html=True)

uploaded_log  = st.sidebar.file_uploader(
    "Access Log Excel",
    type=["xlsx"],
    help="アクセスログファイル（Sheet1、ヘッダー2行目）をアップロードしてください。",
)
uploaded_user = st.sidebar.file_uploader(
    "User Info",
    type=["xlsx"],
    help="ユーザー情報ファイル（Personal User list シート）をアップロードしてください。",
)

# 両ファイルがアップロードされていない場合は案内を表示して停止する
if uploaded_log is None or uploaded_user is None:
    st.title("📊 ALPFA Tableau Access Dashboard")
    st.info("👈 サイドバーから **Access Log Excel** と **User Info** の両方をアップロードしてください。")
    st.stop()

# アップロード済みファイル名をサイドバーに表示する
st.sidebar.caption(f"✅ {uploaded_log.name}")
st.sidebar.caption(f"✅ {uploaded_user.name}")

st.sidebar.divider()

# アップロードされたファイルからデータを構築する
try:
    df_raw = build_df_raw(uploaded_log, uploaded_user)
except Exception as e:
    # ファイル読み込みに失敗した場合はエラーを表示してアプリを停止する
    st.error(f"ファイルの読み込みに失敗しました。ファイル形式を確認してください。\n\n{e}")
    st.stop()

# ─────────────────────────────────────────
# ヘルパー関数
# ─────────────────────────────────────────

def make_ranking_df(series: pd.Series, col_name: str) -> pd.DataFrame:
    """
    指定列の値カウントからランキング表を生成する。

    ランキング表示・構成比の計算を一元化することで、
    各タブでの重複実装を避け保守性を高める。

    出力列:
    - Rank         : 順位（1始まり）
    - {col_name}   : 対象の値
    - Access Count : アクセス件数
    - Share (%)    : 全体に対する構成比（小数点1桁）
    """
    df_rank = series.value_counts().reset_index()
    df_rank.columns = [col_name, "Access Count"]
    df_rank.insert(0, "Rank", range(1, len(df_rank) + 1))
    df_rank["Share (%)"] = (df_rank["Access Count"] / df_rank["Access Count"].sum() * 100).round(1)
    return df_rank


def add_rank_and_share(df_agg: pd.DataFrame, name_col: str, count_col: str) -> pd.DataFrame:
    """
    集計済み DataFrame に順位・構成比列を付与して返す。

    make_ranking_df は pd.Series.value_counts() を起点とするため、
    groupby などで既に集計済みの DataFrame には直接使えない。
    本関数はその補完として、同じ「Rank / Share (%)」付与ロジックを提供する。

    引数:
    - df_agg   : 集計済み DataFrame（降順ソート済みを想定）
    - name_col : 名称列の列名
    - count_col: 件数列の列名

    出力列:
    - Rank       : 順位（1始まり）
    - {name_col} : 名称列（元の列名をそのまま使用）
    - {count_col}: 件数列（元の列名をそのまま使用）
    - Share (%)  : 全体に対する構成比（小数点1桁）
    """
    result = df_agg[[name_col, count_col]].reset_index(drop=True).copy()
    result.insert(0, "Rank", range(1, len(result) + 1))
    result["Share (%)"] = (result[count_col] / result[count_col].sum() * 100).round(1)
    return result


def plot_hbar(df_plot: pd.DataFrame, x: str, y: str, color_scale: str,
              height: int = 440, top_n: int | None = None) -> go.Figure:
    """
    水平棒グラフを生成して返す。

    複数タブで同じ形式の棒グラフを使用するため、共通化して重複を排除している。
    - top_n を指定すると上位N件のみ表示（省略時は全件）
    - カラースケールは呼び出し元で指定し、凡例は非表示にする
    """
    data = df_plot.head(top_n) if top_n else df_plot
    fig = px.bar(
        data, x=x, y=y, orientation="h",
        color=x, color_continuous_scale=color_scale,
        labels={x: "Access Count"},
    )
    fig.update_layout(
        height=height,
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False,
    )
    return fig


def plot_heatmap(z, x_labels, y_labels, colorscale: str,
                 hover_template: str, height: int = 360,
                 xaxis_title: str = "", yaxis_title: str = "") -> go.Figure:
    """
    ヒートマップを生成して返す。

    曜日×時間帯・会社×ワークブックなど複数の用途で使用するため共通化している。
    - z             : 2次元の数値配列（行=y_labels、列=x_labels）
    - hover_template: マウスオーバー時の表示フォーマット
    """
    fig = go.Figure(go.Heatmap(
        z=z, x=x_labels, y=y_labels,
        colorscale=colorscale,
        hovertemplate=hover_template,
    ))
    fig.update_layout(
        xaxis_title=xaxis_title, yaxis_title=yaxis_title, height=height,
    )
    return fig


def plot_top_n_section(
    df_src: pd.DataFrame,
    col: str,
    label: str,
    icon: str,
    color_scale: str,
    top_n: int = 5,
    height: int = 320,
) -> None:
    """
    指定列の上位N件をグラフ＋表で Streamlit に描画する。

    User Deep Dive の Dashboard Top 5 / Project Top 5 は、
    対象列・表示名・カラースケールが異なるだけで処理が同一のため、
    本関数に統合して重複を排除している。

    引数:
    - df_src      : 対象ユーザーに絞り込み済みの DataFrame
    - col         : 集計対象の列名（例: "workbook", "project"）
    - label       : 表示用の列ヘッダー名（例: "Dashboard", "Project"）
    - icon        : セクションタイトルに付与する絵文字
    - color_scale : 棒グラフのカラースケール名（Plotly 形式）
    - top_n       : 表示する上位件数（デフォルト: 5）
    - height      : グラフの高さ（px）
    """
    st.markdown(f"#### {icon} {label} Top {top_n}")

    # 上位N件を集計し、順位・構成比を付与する
    top_df = df_src[col].value_counts().head(top_n).reset_index()
    top_df.columns = [label, "Access Count"]
    top_df.insert(0, "Rank", range(1, len(top_df) + 1))
    top_df["Share (%)"] = (
        top_df["Access Count"] / top_df["Access Count"].sum() * 100
    ).round(1)

    # 水平棒グラフ：アクセス数をバーの外側にテキスト表示する
    fig = px.bar(
        top_df, x="Access Count", y=label, orientation="h",
        color="Access Count", color_continuous_scale=color_scale,
        text="Access Count",
    )
    fig.update_layout(
        height=height,
        yaxis={"categoryorder": "total ascending"},
        coloraxis_showscale=False,
    )
    fig.update_traces(textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

    # 表形式でも同データを表示し、数値の詳細を確認できるようにする
    st.dataframe(top_df, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────
# カスタムCSS – サイドバーのスタイル
# ─────────────────────────────────────────
# サイドバーの視認性・操作性を高めるためのスタイル定義。
# Streamlit のデフォルトスタイルを上書きしてダークテーマに統一する。
st.markdown("""
<style>
/* サイドバー背景：ダークグラデーション */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f1117 0%, #1a1d2e 100%);
}

/* フィルターラベル：小文字・大文字変換・強調 */
[data-testid="stSidebar"] label {
    color: #a0aec0 !important;
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}

/* multiselect タグ：ダーク背景 */
[data-testid="stSidebar"] [data-baseweb="tag"] {
    background-color: #2d3748 !important;
    border-radius: 4px !important;
}

/* セクション区切り線 */
[data-testid="stSidebar"] hr {
    border-color: #2d3748 !important;
    margin: 0.6rem 0 !important;
}

/* フィルター件数バッジ：グラデーション背景 */
.filter-badge {
    background: linear-gradient(90deg, #667eea, #764ba2);
    color: white;
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 0.82rem;
    font-weight: 600;
    display: inline-block;
    margin-top: 4px;
    letter-spacing: 0.03em;
}

/* セクションヘッダー：左ボーダーで視覚的に区切る */
.sidebar-section {
    color: #667eea;
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin: 14px 0 4px 0;
    border-left: 3px solid #667eea;
    padding-left: 8px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# サイドバー – グローバルフィルター
# ─────────────────────────────────────────
# 全タブに共通するフィルターをサイドバーに集約する。
# フィルター変更時は df が再計算され、全タブの表示に即時反映される。
st.sidebar.markdown("## 🔍 Filters")

# ── 期間フィルター ────────────────────────
st.sidebar.markdown('<p class="sidebar-section">📅 Date Range</p>', unsafe_allow_html=True)
min_date = df_raw["timestamp"].min().date()
max_date = df_raw["timestamp"].max().date()
date_range = st.sidebar.date_input(
    "Date Range", label_visibility="collapsed",
    value=(min_date, max_date), min_value=min_date, max_value=max_date,
)

st.sidebar.divider()

# ── 組織フィルター ────────────────────────
st.sidebar.markdown('<p class="sidebar-section">🏢 Organization</p>', unsafe_allow_html=True)
all_companies = sorted(df_raw["company"].dropna().unique())
sel_companies = st.sidebar.multiselect("Company", all_companies, default=all_companies)

all_countries = sorted(df_raw["project"].dropna().unique())
sel_countries = st.sidebar.multiselect("Project Name", all_countries, default=all_countries)

st.sidebar.divider()

# ── コンテンツフィルター ──────────────────
st.sidebar.markdown('<p class="sidebar-section">📚 Content</p>', unsafe_allow_html=True)
all_workbooks = sorted(df_raw["workbook"].dropna().unique())
sel_workbooks = st.sidebar.multiselect("Dashboard", all_workbooks, default=all_workbooks)

all_views = sorted(df_raw["view"].dropna().unique())
sel_views = st.sidebar.multiselect("View (Tab)", all_views, default=all_views)

st.sidebar.divider()

# ── フィルター適用 ────────────────────────
# 各フィルターが選択されている場合のみ絞り込みを実施する。
# 未選択（空リスト）の場合は全件を対象とする。
# これにより「全選択」と「未選択」を同じ挙動にし、ユーザーの誤操作を防ぐ。
df = df_raw.copy()
if sel_companies: df = df[df["company"].isin(sel_companies)]
if sel_countries: df = df[df["project"].isin(sel_countries)]
if sel_workbooks: df = df[df["workbook"].isin(sel_workbooks)]
if sel_views:     df = df[df["view"].isin(sel_views)]
if len(date_range) == 2:
    # date_input は選択途中（開始日のみ）の場合に要素数1のタプルを返すため、
    # 2要素が揃った場合のみ期間フィルターを適用する。
    s, e = date_range
    df = df[(df["timestamp"].dt.date >= s) & (df["timestamp"].dt.date <= e)]

# フィルター後の件数と全体に対する割合をバッジ表示する。
# ユーザーが現在どの範囲のデータを見ているかを一目で把握できるようにする。
pct = len(df) / len(df_raw) * 100 if len(df_raw) > 0 else 0
st.sidebar.markdown(
    f'<div class="filter-badge">⚡ {len(df):,} / {len(df_raw):,} rows &nbsp;·&nbsp; {pct:.0f}%</div>',
    unsafe_allow_html=True,
)

# ─────────────────────────────────────────
# タブ定義
# ─────────────────────────────────────────
tab_overview, tab_ranking, tab_user, tab_org = st.tabs([
    "📊 Overview",
    "🏆 Rankings",
    "🔬 User Analysis",
    "🏢 Org Analysis",
])

# ══════════════════════════════════════════
# TAB 1 – OVERVIEW
# アクセス全体の傾向を把握するためのサマリービュー
# ══════════════════════════════════════════
with tab_overview:
    st.title("📊 ALPFA Tableau Access Dashboard")

    # KPI メトリクス：フィルター後データの主要指標を4列で表示
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Access",  f"{len(df):,}")
    c2.metric("Unique Users",  f"{df['user'].nunique():,}")
    c3.metric("Countries",     f"{df['project'].nunique():,}")
    c4.metric("Dashboards",    f"{df['workbook'].nunique():,}")

    st.divider()

    # 月次アクセス推移：月ごとのアクセス件数を棒グラフで可視化
    st.subheader("📈 Monthly Access Trend")
    monthly = df.groupby("year_month").size().reset_index(name="count")
    fig = px.bar(monthly, x="year_month", y="count",
                 labels={"year_month": "Month", "count": "Access Count"},
                 color_discrete_sequence=["#1f77b4"])
    fig.update_layout(xaxis_tickangle=-45, height=340)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    col_l, col_r = st.columns(2)

    # 国別アクセス：上位20カ国を水平棒グラフで表示
    # project 列が国/プロジェクト識別子として使われているため、表示名を "country" に統一する
    with col_l:
        st.subheader("🌍 Access by Country")
        country_cnt = df["project"].value_counts().reset_index(name="count")
        country_cnt.columns = ["country", "count"]
        st.plotly_chart(
            plot_hbar(country_cnt, x="count", y="country", color_scale="Blues", top_n=20),
            use_container_width=True,
        )

    # ダッシュボード別アクセス：上位15件を水平棒グラフで表示
    with col_r:
        st.subheader("📚 Access by Dashboard")
        workbook_cnt = df["workbook"].value_counts().reset_index(name="count")
        workbook_cnt.columns = ["workbook", "count"]
        st.plotly_chart(
            plot_hbar(workbook_cnt, x="count", y="workbook", color_scale="Greens", top_n=15),
            use_container_width=True,
        )

    # 曜日×時間帯ヒートマップ：アクセスが集中する時間帯を把握するために使用
    st.divider()
    st.subheader("🗓️ Access Heatmap – Weekday × Hour (JST)")
    # groupby で曜日×時間帯の件数を集計し、pivot で行=曜日・列=時間帯の2次元配列に変換する。
    # reindex で月曜始まりの表示順に並べ替え、データが存在しないセルは 0 で補完する。
    heat = (
        df.groupby(["weekday", "hour"]).size()
          .reset_index(name="count")
          .pivot(index="weekday", columns="hour", values="count")
          .reindex(WEEKDAY_ORDER)
          .fillna(0)
    )
    st.plotly_chart(
        plot_heatmap(
            z=heat.values,
            x_labels=[f"{h:02d}:00" for h in heat.columns],
            y_labels=heat.index.tolist(),
            colorscale="YlOrRd",
            hover_template="Day: %{y}<br>Hour: %{x}<br>Count: %{z}<extra></extra>",
            height=360,
            xaxis_title="Hour (JST)",
            yaxis_title="Weekday",
        ),
        use_container_width=True,
    )


# ══════════════════════════════════════════
# TAB 2 – RANKINGS
# ワークブック・ビュー・国別のアクセスランキングを一覧表示する
# ══════════════════════════════════════════
with tab_ranking:
    st.title("🏆 Usage Rankings")

    col_l, col_r = st.columns(2)

    # ダッシュボードランキング：アクセス数順に全ダッシュボードを表示
    with col_l:
        st.subheader("📚 Dashboard Ranking")
        st.dataframe(
            make_ranking_df(df["workbook"], "Dashboard"),
            use_container_width=True, height=420, hide_index=True,
        )

    # ビュー（タブ）ランキング：アクセス数順に全ビューを表示
    with col_r:
        st.subheader("📑 View (Tab) Ranking")
        st.dataframe(
            make_ranking_df(df["view"], "View"),
            use_container_width=True, height=420, hide_index=True,
        )

    st.divider()

    # 国別ランキング：表とドーナツグラフを並べて表示
    st.subheader("🌍 Country Ranking")
    cr = make_ranking_df(df["project"], "Country")

    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(cr, use_container_width=True, height=400, hide_index=True)
    with c2:
        fig = px.pie(cr.head(15), values="Access Count", names="Country",
                     title="Top 15 Countries", hole=0.35)
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════
# TAB 3 – USER ANALYSIS
# 個人ユーザー単位でのアクセス傾向を深掘りするビュー
# ══════════════════════════════════════════
with tab_user:
    st.title("🔬 User Analysis")

    # 月次アクティブユーザー数：月ごとのユニークユーザー数の推移を折れ線グラフで表示
    st.subheader("👥 Monthly Active Users")
    mau = df.groupby("year_month")["user"].nunique().reset_index(name="active_users")
    fig = px.line(mau, x="year_month", y="active_users", markers=True,
                  labels={"year_month": "Month", "active_users": "Unique Active Users"},
                  color_discrete_sequence=["#e377c2"])
    fig.update_layout(xaxis_tickangle=-45, height=300)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ヘビーユーザー Top 20：アクセス数上位20名を表とグラフで並べて表示
    # groupby でユーザーごとのアクセス件数を集計し、上位20名に絞る。
    # rename をメソッドチェーン内に統合し、後続の再代入を不要にする。
    st.subheader("🔥 Heavy Users – Top 20")
    user_cnt = (
        df.groupby(["user", "display_name", "company", "dept"])
          .size().reset_index(name="Access Count")
          .sort_values("Access Count", ascending=False)
          .head(20)
          .rename(columns={"user": "User ID", "display_name": "Name",
                            "company": "Company", "dept": "Dept"})
          .reset_index(drop=True)
    )
    # 順位列を先頭に挿入する（reset_index 後に付与することで連番が確定する）
    user_cnt.insert(0, "Rank", range(1, len(user_cnt) + 1))

    u1, u2 = st.columns(2)
    with u1:
        st.dataframe(
            user_cnt[["Rank", "Name", "User ID", "Company", "Dept", "Access Count"]],
            use_container_width=True, height=420, hide_index=True,
        )
    with u2:
        st.plotly_chart(
            plot_hbar(user_cnt, x="Access Count", y="Name", color_scale="Reds", height=420),
            use_container_width=True,
        )

    st.divider()

    # ── ユーザー個別分析 ──────────────────────────────────────────────────────
    # ユーザーを選択すると、そのユーザーがよく見ているDashboard・Projectの
    # Top 5 ランキングをグラフと表で表示する。
    st.subheader("🔍 User Deep Dive")

    # 表示名（display_name）でユーザーを選択できるようにする。
    # 「表示名 (UserID)」形式にすることで同姓同名でも一意に識別できる。
    user_options = (
        df[["user", "display_name"]]
        .drop_duplicates()
        .sort_values("display_name")
    )
    # iterrows より apply の方がベクトル化されており、ユーザー数が多い場合でも高速に動作する
    user_labels = (
        user_options["display_name"] + " (" + user_options["user"] + ")"
    ).tolist()
    selected_label = st.selectbox(
        "Select User",
        options=user_labels,
        index=0,
        help="分析したいユーザーを選択してください。",
    )

    if selected_label:
        # 選択ラベルの末尾の括弧内から user ID を逆引きする
        selected_user_id = selected_label.split("(")[-1].rstrip(")")
        df_selected = df[df["user"] == selected_user_id]

        total_access = len(df_selected)
        selected_name = df_selected["display_name"].iloc[0] if total_access > 0 else selected_user_id

        st.markdown(
            f'<div class="filter-badge">👤 {selected_name} &nbsp;·&nbsp; {total_access:,} accesses</div>',
            unsafe_allow_html=True,
        )
        st.write("")  # スペーサー

        if total_access == 0:
            st.info("選択したユーザーのアクセスデータが見つかりません。")
        else:
            d1, d2 = st.columns(2)

            # Dashboard Top 5 / Project Top 5 は処理が同一のため plot_top_n_section で共通化する
            with d1:
                plot_top_n_section(
                    df_src=df_selected, col="workbook",
                    label="Dashboard", icon="📚", color_scale="Greens",
                )
            with d2:
                plot_top_n_section(
                    df_src=df_selected, col="project",
                    label="Project", icon="🌍", color_scale="Blues",
                )


# ══════════════════════════════════════════
# TAB 4 – ORG ANALYSIS
# 会社・部署単位でのアクセス傾向を深掘りするビュー
# ══════════════════════════════════════════
with tab_org:
    st.title("🏢 Org Analysis")

    # 会社（所属）別アクセス件数を降順で集計する。
    # NaN 行は dropna で除外し、後続の表示・ヒートマップ両方で再利用する。
    comp_cnt = (
        df.groupby("company").size()
          .reset_index(name="count")
          .sort_values("count", ascending=False)
          .dropna(subset=["company"])
    )

    # 会社（所属）別アクセス：表とドーナツグラフを並べて表示
    st.subheader("🏢 Access by Company (所属)")
    c1, c2 = st.columns(2)
    with c1:
        # make_ranking_df は value_counts() ベースのため、groupby 集計済みの comp_cnt には使えない。
        # add_rank_and_share を使って順位・構成比を付与し、表示用テーブルを生成する。
        comp_tbl = add_rank_and_share(comp_cnt, name_col="company", count_col="count")
        comp_tbl.columns = ["Rank", "Company", "Access Count", "Share (%)"]
        st.dataframe(comp_tbl, use_container_width=True, height=380, hide_index=True)
    with c2:
        fig = px.pie(comp_cnt.head(12), values="count", names="company",
                     title="Top 12 Companies", hole=0.35)
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)


