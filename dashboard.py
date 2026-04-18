import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# ─────────────────────────────────────────
# Page config
# ─────────────────────────────────────────
st.set_page_config(
    page_title="ALPFA Access Log Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────
# Constants
# ─────────────────────────────────────────
LOG_PATH  = Path(r"C:\Users\N206876\Downloads\CP-260413-1008.xlsx")
USER_PATH = Path(r"C:\Users\N206876\Downloads\【PROD】ALPFA Tableau users Information.xlsx")

WEEKDAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# ─────────────────────────────────────────
# Data loading
# ─────────────────────────────────────────
@st.cache_data
def load_log(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Sheet1", header=1)
    df.columns = ["id", "project", "event_type", "timestamp", "user", "workbook", "view"]
    df = df.dropna(subset=["timestamp"])
    df["timestamp"]  = pd.to_datetime(df["timestamp"], errors="coerce")
    df = df.dropna(subset=["timestamp"])
    df["year_month"] = df["timestamp"].dt.to_period("M").astype(str)
    df["weekday"]    = df["timestamp"].dt.day_name()
    df["hour"]       = df["timestamp"].dt.hour
    return df


@st.cache_data
def load_users(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Personal User list")
    df = df[["First Name", "Last Name", "User Name", "所属", "Dept", "Status"]].copy()
    df.columns = ["first_name", "last_name", "user", "company", "dept", "status"]
    df["full_name"] = df["first_name"].fillna("") + " " + df["last_name"].fillna("")
    df["full_name"] = df["full_name"].str.strip()
    df["user"] = df["user"].astype(str).str.strip()
    return df


df_log  = load_log(LOG_PATH)
df_user = load_users(USER_PATH)

# Join: log ← user info
df_raw = df_log.merge(df_user[["user", "full_name", "company", "dept"]], on="user", how="left")
df_raw["display_name"] = df_raw["full_name"].where(df_raw["full_name"].notna() & (df_raw["full_name"] != ""), df_raw["user"])

# ─────────────────────────────────────────
# Sidebar – Global Filters
# ─────────────────────────────────────────
st.sidebar.title("🔍 Filters")

all_countries = sorted(df_raw["project"].dropna().unique())
sel_countries = st.sidebar.multiselect("Country / Project", all_countries, default=all_countries)

all_workbooks = sorted(df_raw["workbook"].dropna().unique())
sel_workbooks = st.sidebar.multiselect("Workbook", all_workbooks, default=all_workbooks)

all_views = sorted(df_raw["view"].dropna().unique())
sel_views = st.sidebar.multiselect("View (Tab)", all_views, default=all_views)

min_date = df_raw["timestamp"].min().date()
max_date = df_raw["timestamp"].max().date()
date_range = st.sidebar.date_input("Date Range", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date)

# Apply filters
df = df_raw.copy()
if sel_countries: df = df[df["project"].isin(sel_countries)]
if sel_workbooks: df = df[df["workbook"].isin(sel_workbooks)]
if sel_views:     df = df[df["view"].isin(sel_views)]
if len(date_range) == 2:
    s, e = date_range
    df = df[(df["timestamp"].dt.date >= s) & (df["timestamp"].dt.date <= e)]

st.sidebar.markdown(f"**Filtered:** {len(df):,} / {len(df_raw):,} rows")

# ─────────────────────────────────────────
# Tabs
# ─────────────────────────────────────────
tab_overview, tab_ranking, tab_user = st.tabs([
    "📊 Overview",
    "🏆 Rankings",
    "🔬 User & Org Analysis",
])

# ══════════════════════════════════════════
# TAB 1 – OVERVIEW
# ══════════════════════════════════════════
with tab_overview:
    st.title("📊 ALPFA Tableau Access Dashboard")

    # KPIs
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Access",    f"{len(df):,}")
    c2.metric("Unique Users",    f"{df['user'].nunique():,}")
    c3.metric("Countries",       f"{df['project'].nunique():,}")
    c4.metric("Workbooks",       f"{df['workbook'].nunique():,}")

    st.divider()

    # Monthly trend
    st.subheader("📈 Monthly Access Trend")
    monthly = df.groupby("year_month").size().reset_index(name="count")
    fig = px.bar(monthly, x="year_month", y="count",
                 labels={"year_month": "Month", "count": "Access Count"},
                 color_discrete_sequence=["#1f77b4"])
    fig.update_layout(xaxis_tickangle=-45, height=340)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    col_l, col_r = st.columns(2)

    # Country breakdown
    with col_l:
        st.subheader("🌍 Access by Country")
        cnt = df["project"].value_counts().reset_index()
        cnt.columns = ["country", "count"]
        fig = px.bar(cnt.head(20), x="count", y="country", orientation="h",
                     color="count", color_continuous_scale="Blues",
                     labels={"count": "Access Count", "country": "Country"})
        fig.update_layout(height=440, yaxis={"categoryorder": "total ascending"},
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    # Workbook breakdown
    with col_r:
        st.subheader("📚 Access by Workbook")
        cnt = df["workbook"].value_counts().reset_index()
        cnt.columns = ["workbook", "count"]
        fig = px.bar(cnt.head(15), x="count", y="workbook", orientation="h",
                     color="count", color_continuous_scale="Greens",
                     labels={"count": "Access Count", "workbook": "Workbook"})
        fig.update_layout(height=440, yaxis={"categoryorder": "total ascending"},
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    # Heatmap
    st.divider()
    st.subheader("�️ Access Heatmap – Weekday × Hour (JST)")
    heat = (df.groupby(["weekday", "hour"]).size().reset_index(name="count")
              .pivot(index="weekday", columns="hour", values="count")
              .reindex(WEEKDAY_ORDER).fillna(0))
    fig = go.Figure(go.Heatmap(
        z=heat.values,
        x=[f"{h:02d}:00" for h in heat.columns],
        y=heat.index.tolist(),
        colorscale="YlOrRd",
        hovertemplate="Day: %{y}<br>Hour: %{x}<br>Count: %{z}<extra></extra>",
    ))
    fig.update_layout(xaxis_title="Hour (JST)", yaxis_title="Weekday", height=360)
    st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════
# TAB 2 – RANKINGS
# ══════════════════════════════════════════
with tab_ranking:
    st.title("🏆 Usage Rankings")

    col_l, col_r = st.columns(2)

    # Workbook ranking
    with col_l:
        st.subheader("📚 Workbook Ranking")
        wb = df["workbook"].value_counts().reset_index()
        wb.columns = ["Workbook", "Access Count"]
        wb.insert(0, "Rank", range(1, len(wb) + 1))
        wb["Share (%)"] = (wb["Access Count"] / wb["Access Count"].sum() * 100).round(1)
        st.dataframe(wb, use_container_width=True, height=420, hide_index=True)

    # View ranking
    with col_r:
        st.subheader("📑 View (Tab) Ranking")
        vw = df["view"].value_counts().reset_index()
        vw.columns = ["View", "Access Count"]
        vw.insert(0, "Rank", range(1, len(vw) + 1))
        vw["Share (%)"] = (vw["Access Count"] / vw["Access Count"].sum() * 100).round(1)
        st.dataframe(vw, use_container_width=True, height=420, hide_index=True)

    st.divider()

    # Country ranking
    st.subheader("🌍 Country Ranking")
    cr = df["project"].value_counts().reset_index()
    cr.columns = ["Country", "Access Count"]
    cr.insert(0, "Rank", range(1, len(cr) + 1))
    cr["Share (%)"] = (cr["Access Count"] / cr["Access Count"].sum() * 100).round(1)

    c1, c2 = st.columns([1, 1])
    with c1:
        st.dataframe(cr, use_container_width=True, height=400, hide_index=True)
    with c2:
        fig = px.pie(cr.head(15), values="Access Count", names="Country",
                     title="Top 15 Countries", hole=0.35)
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════
# TAB 3 – USER & ORG ANALYSIS
# ══════════════════════════════════════════
with tab_user:
    st.title("🔬 User & Organization Analysis")

    # ── Monthly Active Users ──────────────
    st.subheader("👥 Monthly Active Users")
    mau = (df.groupby("year_month")["user"].nunique()
             .reset_index(name="active_users"))
    fig = px.line(mau, x="year_month", y="active_users", markers=True,
                  labels={"year_month": "Month", "active_users": "Unique Active Users"},
                  color_discrete_sequence=["#e377c2"])
    fig.update_layout(xaxis_tickangle=-45, height=300)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Heavy Users Top 20 (with name) ────
    st.subheader("🔥 Heavy Users – Top 20")
    user_cnt = (df.groupby(["user", "display_name", "company", "dept"])
                  .size().reset_index(name="Access Count")
                  .sort_values("Access Count", ascending=False)
                  .head(20)
                  .reset_index(drop=True))
    user_cnt.insert(0, "Rank", range(1, len(user_cnt) + 1))
    user_cnt = user_cnt.rename(columns={
        "user": "User ID", "display_name": "Name",
        "company": "Company", "dept": "Dept"
    })

    u1, u2 = st.columns([1, 1])
    with u1:
        st.dataframe(user_cnt[["Rank", "Name", "User ID", "Company", "Dept", "Access Count"]],
                     use_container_width=True, height=420, hide_index=True)
    with u2:
        fig = px.bar(user_cnt, x="Access Count", y="Name", orientation="h",
                     color="Access Count", color_continuous_scale="Reds",
                     labels={"Access Count": "Access Count", "Name": "User"})
        fig.update_layout(height=420, yaxis={"categoryorder": "total ascending"},
                          coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Access by Company (所属) ──────────
    st.subheader("🏢 Access by Company (所属)")
    comp_cnt = (df.groupby("company").size().reset_index(name="count")
                  .sort_values("count", ascending=False)
                  .dropna(subset=["company"]))

    c1, c2 = st.columns([1, 1])
    with c1:
        comp_tbl = comp_cnt.copy()
        comp_tbl.insert(0, "Rank", range(1, len(comp_tbl) + 1))
        comp_tbl.columns = ["Rank", "Company", "Access Count"]
        comp_tbl["Share (%)"] = (comp_tbl["Access Count"] / comp_tbl["Access Count"].sum() * 100).round(1)
        st.dataframe(comp_tbl, use_container_width=True, height=380, hide_index=True)
    with c2:
        fig = px.pie(comp_cnt.head(12), values="count", names="company",
                     title="Top 12 Companies", hole=0.35)
        fig.update_layout(height=380)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Access by Dept ────────────────────
    st.subheader("🗂️ Access by Department (Dept)")
    dept_cnt = (df.groupby("dept").size().reset_index(name="count")
                  .sort_values("count", ascending=False)
                  .dropna(subset=["dept"])
                  .head(20))
    fig = px.bar(dept_cnt, x="count", y="dept", orientation="h",
                 color="count", color_continuous_scale="Purples",
                 labels={"count": "Access Count", "dept": "Department"})
    fig.update_layout(height=480, yaxis={"categoryorder": "total ascending"},
                      coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # ── Company × Workbook heatmap ────────
    st.subheader("🔥 Company × Workbook Usage Heatmap")
    top_companies = comp_cnt.head(15)["company"].tolist()
    top_workbooks = df["workbook"].value_counts().head(10).index.tolist()
    heat_df = (df[df["company"].isin(top_companies) & df["workbook"].isin(top_workbooks)]
               .groupby(["company", "workbook"]).size().reset_index(name="count")
               .pivot(index="company", columns="workbook", values="count")
               .fillna(0))
    fig = go.Figure(go.Heatmap(
        z=heat_df.values,
        x=heat_df.columns.tolist(),
        y=heat_df.index.tolist(),
        colorscale="Blues",
        hovertemplate="Company: %{y}<br>Workbook: %{x}<br>Count: %{z}<extra></extra>",
    ))
    fig.update_layout(xaxis_tickangle=-35, height=420,
                      xaxis_title="Workbook", yaxis_title="Company")
    st.plotly_chart(fig, use_container_width=True)
