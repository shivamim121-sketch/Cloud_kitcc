# =========================================================
# KITCHEN P&L INTELLIGENCE SUITE
# FULL ENTERPRISE EDITION
# =========================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from pathlib import Path
import numpy as np
import hashlib

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Kitchen P&L Intelligence Suite",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# FILE PATH
# =========================================================

BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

# =========================================================
# WHITE ENTERPRISE CSS
# =========================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"]{
    font-family:'Inter',sans-serif;
}

.stApp{
    background:#f5f7fb;
}

/* SIDEBAR */

[data-testid="stSidebar"]{
    background:#ffffff;
    border-right:1px solid #e5e7eb;
}

[data-testid="stSidebar"] *{
    color:#111827;
}

/* HERO */

.hero-grid{
    display:grid;
    grid-template-columns:2fr 1fr;
    gap:24px;
    margin-bottom:28px;
}

.hero-main{
    background:white;
    border-radius:28px;
    padding:42px;
    border:1px solid #e5e7eb;
    box-shadow:0 10px 35px rgba(0,0,0,0.05);
}

.hero-badge{
    display:inline-block;
    background:#eef2ff;
    color:#4338ca;
    padding:8px 18px;
    border-radius:999px;
    font-size:12px;
    font-weight:800;
    letter-spacing:1px;
    margin-bottom:22px;
}

.hero-title{
    font-size:58px;
    font-weight:900;
    line-height:1;
    color:#111827;
    margin-bottom:20px;
}

.hero-title span{
    display:block;
    color:#4f46e5;
}

.hero-sub{
    font-size:16px;
    color:#6b7280;
    line-height:1.9;
}

.hero-chips{
    display:flex;
    flex-wrap:wrap;
    gap:12px;
    margin-top:28px;
}

.hero-chip{
    background:#f3f4f6;
    border-radius:14px;
    padding:10px 16px;
    font-size:14px;
    font-weight:600;
    color:#374151;
}

.hero-card{
    background:linear-gradient(135deg,#4f46e5 0%,#7c3aed 100%);
    border-radius:28px;
    padding:36px;
    color:white;
    box-shadow:0 18px 45px rgba(79,70,229,0.28);
}

.hero-card-title{
    font-size:12px;
    letter-spacing:2px;
    font-weight:700;
    opacity:0.85;
}

.hero-card-metric{
    font-size:56px;
    font-weight:900;
    margin-top:24px;
}

.hero-card-desc{
    margin-top:22px;
    line-height:1.9;
    color:rgba(255,255,255,0.92);
}

/* KPI */

.metric-card{
    background:white;
    border-radius:24px;
    padding:28px;
    border:1px solid #e5e7eb;
    box-shadow:0 8px 24px rgba(0,0,0,0.04);
}

.metric-title{
    font-size:12px;
    color:#6b7280;
    font-weight:800;
    letter-spacing:1px;
    text-transform:uppercase;
}

.metric-value{
    font-size:38px;
    font-weight:900;
    color:#111827;
    margin-top:10px;
}

.metric-delta{
    margin-top:12px;
    font-size:14px;
    font-weight:600;
}

.delta-pos{
    color:#059669;
}

.delta-neg{
    color:#dc2626;
}

/* SECTION */

.section-header{
    font-size:30px;
    font-weight:900;
    color:#111827;
    margin-top:40px;
    margin-bottom:18px;
}

/* INSIGHTS */

.insight-card{
    background:white;
    border-radius:24px;
    padding:30px;
    border:1px solid #e5e7eb;
    box-shadow:0 8px 20px rgba(0,0,0,0.04);
    height:100%;
}

.insight-title{
    font-size:12px;
    color:#6b7280;
    font-weight:800;
    letter-spacing:1px;
}

.insight-value{
    font-size:30px;
    font-weight:900;
    color:#111827;
    margin-top:14px;
}

.insight-desc{
    margin-top:14px;
    color:#6b7280;
    line-height:1.7;
}

/* TABS */

.stTabs [data-baseweb="tab"]{
    font-size:16px;
    font-weight:700;
}

/* TABLE */

[data-testid="stDataFrame"]{
    border-radius:20px;
    overflow:hidden;
}

/* FILTER PILLS */

.filter-pill{
    display:inline-block;
    background:#eef2ff;
    color:#4338ca;
    padding:8px 14px;
    border-radius:999px;
    font-size:13px;
    font-weight:700;
    margin-right:8px;
    margin-bottom:8px;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# DATA LOADING
# =========================================================

@st.cache_data(ttl=300)
def load_data(filepath):

    df = pd.read_excel(filepath, header=1)

    df["GM%"] = (
        df["GROSS MARGIN"] /
        df["NET REVENUE"] * 100
    ).round(2)

    df["EBITDA%"] = (
        df["KITCHEN EBITDA"] /
        df["NET REVENUE"] * 100
    ).round(2)

    df["VARIANCE%"] = (
        df["VARIANCE"] /
        df["NET REVENUE"] * 100
    ).round(2)

    df["CM"] = df["GROSS MARGIN"]
    df["CM%"] = df["GM%"]

    month_order = [
        "Oct-2023",
        "Nov-2023",
        "Dec-2023",
        "Jan-2024",
        "Feb-2024",
        "Mar-2024"
    ]

    df["MONTH"] = pd.Categorical(
        df["MONTH"],
        categories=month_order,
        ordered=True
    )

    df["MONTH_STR"] = df["MONTH"].astype(str)

    return df

df = load_data(DATA_PATH)

# =========================================================
# SIDEBAR
# =========================================================

def render_sidebar(df):

    st.sidebar.markdown("""
    <div style='padding-top:10px;padding-bottom:20px'>
        <div style='font-size:30px;font-weight:900;color:#111827'>
            🍳 Kitchen Intelligence
        </div>

        <div style='font-size:13px;color:#6b7280;margin-top:6px'>
            Enterprise Analytics Platform
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("---")

    dashboard = st.sidebar.radio(
        "Select Dashboard",
        [
            "🏪 Kitchen Level PNL",
            "📉 Variance Level PNL"
        ]
    )

    st.sidebar.markdown("---")

    filters = {}

    filters["month"] = st.sidebar.multiselect(
        "📅 Month",
        options=df["MONTH_STR"].unique(),
        default=df["MONTH_STR"].unique()
    )

    filters["city"] = st.sidebar.multiselect(
        "🏙️ City",
        options=sorted(df["CITY"].unique()),
        default=sorted(df["CITY"].unique())
    )

    filters["status"] = st.sidebar.multiselect(
        "🔵 Status",
        options=df["STATUS"].unique(),
        default=df["STATUS"].unique()
    )

    filters["store"] = st.sidebar.multiselect(
        "🏪 Store",
        options=sorted(df["STORE"].unique())
    )

    filters["zone"] = st.sidebar.multiselect(
        "🗺️ Zone",
        options=sorted(df["ZONE MAPPING"].unique()),
        default=sorted(df["ZONE MAPPING"].unique())
    )

    st.sidebar.markdown("---")

    rev_min = float(df["NET REVENUE"].min())
    rev_max = float(df["NET REVENUE"].max())

    filters["rev_range"] = st.sidebar.slider(
        "💰 Revenue Range",
        min_value=rev_min,
        max_value=rev_max,
        value=(rev_min, rev_max)
    )

    st.sidebar.markdown("---")

    st.sidebar.markdown("""
    <div style='text-align:center;font-size:12px;color:#6b7280'>
        Financial analytics environment for
        multi-city kitchen performance review.
    </div>
    """, unsafe_allow_html=True)

    return dashboard, filters

dashboard, filters = render_sidebar(df)

# =========================================================
# FILTER LOGIC
# =========================================================

fdf = df.copy()

if filters["month"]:
    fdf = fdf[fdf["MONTH_STR"].isin(filters["month"])]

if filters["city"]:
    fdf = fdf[fdf["CITY"].isin(filters["city"])]

if filters["status"]:
    fdf = fdf[fdf["STATUS"].isin(filters["status"])]

if filters["store"]:
    fdf = fdf[fdf["STORE"].isin(filters["store"])]

if filters["zone"]:
    fdf = fdf[fdf["ZONE MAPPING"].isin(filters["zone"])]

fdf = fdf[
    fdf["NET REVENUE"].between(
        filters["rev_range"][0],
        filters["rev_range"][1]
    )
]

# =========================================================
# METRICS
# =========================================================

total_rev = fdf["NET REVENUE"].sum()
avg_gm = fdf["GM%"].mean()
avg_ebitda = fdf["EBITDA%"].mean()
stores = fdf["STORE"].nunique()
best_city = fdf.groupby("CITY")["EBITDA%"].mean().idxmax()

# =========================================================
# HERO
# =========================================================

st.markdown(f"""
<div class='hero-grid'>

    <div class='hero-main'>

        <div class='hero-badge'>
            ENTERPRISE PERFORMANCE DASHBOARD
        </div>

        <div class='hero-title'>
            Kitchen P&L
            <span>Intelligence Suite</span>
        </div>

        <div class='hero-sub'>

            Unified operational reporting platform for
            multi-city cloud kitchen performance analysis,
            revenue benchmarking, profitability monitoring,
            variance tracking, and executive financial review.

            <br><br>

            Monitor kitchen efficiency, identify operational
            trends, compare market performance, and evaluate
            store-level profitability through a centralized
            analytics environment.

        </div>

        <div class='hero-chips'>

            <div class='hero-chip'>📊 Revenue Monitoring</div>
            <div class='hero-chip'>📈 EBITDA Tracking</div>
            <div class='hero-chip'>🏪 Multi-City Analysis</div>
            <div class='hero-chip'>📉 Variance Reporting</div>
            <div class='hero-chip'>⚡ Operational Intelligence</div>

        </div>

    </div>

    <div class='hero-card'>

        <div class='hero-card-title'>
            PERFORMANCE SNAPSHOT
        </div>

        <div class='hero-card-metric'>
            ₹{total_rev/1e7:.1f}Cr
        </div>

        <div class='hero-card-desc'>

            Financial review across
            <b>{stores}</b> kitchens with
            average EBITDA margin of
            <b>{avg_ebitda:.1f}%</b>.

            <br><br>

            Highest operational performance observed in
            <b>{best_city}</b>.

        </div>

    </div>

</div>
""", unsafe_allow_html=True)

# =========================================================
# FILTER PILLS
# =========================================================

st.markdown(
    f"""
    <div class='filter-pill'>
        Cities: {len(filters["city"])}
    </div>

    <div class='filter-pill'>
        Months: {len(filters["month"])}
    </div>

    <div class='filter-pill'>
        Stores: {fdf["STORE"].nunique()}
    </div>
    """,
    unsafe_allow_html=True
)

# =========================================================
# KPI STRIP
# =========================================================

col1, col2, col3, col4 = st.columns(4)

cards = [
    ("💰 Net Revenue", f"₹{total_rev/1e7:.1f}Cr", "Across selected kitchens"),
    ("📊 Avg GM%", f"{avg_gm:.1f}%", "Gross margin performance"),
    ("📈 Avg EBITDA%", f"{avg_ebitda:.1f}%", "Operational profitability"),
    ("🏪 Kitchens", f"{stores}", "Active locations")
]

for col, card in zip([col1,col2,col3,col4], cards):

    with col:

        st.markdown(f"""
        <div class='metric-card'>

            <div class='metric-title'>
                {card[0]}
            </div>

            <div class='metric-value'>
                {card[1]}
            </div>

            <div class='metric-delta delta-pos'>
                {card[2]}
            </div>

        </div>
        """, unsafe_allow_html=True)

# =========================================================
# TABS
# =========================================================

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Performance Table",
    "📈 Trends",
    "🏙️ Market Analysis",
    "🔮 Store Intelligence"
])

# =========================================================
# TAB 1
# =========================================================

with tab1:

    st.markdown(
        "<div class='section-header'>Kitchen Performance Overview</div>",
        unsafe_allow_html=True
    )

    st.dataframe(
        fdf[[
            "STORE",
            "CITY",
            "MONTH_STR",
            "NET REVENUE",
            "GM%",
            "EBITDA%",
            "VARIANCE%"
        ]],
        use_container_width=True,
        height=600
    )

# =========================================================
# TAB 2
# =========================================================

with tab2:

    st.markdown(
        "<div class='section-header'>Revenue & Profitability Trends</div>",
        unsafe_allow_html=True
    )

    monthly = fdf.groupby("MONTH_STR").agg({
        "NET REVENUE":"sum",
        "EBITDA%":"mean",
        "GM%":"mean"
    }).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(
        go.Bar(
            x=monthly["MONTH_STR"],
            y=monthly["NET REVENUE"]/1e7,
            name="Revenue (Cr)"
        ),
        secondary_y=False
    )

    fig.add_trace(
        go.Scatter(
            x=monthly["MONTH_STR"],
            y=monthly["EBITDA%"],
            mode="lines+markers",
            name="EBITDA %"
        ),
        secondary_y=True
    )

    fig.add_trace(
        go.Scatter(
            x=monthly["MONTH_STR"],
            y=monthly["GM%"],
            mode="lines+markers",
            name="GM %"
        ),
        secondary_y=True
    )

    fig.update_layout(
        template="plotly_white",
        height=520
    )

    st.plotly_chart(fig, use_container_width=True)

# =========================================================
# TAB 3
# =========================================================

with tab3:

    st.markdown(
        "<div class='section-header'>Market Performance Analysis</div>",
        unsafe_allow_html=True
    )

    city_perf = fdf.groupby("CITY").agg({
        "NET REVENUE":"sum",
        "EBITDA%":"mean"
    }).reset_index()

    fig2 = px.bar(
        city_perf,
        x="CITY",
        y="NET REVENUE",
        color="EBITDA%",
        text_auto=True,
        template="plotly_white"
    )

    fig2.update_layout(height=500)

    st.plotly_chart(fig2, use_container_width=True)

# =========================================================
# TAB 4
# =========================================================

with tab4:

    st.markdown(
        "<div class='section-header'>Store Intelligence Map</div>",
        unsafe_allow_html=True
    )

    store_perf = fdf.groupby(["STORE","CITY"]).agg({
        "NET REVENUE":"mean",
        "EBITDA%":"mean"
    }).reset_index()

    fig3 = px.scatter(
        store_perf,
        x="NET REVENUE",
        y="EBITDA%",
        color="CITY",
        size="NET REVENUE",
        hover_name="STORE",
        template="plotly_white",
        height=600
    )

    st.plotly_chart(fig3, use_container_width=True)

# =========================================================
# EXECUTIVE INSIGHTS
# =========================================================

st.markdown(
    "<div class='section-header'>Executive Insights</div>",
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

top_store = fdf.groupby("STORE")["EBITDA%"].mean().idxmax()

avg_var = fdf["VARIANCE%"].mean()

highest_city = fdf.groupby("CITY")["NET REVENUE"].sum().idxmax()

with col1:

    st.markdown(f"""
    <div class='insight-card'>

        <div class='insight-title'>
            TOP PERFORMING KITCHEN
        </div>

        <div class='insight-value'>
            {top_store}
        </div>

        <div class='insight-desc'>
            Highest average EBITDA margin
            among selected locations.
        </div>

    </div>
    """, unsafe_allow_html=True)

with col2:

    st.markdown(f"""
    <div class='insight-card'>

        <div class='insight-title'>
            AVERAGE VARIANCE
        </div>

        <div class='insight-value'>
            {avg_var:.2f}%
        </div>

        <div class='insight-desc'>
            Operational variance across
            selected reporting periods.
        </div>

    </div>
    """, unsafe_allow_html=True)

with col3:

    st.markdown(f"""
    <div class='insight-card'>

        <div class='insight-title'>
            HIGHEST REVENUE MARKET
        </div>

        <div class='insight-value'>
            {highest_city}
        </div>

        <div class='insight-desc'>
            Leading contribution in
            overall kitchen revenue.
        </div>

    </div>
    """, unsafe_allow_html=True)
