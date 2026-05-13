"""
╔══════════════════════════════════════════════════════════════╗
║         KITCHEN P&L INTELLIGENCE DASHBOARD                  ║
║         Enterprise Cloud Kitchen Analytics Suite            ║
╚══════════════════════════════════════════════════════════════╝

Python: 3.10+
Packages:
streamlit>=1.32
plotly>=5.20
pandas>=2.0
numpy>=1.24
openpyxl>=3.1

Run:
streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from plotly.subplots import make_subplots
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────
# DATA PATH
# ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

# ─────────────────────────────────────────────────────────────
# PREMIUM WHITE THEME
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>

/* ========= GLOBAL ========= */

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"]{
    font-family:'Inter', sans-serif;
}

.stApp{
    background:
        radial-gradient(circle at top left, rgba(99,102,241,0.05), transparent 30%),
        radial-gradient(circle at bottom right, rgba(14,165,233,0.05), transparent 30%),
        #f8fafc;
}

/* ========= SIDEBAR ========= */

[data-testid="stSidebar"]{
    background:white;
    border-right:1px solid #e5e7eb;
}

[data-testid="stSidebar"] label{
    color:#374151 !important;
    font-weight:600 !important;
}

.brand{
    padding-top:10px;
    text-align:center;
}

.brand-title{
    font-size:1.4rem;
    font-weight:900;
    color:#111827;
}

.brand-sub{
    color:#6b7280;
    font-size:0.78rem;
    margin-top:4px;
}

/* ========= HERO ========= */

.hero-container{
    position:relative;
    overflow:hidden;
    padding:42px;
    border-radius:30px;
    background:
        linear-gradient(
            135deg,
            #4f46e5 0%,
            #7c3aed 35%,
            #2563eb 70%,
            #0891b2 100%
        );

    box-shadow:
        0 25px 60px rgba(79,70,229,0.25),
        inset 0 1px 0 rgba(255,255,255,0.2);

    margin-bottom:28px;
}

.hero-container::before{
    content:'';
    position:absolute;
    width:520px;
    height:520px;
    top:-240px;
    right:-120px;
    background:
        radial-gradient(circle,
        rgba(255,255,255,0.22),
        transparent 70%);
}

.hero-grid{
    display:flex;
    justify-content:space-between;
    align-items:center;
    gap:30px;
    position:relative;
    z-index:2;
}

.hero-title{
    color:white;
    font-size:3rem;
    font-weight:900;
    line-height:1;
    letter-spacing:-0.05em;
}

.hero-title span{
    color:#dbeafe;
}

.hero-sub{
    color:rgba(255,255,255,0.92);
    margin-top:18px;
    font-size:1rem;
    line-height:1.8;
    max-width:850px;
}

.hero-badge{
    display:inline-block;
    padding:8px 16px;
    border-radius:999px;
    background:rgba(255,255,255,0.14);
    color:white;
    font-size:0.76rem;
    font-weight:700;
    letter-spacing:0.08em;
    margin-bottom:18px;
    border:1px solid rgba(255,255,255,0.18);
}

.hero-chips{
    display:flex;
    gap:10px;
    flex-wrap:wrap;
    margin-top:20px;
}

.hero-chip{
    padding:8px 14px;
    border-radius:999px;
    background:rgba(255,255,255,0.12);
    color:white;
    font-size:0.78rem;
    font-weight:600;
    border:1px solid rgba(255,255,255,0.14);
}

.hero-card{
    min-width:320px;
    background:rgba(255,255,255,0.12);
    border:1px solid rgba(255,255,255,0.14);
    border-radius:24px;
    padding:24px;
    backdrop-filter:blur(8px);
}

.hero-card-title{
    color:#e0e7ff;
    font-size:0.75rem;
    font-weight:700;
    letter-spacing:0.08em;
}

.hero-card-metric{
    color:white;
    font-size:2.4rem;
    font-weight:900;
    margin-top:12px;
}

.hero-card-desc{
    color:rgba(255,255,255,0.88);
    margin-top:10px;
    line-height:1.7;
    font-size:0.92rem;
}

/* ========= KPI CARDS ========= */

.metric-card{
    background:white;
    border:1px solid #e5e7eb;
    border-radius:20px;
    padding:22px;
    box-shadow:0 8px 24px rgba(15,23,42,0.05);
    transition:0.25s ease;
}

.metric-card:hover{
    transform:translateY(-4px);
    box-shadow:0 16px 40px rgba(79,70,229,0.12);
}

.metric-title{
    color:#6b7280;
    font-size:0.75rem;
    text-transform:uppercase;
    font-weight:700;
    letter-spacing:0.08em;
}

.metric-value{
    color:#111827;
    font-size:1.9rem;
    font-weight:900;
    margin-top:10px;
}

.metric-delta{
    margin-top:8px;
    font-size:0.8rem;
    font-weight:600;
}

.delta-pos{
    color:#059669;
}

.delta-neg{
    color:#dc2626;
}

/* ========= SECTION ========= */

.section-header{
    margin-top:28px;
    margin-bottom:16px;
    font-size:1.05rem;
    font-weight:800;
    color:#111827;
    padding-left:14px;
    border-left:5px solid #4f46e5;
}

/* ========= TABS ========= */

button[data-baseweb="tab"]{
    font-weight:700;
}

/* ========= DATAFRAME ========= */

.stDataFrame{
    border-radius:18px;
    overflow:hidden;
    border:1px solid #e5e7eb;
}

/* ========= INSIGHTS ========= */

.insight-box{
    background:white;
    border:1px solid #e5e7eb;
    border-radius:20px;
    padding:24px;
    box-shadow:0 10px 28px rgba(15,23,42,0.05);
}

.insight-title{
    font-size:0.75rem;
    color:#6b7280;
    font-weight:700;
    letter-spacing:0.08em;
    text-transform:uppercase;
}

.insight-value{
    margin-top:10px;
    font-size:2rem;
    font-weight:900;
    color:#111827;
}

.insight-desc{
    margin-top:10px;
    color:#4b5563;
    line-height:1.7;
}

/* ========= FILTER PILL ========= */

.filter-pill{
    display:inline-block;
    background:#eef2ff;
    color:#4338ca;
    border-radius:999px;
    padding:6px 14px;
    margin:4px;
    font-size:0.76rem;
    font-weight:700;
}

/* ========= HIDE STREAMLIT ========= */

#MainMenu {visibility:hidden;}
footer {visibility:hidden;}

</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data(path):

    df = pd.read_excel(path, header=1)

    df["GM%"] = (
        df["GROSS MARGIN"] / df["NET REVENUE"] * 100
    ).round(2)

    df["EBITDA%"] = (
        df["KITCHEN EBITDA"] / df["NET REVENUE"] * 100
    ).round(2)

    df["VARIANCE%"] = (
        df["VARIANCE"] / df["NET REVENUE"] * 100
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

# ─────────────────────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────────────────────
PURPLE = "#4f46e5"
CYAN = "#0891b2"
GREEN = "#059669"
RED = "#dc2626"
AMBER = "#d97706"

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
def render_sidebar(df):

    st.sidebar.markdown("""
    <div class='brand'>
        <div class='brand-title'>🍳 Kitchen Intelligence</div>
        <div class='brand-sub'>
            Financial Performance Analytics
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("---")

    dashboard = st.sidebar.radio(
        "Dashboard",
        [
            "🏪 Kitchen Level PNL",
            "📉 Variance Level PNL"
        ]
    )

    st.sidebar.markdown("---")

    filters = {}

    if "Kitchen" in dashboard:

        st.sidebar.markdown("### 🎛️ Operational Filters")

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

        filters["store"] = st.sidebar.multiselect(
            "🏪 Store",
            options=sorted(df["STORE"].unique())
        )

        filters["status"] = st.sidebar.multiselect(
            "🔵 Status",
            options=df["STATUS"].unique(),
            default=df["STATUS"].unique()
        )

    else:

        st.sidebar.markdown("### 📉 Variance Filters")

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

    st.sidebar.markdown("---")

    st.sidebar.markdown("""
    <div style='text-align:center;color:#6b7280;font-size:0.75rem'>
    Updated operational reporting dashboard
    </div>
    """, unsafe_allow_html=True)

    return dashboard, filters

# ─────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────
def apply_filters(df, filters):

    fdf = df.copy()

    if filters.get("month"):
        fdf = fdf[
            fdf["MONTH_STR"].isin(filters["month"])
        ]

    if filters.get("city"):
        fdf = fdf[
            fdf["CITY"].isin(filters["city"])
        ]

    if filters.get("store"):
        fdf = fdf[
            fdf["STORE"].isin(filters["store"])
        ]

    if filters.get("status"):
        fdf = fdf[
            fdf["STATUS"].isin(filters["status"])
        ]

    return fdf

# ─────────────────────────────────────────────────────────────
# KPI STRIP
# ─────────────────────────────────────────────────────────────
def render_kpis(df):

    total_rev = df["NET REVENUE"].sum()
    avg_gm = df["GM%"].mean()
    avg_ebitda = df["EBITDA%"].mean()
    stores = df["STORE"].nunique()

    cols = st.columns(4)

    cards = [
        (
            "💰 Net Revenue",
            f"₹{total_rev/1e7:.1f}Cr",
            "Across selected kitchens"
        ),
        (
            "📊 Avg GM%",
            f"{avg_gm:.1f}%",
            "Gross margin performance"
        ),
        (
            "📈 Avg EBITDA%",
            f"{avg_ebitda:.1f}%",
            "Operational profitability"
        ),
        (
            "🏪 Kitchens",
            f"{stores}",
            "Active locations"
        )
    ]

    for col, card in zip(cols, cards):

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

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():

    df = load_data(DATA_PATH)

    dashboard, filters = render_sidebar(df)

    fdf = apply_filters(df, filters)

    total_rev = fdf["NET REVENUE"].sum() / 1e7
    total_stores = fdf["STORE"].nunique()
    avg_ebitda = fdf["EBITDA%"].mean()

    best_city = (
        fdf.groupby("CITY")["EBITDA%"]
        .mean()
        .idxmax()
    )

    # ───────────────── HERO ─────────────────

    st.markdown(f"""
    <div class='hero-container'>

        <div class='hero-grid'>

            <div>

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

                    <div class='hero-chip'>
                        📊 Revenue Monitoring
                    </div>

                    <div class='hero-chip'>
                        📈 EBITDA Tracking
                    </div>

                    <div class='hero-chip'>
                        🏪 Multi-City Analysis
                    </div>

                    <div class='hero-chip'>
                        📉 Variance Reporting
                    </div>

                    <div class='hero-chip'>
                        ⚡ Operational Intelligence
                    </div>

                </div>

            </div>

            <div class='hero-card'>

                <div class='hero-card-title'>
                    PERFORMANCE SNAPSHOT
                </div>

                <div class='hero-card-metric'>
                    ₹{total_rev:.1f}Cr
                </div>

                <div class='hero-card-desc'>

                    Financial review across
                    <b>{total_stores}</b> kitchens with
                    average EBITDA margin of
                    <b>{avg_ebitda:.1f}%</b>.

                    <br><br>

                    Highest operational performance observed in
                    <b>{best_city}</b>.

                </div>

            </div>

        </div>

    </div>
    """, unsafe_allow_html=True)

    # ───────────────── KPIs ─────────────────

    render_kpis(fdf)

    # ───────────────── FILTER PILL ─────────────────

    st.markdown("<br>", unsafe_allow_html=True)

    active_filters = []

    if filters.get("city"):
        active_filters.append(
            f"Cities: {len(filters['city'])}"
        )

    if filters.get("month"):
        active_filters.append(
            f"Months: {len(filters['month'])}"
        )

    pills = "".join([
        f"<span class='filter-pill'>{x}</span>"
        for x in active_filters
    ])

    st.markdown(
        f"<div>{pills}</div>",
        unsafe_allow_html=True
    )

    # ───────────────── TABS ─────────────────

    tab1, tab2, tab3 = st.tabs([
        "📋 Performance Table",
        "📈 Trends",
        "🏙️ Market Analysis"
    ])

    # ───────────────── TABLE ─────────────────

    with tab1:

        st.markdown("""
        <div class='section-header'>
            Kitchen Performance Overview
        </div>
        """, unsafe_allow_html=True)

        table = fdf[[
            "STORE",
            "CITY",
            "MONTH_STR",
            "NET REVENUE",
            "GM%",
            "EBITDA%",
            "VARIANCE%"
        ]].copy()

        st.dataframe(
            table,
            use_container_width=True,
            height=500
        )

    # ───────────────── TRENDS ─────────────────

    with tab2:

        st.markdown("""
        <div class='section-header'>
            Revenue & EBITDA Trend
        </div>
        """, unsafe_allow_html=True)

        trend = (
            fdf.groupby("MONTH_STR")
            .agg({
                "NET REVENUE":"sum",
                "KITCHEN EBITDA":"sum"
            })
            .reset_index()
        )

        fig = go.Figure()

        fig.add_trace(
            go.Bar(
                x=trend["MONTH_STR"],
                y=trend["NET REVENUE"]/1e5,
                name="Revenue (₹L)",
                marker_color=PURPLE
            )
        )

        fig.add_trace(
            go.Scatter(
                x=trend["MONTH_STR"],
                y=trend["KITCHEN EBITDA"]/1e5,
                mode="lines+markers",
                name="EBITDA (₹L)",
                line=dict(
                    color=GREEN,
                    width=4
                )
            )
        )

        fig.update_layout(
            height=500,
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(color="#111827"),
            xaxis=dict(
                title="Month",
                gridcolor="#f3f4f6"
            ),
            yaxis=dict(
                title="₹ Lacs",
                gridcolor="#f3f4f6"
            )
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # ───────────────── CITY ANALYSIS ─────────────────

    with tab3:

        st.markdown("""
        <div class='section-header'>
            City Performance Distribution
        </div>
        """, unsafe_allow_html=True)

        city = (
            fdf.groupby("CITY")
            .agg({
                "NET REVENUE":"sum",
                "EBITDA%":"mean"
            })
            .reset_index()
        )

        fig2 = px.bar(
            city,
            x="CITY",
            y="NET REVENUE",
            color="EBITDA%",
            color_continuous_scale="Blues"
        )

        fig2.update_layout(
            height=500,
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(color="#111827")
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

    # ───────────────── INSIGHTS ─────────────────

    st.markdown("""
    <div class='section-header'>
        Executive Insights
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:

        best_store = (
            fdf.groupby("STORE")["EBITDA%"]
            .mean()
            .idxmax()
        )

        st.markdown(f"""
        <div class='insight-box'>

            <div class='insight-title'>
                TOP PERFORMING KITCHEN
            </div>

            <div class='insight-value'>
                {best_store}
            </div>

            <div class='insight-desc'>
                Highest average EBITDA margin
                among selected locations.
            </div>

        </div>
        """, unsafe_allow_html=True)

    with c2:

        avg_var = fdf["VARIANCE%"].mean()

        st.markdown(f"""
        <div class='insight-box'>

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

    with c3:

        top_city = (
            fdf.groupby("CITY")["NET REVENUE"]
            .sum()
            .idxmax()
        )

        st.markdown(f"""
        <div class='insight-box'>

            <div class='insight-title'>
                HIGHEST REVENUE MARKET
            </div>

            <div class='insight-value'>
                {top_city}
            </div>

            <div class='insight-desc'>
                Leading contribution in
                overall kitchen revenue.
            </div>

        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# RUN APP
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
