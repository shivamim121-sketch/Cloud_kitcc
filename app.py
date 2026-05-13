"""
╔══════════════════════════════════════════════════════════════╗
║         KITCHEN P&L INTELLIGENCE DASHBOARD                  ║
║         Cloud Kitchen Analytics Suite                       ║
╚══════════════════════════════════════════════════════════════╝

Python: 3.10+
Packages: streamlit>=1.32, plotly>=5.20, pandas>=2.0, openpyxl>=3.1
Run: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
import hashlib
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# PATH CONFIG
# ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────
# THEME / CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: linear-gradient(135deg, #0f0f1a 0%, #1a1a2e 50%, #16213e 100%);
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d0d1a 0%, #1a0a2e 100%);
    border-right: 1px solid rgba(139,92,246,0.3);
}

/* Metric Cards */
.metric-card {
    background: linear-gradient(
        135deg,
        rgba(139,92,246,0.15) 0%,
        rgba(59,130,246,0.10) 100%
    );
    border: 1px solid rgba(139,92,246,0.4);
    border-radius: 16px;
    padding: 20px;
    text-align: center;
    backdrop-filter: blur(10px);
    margin-bottom: 10px;
}

.metric-title {
    color: #94a3b8;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.metric-value {
    color: white;
    font-size: 1.7rem;
    font-weight: 800;
}

.metric-delta {
    margin-top: 5px;
    font-size: 0.75rem;
    font-weight: 600;
}

.delta-pos { color: #34d399; }
.delta-neg { color: #f87171; }

.section-header {
    background: linear-gradient(
        90deg,
        rgba(139,92,246,0.2),
        transparent
    );
    border-left: 4px solid #8b5cf6;
    padding: 10px 16px;
    border-radius: 0 8px 8px 0;
    margin: 20px 0 12px 0;
    color: #c4b5fd;
    font-weight: 700;
    font-size: 1rem;
}

.filter-pill {
    display: inline-block;
    background: rgba(139,92,246,0.2);
    border: 1px solid rgba(139,92,246,0.5);
    border-radius: 20px;
    padding: 3px 12px;
    font-size: 0.72rem;
    color: #c4b5fd;
    margin: 2px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────────────────────
CHART_BG = "rgba(0,0,0,0)"
PAPER_BG = "rgba(0,0,0,0)"
FONT_COLOR = "#e2e8f0"
GRID_COLOR = "rgba(139,92,246,0.15)"

PURPLE = "#8b5cf6"
CYAN = "#22d3ee"
EMERALD = "#34d399"
AMBER = "#fbbf24"
ROSE = "#f87171"

COLORS = [
    PURPLE,
    CYAN,
    EMERALD,
    AMBER,
    ROSE,
    "#a78bfa",
    "#67e8f9",
    "#6ee7b7",
]

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def base_layout(**kwargs):
    return dict(
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=CHART_BG,
        font=dict(family="Inter", color=FONT_COLOR),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(
            bgcolor="rgba(0,0,0,0.3)",
            bordercolor="rgba(139,92,246,0.3)",
            borderwidth=1,
        ),
        **kwargs,
    )

def styled_axis(title="", fmt="", showgrid=True):
    return dict(
        title=title,
        gridcolor=GRID_COLOR if showgrid else "rgba(0,0,0,0)",
        gridwidth=1,
        linecolor="rgba(139,92,246,0.3)",
        tickfont=dict(color="#94a3b8", size=11),
        title_font=dict(color="#c4b5fd", size=12),
        tickformat=fmt,
        zeroline=False,
    )

# ─────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data(filepath):
    df = pd.read_excel(filepath, header=1)

    df["GM%"] = (
        df["GROSS MARGIN"] / df["NET REVENUE"] * 100
    ).round(2)

    df["EBITDA%"] = (
        df["KITCHEN EBITDA"] / df["NET REVENUE"] * 100
    ).round(2)

    df["VARIANCE%"] = (
        df["VARIANCE"] / df["NET REVENUE"] * 100
    ).round(4)

    df["CM"] = df["GROSS MARGIN"]
    df["CM%"] = df["GM%"]

    month_order = [
        "Oct-2023",
        "Nov-2023",
        "Dec-2023",
        "Jan-2024",
        "Feb-2024",
        "Mar-2024",
    ]

    df["MONTH"] = pd.Categorical(
        df["MONTH"],
        categories=month_order,
        ordered=True,
    )

    df = df.sort_values("MONTH")
    df["MONTH_STR"] = df["MONTH"].astype(str)

    def var_bucket(v):
        if v < 15000:
            return "(a) Var < ₹15K"
        elif v < 20000:
            return "(b) Var ₹15K–20K"
        elif v < 25000:
            return "(c) Var ₹20K–25K"
        else:
            return "(d) Var > ₹25K"

    df["VAR_BUCKET"] = df["VARIANCE"].apply(var_bucket)

    def rev_bucket(r):
        lacs = r / 1e5

        if lacs < 15:
            return "(a) Below INR 15 lacs"
        elif lacs < 25:
            return "(b) INR 15 to 25 lacs"
        elif lacs < 35:
            return "(c) INR 25 to 35 lacs"
        elif lacs < 45:
            return "(d) INR 35 to 45 lacs"
        else:
            return "(e) Above INR 45 lacs"

    df["REV_BUCKET"] = df["NET REVENUE"].apply(rev_bucket)

    return df

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
def render_sidebar(df):

    st.sidebar.markdown("""
    <div style='text-align:center;padding:15px 0'>
        <h2 style='color:white;'>🍳 Kitchen Intelligence</h2>
        <p style='color:#888;'>P&L Analytics Suite</p>
    </div>
    """, unsafe_allow_html=True)

    dashboard = st.sidebar.radio(
        "Select Dashboard",
        [
            "🏪 Kitchen Level PNL",
            "📉 Variance Level PNL"
        ]
    )

    filters = {}

    filters["month"] = st.sidebar.multiselect(
        "Month",
        options=df["MONTH_STR"].unique().tolist(),
        default=df["MONTH_STR"].unique().tolist()
    )

    filters["city"] = st.sidebar.multiselect(
        "City",
        options=sorted(df["CITY"].unique().tolist()),
        default=sorted(df["CITY"].unique().tolist())
    )

    filters["store"] = st.sidebar.multiselect(
        "Store",
        options=sorted(df["STORE"].unique().tolist()),
        default=[]
    )

    return dashboard, filters

# ─────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────
def apply_filters(df, filters):

    mask = pd.Series(True, index=df.index)

    if filters["month"]:
        mask &= df["MONTH_STR"].isin(filters["month"])

    if filters["city"]:
        mask &= df["CITY"].isin(filters["city"])

    if filters["store"]:
        mask &= df["STORE"].isin(filters["store"])

    return df[mask]

# ─────────────────────────────────────────────────────────────
# KPI STRIP
# ─────────────────────────────────────────────────────────────
def render_kpi_strip(df):

    total_rev = df["NET REVENUE"].sum()
    avg_gm = df["GM%"].mean()
    avg_ebitda = df["EBITDA%"].mean()
    total_ebitda = df["KITCHEN EBITDA"].sum()

    cols = st.columns(4)

    metrics = [
        ("💰 Net Revenue", f"₹{total_rev/1e7:.2f} Cr"),
        ("📈 Avg GM%", f"{avg_gm:.1f}%"),
        ("📊 Avg EBITDA%", f"{avg_ebitda:.1f}%"),
        ("🏦 Total EBITDA", f"₹{total_ebitda/1e7:.2f} Cr"),
    ]

    for col, metric in zip(cols, metrics):
        with col:
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-title'>{metric[0]}</div>
                <div class='metric-value'>{metric[1]}</div>
            </div>
            """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# TREND CHARTS
# ─────────────────────────────────────────────────────────────
def render_trends(df):

    st.markdown(
        "<div class='section-header'>📈 Trend Analysis</div>",
        unsafe_allow_html=True
    )

    monthly = df.groupby("MONTH_STR").agg(
        NET_REVENUE=("NET REVENUE", "sum"),
        EBITDA=("KITCHEN EBITDA", "sum"),
        GM=("GM%", "mean"),
        EBITDA_PCT=("EBITDA%", "mean")
    ).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        fig = go.Figure()

        fig.add_trace(
            go.Bar(
                x=monthly["MONTH_STR"],
                y=monthly["NET_REVENUE"] / 1e5,
                name="Revenue",
                marker_color=PURPLE,
            )
        )

        fig.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["EBITDA"] / 1e5,
                name="EBITDA",
                mode="lines+markers",
                line=dict(color=EMERALD, width=3),
            )
        )

        fig.update_layout(
            title="Revenue & EBITDA Trend",
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("₹ Lacs"),
            **base_layout()
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        fig2 = go.Figure()

        fig2.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["GM"],
                name="GM%",
                mode="lines+markers",
                line=dict(color=CYAN, width=3),
            )
        )

        fig2.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["EBITDA_PCT"],
                name="EBITDA%",
                mode="lines+markers",
                line=dict(color=AMBER, width=3),
            )
        )

        fig2.update_layout(
            title="Margin Trends",
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("Margin %"),
            **base_layout()
        )

        st.plotly_chart(fig2, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# CITY ANALYSIS
# ─────────────────────────────────────────────────────────────
def render_city_analysis(df):

    st.markdown(
        "<div class='section-header'>🏙️ City Analysis</div>",
        unsafe_allow_html=True
    )

    city = df.groupby("CITY").agg(
        REVENUE=("NET REVENUE", "sum"),
        EBITDA=("KITCHEN EBITDA", "sum"),
    ).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        fig = px.bar(
            city,
            x="CITY",
            y="REVENUE",
            color="CITY",
            color_discrete_sequence=COLORS,
            title="Revenue by City"
        )

        fig.update_layout(
            xaxis=styled_axis("City"),
            yaxis=styled_axis("Revenue"),
            **base_layout()
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        fig2 = go.Figure(
            go.Pie(
                labels=city["CITY"],
                values=city["EBITDA"],
                hole=0.5,
                marker_colors=COLORS,
            )
        )

        fig2.update_layout(
            title="EBITDA Share by City",
            **base_layout()
        )

        st.plotly_chart(fig2, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# STORE SCATTER
# ─────────────────────────────────────────────────────────────
def render_scatter(df):

    st.markdown(
        "<div class='section-header'>🔮 Store Intelligence</div>",
        unsafe_allow_html=True
    )

    store = df.groupby(["STORE", "CITY"]).agg(
        NET_REVENUE=("NET REVENUE", "mean"),
        EBITDA_PCT=("EBITDA%", "mean"),
    ).reset_index()

    fig = px.scatter(
        store,
        x="NET_REVENUE",
        y="EBITDA_PCT",
        color="CITY",
        size="NET_REVENUE",
        hover_name="STORE",
        color_discrete_sequence=COLORS,
        title="Revenue vs EBITDA%"
    )

    fig.update_layout(
        xaxis=styled_axis("Revenue"),
        yaxis=styled_axis("EBITDA%"),
        **base_layout(height=500)
    )

    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# RANKINGS
# ─────────────────────────────────────────────────────────────
def render_rankings(df):

    st.markdown(
        "<div class='section-header'>🏆 Store Rankings</div>",
        unsafe_allow_html=True
    )

    perf = df.groupby("STORE").agg(
        EBITDA=("EBITDA%", "mean")
    ).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        top = perf.nlargest(10, "EBITDA")

        fig = go.Figure(
            go.Bar(
                x=top["EBITDA"],
                y=top["STORE"],
                orientation="h",
                marker_color=EMERALD,
            )
        )

        fig.update_layout(
            title="Top Stores",
            yaxis=dict(autorange="reversed"),
            **base_layout(height=400)
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        bottom = perf.nsmallest(10, "EBITDA")

        fig2 = go.Figure(
            go.Bar(
                x=bottom["EBITDA"],
                y=bottom["STORE"],
                orientation="h",
                marker_color=ROSE,
            )
        )

        fig2.update_layout(
            title="Bottom Stores",
            yaxis=dict(autorange="reversed"),
            **base_layout(height=400)
        )

        st.plotly_chart(fig2, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# VARIANCE DASHBOARD
# ─────────────────────────────────────────────────────────────
def render_variance_dashboard(df):

    st.markdown(
        "<div class='section-header'>📉 Variance Dashboard</div>",
        unsafe_allow_html=True
    )

    var = df.groupby(["MONTH_STR", "VAR_BUCKET"]).size().reset_index(name="COUNT")

    fig = px.bar(
        var,
        x="MONTH_STR",
        y="COUNT",
        color="VAR_BUCKET",
        barmode="stack",
        color_discrete_map={
            "(a) Var < ₹15K": EMERALD,
            "(b) Var ₹15K–20K": AMBER,
            "(c) Var ₹20K–25K": PURPLE,
            "(d) Var > ₹25K": ROSE,
        },
        title="Variance Distribution"
    )

    fig.update_layout(
        xaxis=styled_axis("Month"),
        yaxis=styled_axis("Store Count"),
        **base_layout(height=500)
    )

    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# INSIGHTS
# ─────────────────────────────────────────────────────────────
def render_insights(df):

    st.markdown("---")

    st.markdown(
        "<div class='section-header'>💡 Insights</div>",
        unsafe_allow_html=True
    )

    best_city = df.groupby("CITY")["EBITDA%"].mean().idxmax()
    worst_city = df.groupby("CITY")["EBITDA%"].mean().idxmin()

    col1, col2 = st.columns(2)

    with col1:
        st.success(f"🏆 Best Performing City: {best_city}")

    with col2:
        st.error(f"⚠️ Lowest Performing City: {worst_city}")

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():

    st.markdown("""
    <div style='text-align:center;padding:10px 0 25px 0'>
        <h1 style='color:white;font-size:2.2rem;font-weight:800'>
            🍳 Kitchen P&L Intelligence Suite
        </h1>
        <p style='color:#888'>
            CLOUD KITCHEN ANALYTICS · MULTI-CITY · REAL-TIME FILTERS
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Loading data..."):
        df = load_data(DATA_PATH)

    dashboard, filters = render_sidebar(df)

    filtered_df = apply_filters(df, filters)

    if filtered_df.empty:
        st.warning("No matching data found.")
        return

    render_kpi_strip(filtered_df)

    tabs = st.tabs([
        "📈 Trends",
        "🏙️ City",
        "🔮 Stores",
        "🏆 Rankings",
        "📉 Variance"
    ])

    with tabs[0]:
        render_trends(filtered_df)

    with tabs[1]:
        render_city_analysis(filtered_df)

    with tabs[2]:
        render_scatter(filtered_df)

    with tabs[3]:
        render_rankings(filtered_df)

    with tabs[4]:
        render_variance_dashboard(filtered_df)

    render_insights(filtered_df)

if __name__ == "__main__":
    main()
