"""
╔══════════════════════════════════════════════════════════════╗
║         KITCHEN P&L INTELLIGENCE DASHBOARD                  ║
║         Cloud Kitchen Analytics Suite                       ║
╚══════════════════════════════════════════════════════════════╝

Run:
streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
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
# WHITE THEME CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background: #f5f7fb;
}

/* SIDEBAR */
[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e5e7eb;
}

[data-testid="stSidebar"] label {
    color: #374151 !important;
    font-weight: 600;
}

.brand {
    text-align: center;
    padding: 18px 0 24px 0;
}

.brand-title {
    color: #111827;
    font-size: 1.2rem;
    font-weight: 800;
}

.brand-sub {
    color: #6b7280;
    font-size: 0.75rem;
}

/* METRIC CARDS */
.metric-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 22px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    transition: 0.2s ease;
}

.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 22px rgba(0,0,0,0.08);
}

.metric-title {
    color: #6b7280;
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 10px;
}

.metric-value {
    color: #111827;
    font-size: 1.9rem;
    font-weight: 800;
}

.metric-delta {
    font-size: 0.78rem;
    margin-top: 5px;
    font-weight: 600;
}

.delta-pos { color: #10b981; }
.delta-neg { color: #ef4444; }

/* SECTION HEADERS */
.section-header {
    background: linear-gradient(
        90deg,
        rgba(99,102,241,0.12),
        rgba(99,102,241,0.02)
    );
    border-left: 5px solid #6366f1;
    padding: 12px 18px;
    border-radius: 0 12px 12px 0;
    margin: 28px 0 14px 0;
    color: #111827;
    font-size: 1rem;
    font-weight: 800;
}

/* FILTER PILLS */
.filter-pill {
    display: inline-block;
    background: #eef2ff;
    border: 1px solid #c7d2fe;
    border-radius: 30px;
    padding: 4px 12px;
    font-size: 0.72rem;
    color: #4338ca;
    margin: 3px;
    font-weight: 600;
}

/* DATAFRAME */
.stDataFrame {
    border-radius: 14px;
    overflow: hidden;
}

/* TABS */
button[data-baseweb="tab"] {
    font-weight: 700;
}

</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────────────────────
FONT_COLOR = "#111827"
GRID_COLOR = "rgba(0,0,0,0.08)"

PURPLE = "#6366f1"
CYAN = "#06b6d4"
EMERALD = "#10b981"
AMBER = "#f59e0b"
ROSE = "#ef4444"

COLORS = [
    PURPLE,
    CYAN,
    EMERALD,
    AMBER,
    ROSE,
    "#8b5cf6",
    "#14b8a6",
]

# ─────────────────────────────────────────────────────────────
# LAYOUT HELPERS
# ─────────────────────────────────────────────────────────────
def base_layout(**kwargs):
    return dict(
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(
            family="Inter",
            color=FONT_COLOR
        ),
        margin=dict(l=40, r=20, t=60, b=40),
        legend=dict(
            bgcolor="rgba(255,255,255,0.8)",
            bordercolor="#e5e7eb",
            borderwidth=1
        ),
        **kwargs
    )

def styled_axis(title="", fmt="", showgrid=True):
    return dict(
        title=title,
        gridcolor=GRID_COLOR if showgrid else "rgba(0,0,0,0)",
        linecolor="#d1d5db",
        tickfont=dict(color="#6b7280"),
        title_font=dict(color="#374151"),
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

    df = df.sort_values("MONTH")

    df["MONTH_STR"] = df["MONTH"].astype(str)

    def variance_bucket(v):
        if v < 15000:
            return "(a) Var < ₹15K"
        elif v < 20000:
            return "(b) Var ₹15K–20K"
        elif v < 25000:
            return "(c) Var ₹20K–25K"
        else:
            return "(d) Var > ₹25K"

    df["VAR_BUCKET"] = df["VARIANCE"].apply(variance_bucket)

    return df

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
def render_sidebar(df: pd.DataFrame):

    st.sidebar.markdown("""
    <div class='brand'>
        <div class='brand-title'>🍳 Kitchen Intelligence</div>
        <div class='brand-sub'>P&L Analytics Suite · v2.0</div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("---")

    dashboard = st.sidebar.radio(
        "📊 Select Dashboard",
        [
            "🏪 Kitchen Level PNL",
            "📉 Variance Level PNL"
        ]
    )

    st.sidebar.markdown("---")

    filters = {}

    if "Kitchen" in dashboard:

        st.sidebar.markdown("### 🎛️ PNL Filters")

        filters["month"] = st.sidebar.multiselect(
            "📅 Month",
            options=df["MONTH_STR"].unique().tolist(),
            default=df["MONTH_STR"].unique().tolist()
        )

        filters["store"] = st.sidebar.multiselect(
            "🏪 Store",
            options=sorted(df["STORE"].unique().tolist()),
            default=[]
        )

        filters["city"] = st.sidebar.multiselect(
            "🏙️ City",
            options=sorted(df["CITY"].unique().tolist()),
            default=sorted(df["CITY"].unique().tolist())
        )

        filters["zone"] = st.sidebar.multiselect(
            "🗺️ Zone",
            options=sorted(df["ZONE MAPPING"].unique().tolist()),
            default=sorted(df["ZONE MAPPING"].unique().tolist())
        )

        filters["status"] = st.sidebar.multiselect(
            "🔵 Status",
            options=df["STATUS"].unique().tolist(),
            default=df["STATUS"].unique().tolist()
        )

        st.sidebar.markdown("---")

        st.sidebar.markdown("### 📐 Financial Filters")

        rev_min = float(df["NET REVENUE"].min())
        rev_max = float(df["NET REVENUE"].max())

        filters["rev_range"] = st.sidebar.slider(
            "💰 Net Revenue",
            min_value=rev_min,
            max_value=rev_max,
            value=(rev_min, rev_max),
            format="₹%,.0f"
        )

        ebitda_min = float(df["KITCHEN EBITDA"].min())
        ebitda_max = float(df["KITCHEN EBITDA"].max())

        filters["ebitda_range"] = st.sidebar.slider(
            "📈 EBITDA",
            min_value=ebitda_min,
            max_value=ebitda_max,
            value=(ebitda_min, ebitda_max),
            format="₹%,.0f"
        )

    else:

        st.sidebar.markdown("### 📉 Variance Filters")

        filters["var_bucket"] = st.sidebar.multiselect(
            "Variance Bucket",
            options=df["VAR_BUCKET"].unique().tolist(),
            default=df["VAR_BUCKET"].unique().tolist()
        )

        filters["city"] = st.sidebar.multiselect(
            "🏙️ City",
            options=sorted(df["CITY"].unique().tolist()),
            default=sorted(df["CITY"].unique().tolist())
        )

    st.sidebar.markdown("---")

    st.sidebar.markdown("""
    <div style='text-align:center;color:#6b7280;font-size:0.72rem'>
    Data refreshes every 5 min<br>
    📊 Kitchen P&L Suite
    </div>
    """, unsafe_allow_html=True)

    return dashboard, filters

# ─────────────────────────────────────────────────────────────
# FILTERS
# ─────────────────────────────────────────────────────────────
def apply_filters(df, filters):

    mask = pd.Series(True, index=df.index)

    if "month" in filters:
        mask &= df["MONTH_STR"].isin(filters["month"])

    if "store" in filters and filters["store"]:
        mask &= df["STORE"].isin(filters["store"])

    if "city" in filters:
        mask &= df["CITY"].isin(filters["city"])

    if "zone" in filters:
        mask &= df["ZONE MAPPING"].isin(filters["zone"])

    if "status" in filters:
        mask &= df["STATUS"].isin(filters["status"])

    if "rev_range" in filters:
        mask &= df["NET REVENUE"].between(*filters["rev_range"])

    if "ebitda_range" in filters:
        mask &= df["KITCHEN EBITDA"].between(*filters["ebitda_range"])

    if "var_bucket" in filters:
        mask &= df["VAR_BUCKET"].isin(filters["var_bucket"])

    return df[mask]

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
        ("💰 Revenue", f"₹{total_rev/1e7:.2f} Cr"),
        ("📊 Avg GM%", f"{avg_gm:.1f}%"),
        ("📈 Avg EBITDA%", f"{avg_ebitda:.1f}%"),
        ("🏪 Stores", f"{stores}")
    ]

    for col, card in zip(cols, cards):
        with col:
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-title'>{card[0]}</div>
                <div class='metric-value'>{card[1]}</div>
            </div>
            """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# TREND ANALYSIS
# ─────────────────────────────────────────────────────────────
def render_trends(df):

    st.markdown(
        "<div class='section-header'>📈 Trend Analysis</div>",
        unsafe_allow_html=True
    )

    monthly = df.groupby("MONTH_STR").agg(
        Revenue=("NET REVENUE", "sum"),
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
                y=monthly["Revenue"]/1e5,
                marker_color=PURPLE,
                name="Revenue"
            )
        )

        fig.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["EBITDA"]/1e5,
                mode="lines+markers",
                line=dict(color=EMERALD, width=3),
                name="EBITDA"
            )
        )

        fig.update_layout(
            title="Revenue & EBITDA Trend",
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("₹ Lacs"),
            **base_layout(height=420)
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        fig2 = go.Figure()

        fig2.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["GM"],
                mode="lines+markers",
                line=dict(color=CYAN, width=3),
                name="GM%"
            )
        )

        fig2.add_trace(
            go.Scatter(
                x=monthly["MONTH_STR"],
                y=monthly["EBITDA_PCT"],
                mode="lines+markers",
                line=dict(color=AMBER, width=3),
                name="EBITDA%"
            )
        )

        fig2.update_layout(
            title="Margin Trend",
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("%"),
            **base_layout(height=420)
        )

        st.plotly_chart(fig2, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# CITY ANALYSIS
# ─────────────────────────────────────────────────────────────
def render_city(df):

    st.markdown(
        "<div class='section-header'>🏙️ City Performance</div>",
        unsafe_allow_html=True
    )

    city = df.groupby("CITY").agg(
        Revenue=("NET REVENUE", "sum"),
        EBITDA=("KITCHEN EBITDA", "sum")
    ).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        fig = px.bar(
            city,
            x="CITY",
            y="Revenue",
            color="CITY",
            color_discrete_sequence=COLORS,
            title="Revenue by City"
        )

        fig.update_layout(
            xaxis=styled_axis("City"),
            yaxis=styled_axis("Revenue"),
            **base_layout(height=420)
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        fig2 = go.Figure(
            go.Pie(
                labels=city["CITY"],
                values=city["EBITDA"],
                hole=0.5,
                marker_colors=COLORS
            )
        )

        fig2.update_layout(
            title="EBITDA Share",
            **base_layout(height=420)
        )

        st.plotly_chart(fig2, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# STORE SCATTER
# ─────────────────────────────────────────────────────────────
def render_store_scatter(df):

    st.markdown(
        "<div class='section-header'>🔮 Store Intelligence</div>",
        unsafe_allow_html=True
    )

    store = df.groupby(["STORE", "CITY"]).agg(
        Revenue=("NET REVENUE", "mean"),
        EBITDA=("EBITDA%", "mean")
    ).reset_index()

    fig = px.scatter(
        store,
        x="Revenue",
        y="EBITDA",
        color="CITY",
        size="Revenue",
        hover_name="STORE",
        color_discrete_sequence=COLORS,
        title="Revenue vs EBITDA%"
    )

    fig.update_layout(
        xaxis=styled_axis("Revenue"),
        yaxis=styled_axis("EBITDA%"),
        **base_layout(height=520)
    )

    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# INSIGHTS
# ─────────────────────────────────────────────────────────────
def render_insights(df):

    st.markdown("---")

    st.markdown(
        "<div class='section-header'>💡 Auto Insights</div>",
        unsafe_allow_html=True
    )

    best_city = df.groupby("CITY")["EBITDA%"].mean().idxmax()
    best_val = df.groupby("CITY")["EBITDA%"].mean().max()

    worst_city = df.groupby("CITY")["EBITDA%"].mean().idxmin()
    worst_val = df.groupby("CITY")["EBITDA%"].mean().min()

    high_var = df[df["VARIANCE%"] > 3]["STORE"].nunique()

    c1, c2, c3 = st.columns(3)

    with c1:
        st.success(
            f"🏆 Best City: {best_city} ({best_val:.1f}% EBITDA)"
        )

    with c2:
        st.error(
            f"⚠️ Weakest City: {worst_city} ({worst_val:.1f}% EBITDA)"
        )

    with c3:
        st.info(
            f"📉 High Variance Stores: {high_var}"
        )

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():

    st.markdown("""
    <div style='text-align:center;padding:10px 0 25px 0'>
        <h1 style='color:#111827;font-size:2.4rem;font-weight:900'>
            🍳 Kitchen P&L Intelligence Suite
        </h1>

        <p style='color:#6b7280;font-size:0.9rem'>
            CLOUD KITCHEN ANALYTICS · MULTI-CITY · REAL-TIME FILTERS
        </p>
    </div>
    """, unsafe_allow_html=True)

    with st.spinner("Loading kitchen analytics..."):
        df = load_data(DATA_PATH)

    dashboard, filters = render_sidebar(df)

    filtered_df = apply_filters(df, filters)

    if filtered_df.empty:
        st.warning("No matching data found.")
        return

    render_kpis(filtered_df)

    tabs = st.tabs([
        "📈 Trends",
        "🏙️ Cities",
        "🔮 Store Intelligence",
        "📉 Variance"
    ])

    with tabs[0]:
        render_trends(filtered_df)

    with tabs[1]:
        render_city(filtered_df)

    with tabs[2]:
        render_store_scatter(filtered_df)

    with tabs[3]:

        variance = filtered_df.groupby(
            ["MONTH_STR", "VAR_BUCKET"]
        ).size().reset_index(name="COUNT")

        fig = px.bar(
            variance,
            x="MONTH_STR",
            y="COUNT",
            color="VAR_BUCKET",
            barmode="stack",
            title="Variance Distribution"
        )

        fig.update_layout(
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("Store Count"),
            **base_layout(height=500)
        )

        st.plotly_chart(fig, use_container_width=True)

    render_insights(filtered_df)

if __name__ == "__main__":
    main()
