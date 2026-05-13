import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
from pathlib import Path

# ═══════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════
# PREMIUM UI CSS
# ═══════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(139,92,246,0.18), transparent 30%),
        radial-gradient(circle at bottom right, rgba(34,211,238,0.15), transparent 30%),
        linear-gradient(135deg, #050816 0%, #0b1120 45%, #111827 100%);
    color: white;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #070b1d 0%, #111827 100%);
    border-right: 1px solid rgba(139,92,246,0.25);
}

.metric-card {
    background: rgba(17,24,39,0.75);
    border: 1px solid rgba(139,92,246,0.25);
    border-radius: 22px;
    padding: 22px;
    backdrop-filter: blur(16px);
    transition: all 0.25s ease;
    box-shadow: 0 4px 20px rgba(0,0,0,0.25);
}

.metric-card:hover {
    transform: translateY(-4px);
    border-color: rgba(139,92,246,0.6);
    box-shadow: 0 10px 30px rgba(139,92,246,0.25);
}

.metric-title {
    color: #94a3b8;
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}

.metric-value {
    color: white;
    font-size: 2rem;
    font-weight: 800;
    margin-top: 8px;
}

.metric-delta {
    margin-top: 8px;
    font-size: 0.78rem;
    font-weight: 600;
}

.delta-pos {
    color: #4ade80;
}

.delta-neg {
    color: #fb7185;
}

.section-header {
    background: linear-gradient(90deg,
        rgba(139,92,246,0.28),
        rgba(34,211,238,0.08),
        transparent);
    border-left: 5px solid #8b5cf6;
    padding: 14px 18px;
    border-radius: 14px;
    margin: 24px 0 16px 0;
    color: white;
    font-weight: 800;
    letter-spacing: 0.05em;
}

.stDataFrame {
    border: 1px solid rgba(139,92,246,0.2);
    border-radius: 18px;
    overflow: hidden;
}

.filter-pill {
    display: inline-block;
    background: rgba(139,92,246,0.18);
    border: 1px solid rgba(139,92,246,0.35);
    color: #ddd6fe;
    padding: 6px 14px;
    border-radius: 30px;
    font-size: 0.72rem;
    margin-right: 8px;
}

</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# COLORS
# ═══════════════════════════════════════════════════════════════
PURPLE = "#8b5cf6"
CYAN = "#22d3ee"
EMERALD = "#34d399"
AMBER = "#fbbf24"
ROSE = "#fb7185"
FONT_COLOR = "#f8fafc"
GRID_COLOR = "rgba(139,92,246,0.15)"
COLORS = [PURPLE, CYAN, EMERALD, AMBER, ROSE]

# ═══════════════════════════════════════════════════════════════
# FILE PATH
# ═══════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

if not DATA_PATH.exists():
    st.error(f"Excel file not found: {DATA_PATH}")
    st.stop()

# ═══════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════
@st.cache_data(ttl=600)
def load_data(filepath):
    df = pd.read_excel(filepath, header=1)

    df["GM%"] = (df["GROSS MARGIN"] / df["NET REVENUE"] * 100).round(2)
    df["EBITDA%"] = (df["KITCHEN EBITDA"] / df["NET REVENUE"] * 100).round(2)
    df["VARIANCE%"] = (df["VARIANCE"] / df["NET REVENUE"] * 100).round(2)

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

    return df

# ═══════════════════════════════════════════════════════════════
# SAFE DATAFRAME STYLING
# ═══════════════════════════════════════════════════════════════
def safe_style_dataframe(df, cols=None):

    def color_vals(v):
        if pd.isna(v):
            return ""

        try:
            v = float(v)
        except:
            return ""

        if v < 0:
            return "color:#fb7185;font-weight:700"
        elif v > 0:
            return "color:#4ade80;font-weight:700"

        return ""

    styler = df.style

    if cols:
        try:
            styler = styler.map(color_vals, subset=cols)
        except:
            try:
                styler = styler.applymap(color_vals, subset=cols)
            except:
                pass

    return styler

# ═══════════════════════════════════════════════════════════════
# CHART LAYOUT
# ═══════════════════════════════════════════════════════════════
def base_layout(**kwargs):
    return dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0.02)",
        font=dict(
            family="Inter",
            color=FONT_COLOR,
            size=12
        ),
        margin=dict(l=40, r=30, t=60, b=40),
        legend=dict(
            bgcolor="rgba(17,24,39,0.85)",
            bordercolor="rgba(139,92,246,0.25)",
            borderwidth=1,
        ),
        **kwargs
    )

# ═══════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════
def render_sidebar(df):

    st.sidebar.markdown("## 🍳 Kitchen Intelligence")

    months = st.sidebar.multiselect(
        "Month",
        df["MONTH"].dropna().unique(),
        default=df["MONTH"].dropna().unique()
    )

    cities = st.sidebar.multiselect(
        "City",
        sorted(df["CITY"].dropna().unique()),
        default=sorted(df["CITY"].dropna().unique())
    )

    filtered = df[
        (df["MONTH"].isin(months)) &
        (df["CITY"].isin(cities))
    ]

    return filtered

# ═══════════════════════════════════════════════════════════════
# KPI CARDS
# ═══════════════════════════════════════════════════════════════
def metric_card(title, value, delta, positive=True):

    dclass = "delta-pos" if positive else "delta-neg"

    return f"""
    <div class='metric-card'>
        <div class='metric-title'>{title}</div>
        <div class='metric-value'>{value}</div>
        <div class='metric-delta {dclass}'>{delta}</div>
    </div>
    """

def render_kpis(df):

    total_rev = df["NET REVENUE"].sum()
    total_ebitda = df["KITCHEN EBITDA"].sum()
    avg_gm = df["GM%"].mean()
    avg_ebitda = df["EBITDA%"].mean()
    stores = df["STORE"].nunique()

    cols = st.columns(5)

    cards = [
        ("Revenue", f"₹{total_rev/1e7:.2f}Cr", "Total Revenue", True),
        ("EBITDA", f"₹{total_ebitda/1e7:.2f}Cr", "Total EBITDA", total_ebitda > 0),
        ("GM%", f"{avg_gm:.1f}%", "Gross Margin", avg_gm > 0),
        ("EBITDA%", f"{avg_ebitda:.1f}%", "Avg EBITDA", avg_ebitda > 0),
        ("Stores", str(stores), "Active Stores", True),
    ]

    for col, card in zip(cols, cards):
        with col:
            st.markdown(metric_card(*card), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
# SNAPSHOT TABLE
# ═══════════════════════════════════════════════════════════════
def render_snapshot(df):

    st.markdown("<div class='section-header'>📋 Kitchen Snapshot</div>", unsafe_allow_html=True)

    table = df.groupby("STORE").agg({
        "NET REVENUE": "sum",
        "GM%": "mean",
        "EBITDA%": "mean",
        "KITCHEN EBITDA": "sum"
    }).reset_index()

    table.columns = [
        "Store",
        "Revenue",
        "GM%",
        "EBITDA%",
        "EBITDA"
    ]

    table["Revenue"] = (table["Revenue"] / 1e5).round(1)
    table["EBITDA"] = (table["EBITDA"] / 1e5).round(1)

    styled = safe_style_dataframe(table, cols=["EBITDA", "EBITDA%"])

    st.dataframe(styled, use_container_width=True, height=500)

# ═══════════════════════════════════════════════════════════════
# TREND CHARTS
# ═══════════════════════════════════════════════════════════════
def render_trends(df):

    st.markdown("<div class='section-header'>📈 Revenue Trends</div>", unsafe_allow_html=True)

    monthly = df.groupby("MONTH").agg({
        "NET REVENUE": "sum",
        "KITCHEN EBITDA": "sum",
        "GM%": "mean",
        "EBITDA%": "mean"
    }).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        fig = go.Figure()

        fig.add_trace(go.Bar(
            x=monthly["MONTH"],
            y=monthly["NET REVENUE"] / 1e5,
            marker_color=PURPLE,
            name="Revenue"
        ))

        fig.add_trace(go.Scatter(
            x=monthly["MONTH"],
            y=monthly["KITCHEN EBITDA"] / 1e5,
            mode="lines+markers",
            line=dict(color=EMERALD, width=4),
            name="EBITDA"
        ))

        fig.update_layout(
            title="Revenue vs EBITDA",
            **base_layout()
        )

        st.plotly_chart(fig, use_container_width=True)

    with col2:

        fig2 = go.Figure()

        fig2.add_trace(go.Scatter(
            x=monthly["MONTH"],
            y=monthly["GM%"],
            mode="lines+markers",
            line=dict(color=CYAN, width=4),
            name="GM%"
        ))

        fig2.add_trace(go.Scatter(
            x=monthly["MONTH"],
            y=monthly["EBITDA%"],
            mode="lines+markers",
            line=dict(color=AMBER, width=4),
            name="EBITDA%"
        ))

        fig2.update_layout(
            title="Margin Trends",
            **base_layout()
        )

        st.plotly_chart(fig2, use_container_width=True)

# ═══════════════════════════════════════════════════════════════
# CITY ANALYSIS
# ═══════════════════════════════════════════════════════════════
def render_city_analysis(df):

    st.markdown("<div class='section-header'>🏙️ City Performance</div>", unsafe_allow_html=True)

    city = df.groupby("CITY").agg({
        "NET REVENUE": "sum",
        "KITCHEN EBITDA": "sum"
    }).reset_index()

    col1, col2 = st.columns(2)

    with col1:

        fig = px.bar(
            city,
            x="CITY",
            y="NET REVENUE",
            color="CITY",
            color_discrete_sequence=COLORS
        )

        fig.update_layout(
            title="Revenue by City",
            **base_layout(),
     
