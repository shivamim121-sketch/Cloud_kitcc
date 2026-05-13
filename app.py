"""
╔══════════════════════════════════════════════════════════════╗
║         KITCHEN P&L INTELLIGENCE DASHBOARD                  ║
║         Cloud Kitchen Analytics Suite v3.1                  ║
╚══════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
import hashlib
from pathlib import Path

BASE_DIR  = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════════
#  WHITE THEME CSS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* White background */
.stApp { 
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 50%, #e2e8f0 100%); 
}

/* Sidebar - light */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
    border-right: 1px solid rgba(99,102,241,0.15);
}
[data-testid="stSidebar"] label { 
    color: #4f46e5 !important; 
    font-weight: 600; 
    font-size: 0.78rem; 
    letter-spacing: 0.05em; 
}
[data-testid="stSidebar"] .stMarkdown { color: #334155 !important; }

/* Metric Cards - light */
.metric-card {
    background: linear-gradient(145deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid rgba(99,102,241,0.18);
    border-radius: 20px;
    padding: 22px 18px;
    text-align: center;
    backdrop-filter: blur(16px);
    transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
    margin-bottom: 10px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 2px 12px rgba(99,102,241,0.06);
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #6366f1, #06b6d4, #10b981);
    opacity: 0.8;
}
.metric-card:hover { 
    transform: translateY(-4px); 
    box-shadow: 0 12px 40px rgba(99,102,241,0.12), 0 0 60px rgba(99,102,241,0.04);
    border-color: rgba(99,102,241,0.35);
}
.metric-title { 
    color: #64748b; 
    font-size: 0.68rem; 
    font-weight: 700; 
    letter-spacing: 0.1em; 
    text-transform: uppercase; 
    margin-bottom: 8px; 
}
.metric-value { 
    color: #0f172a; 
    font-size: 1.7rem; 
    font-weight: 900; 
}
.metric-delta { 
    font-size: 0.72rem; 
    font-weight: 600; 
    margin-top: 6px; 
    letter-spacing: 0.03em;
}
.delta-pos { color: #059669; }
.delta-neg { color: #dc2626; }

/* Section Headers - light */
.section-header {
    background: linear-gradient(90deg, rgba(99,102,241,0.1), rgba(6,182,212,0.03), transparent);
    border-left: 4px solid #6366f1;
    padding: 12px 18px;
    border-radius: 0 12px 12px 0;
    margin: 24px 0 14px 0;
    color: #1e293b;
    font-weight: 800;
    font-size: 1.05rem;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    position: relative;
}
.section-header::after {
    content: '';
    position: absolute;
    bottom: -2px; left: 0;
    width: 60px; height: 2px;
    background: linear-gradient(90deg, #6366f1, transparent);
}

/* Tabs - light */
[data-testid="stTab"] { 
    color: #94a3b8; 
    font-weight: 600; 
    font-size: 0.9rem;
    padding: 10px 16px !important;
    transition: all 0.2s;
}
[data-testid="stTab"]:hover { color: #6366f1; }
[data-testid="stTab"][aria-selected="true"] { 
    color: #4f46e5; 
    border-bottom: 3px solid #6366f1;
    background: linear-gradient(180deg, rgba(99,102,241,0.06), transparent);
}

/* Dataframe */
.stDataFrame { border-radius: 14px; overflow: hidden; }

/* Expander */
[data-testid="stExpander"] { 
    border: 1px solid rgba(99,102,241,0.15); 
    border-radius: 12px; 
    background: rgba(255,255,255,0.7);
}

/* Filter pills */
.filter-pill {
    display: inline-block;
    background: linear-gradient(135deg, rgba(99,102,241,0.12), rgba(6,182,212,0.08));
    border: 1px solid rgba(99,102,241,0.25);
    border-radius: 24px;
    padding: 4px 14px;
    font-size: 0.72rem;
    color: #4f46e5;
    margin: 3px;
    font-weight: 500;
}

/* Brand */
.brand {
    text-align: center; padding: 20px 0 28px 0; position: relative;
}
.brand::after {
    content: '';
    position: absolute; bottom: 10px; left: 25%; right: 25%;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(99,102,241,0.3), transparent);
}
.brand-title { 
    color: #0f172a; 
    font-size: 1.15rem; 
    font-weight: 900; 
    letter-spacing: 0.04em;
}
.brand-sub { 
    color: #64748b; 
    font-size: 0.7rem; 
    letter-spacing: 0.1em; 
    margin-top: 4px;
}

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: rgba(99,102,241,0.05); }
::-webkit-scrollbar-thumb { background: rgba(99,102,241,0.25); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: rgba(99,102,241,0.4); }

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, rgba(99,102,241,0.15), rgba(6,182,212,0.1)) !important;
    border: 1px solid rgba(99,102,241,0.3) !important;
    border-radius: 10px !important;
    color: #4f46e5 !important;
    font-weight: 600 !important;
    transition: all 0.3s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, rgba(99,102,241,0.25), rgba(6,182,212,0.18)) !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.15) !important;
    transform: translateY(-1px);
}

/* Alerts */
[data-testid="stAlert"] {
    border-radius: 12px !important;
    border: 1px solid rgba(99,102,241,0.15) !important;
}

/* Header glow - dark text on light */
.header-glow {
    background: linear-gradient(90deg, #4f46e5, #0891b2, #059669, #4f46e5);
    background-size: 300% 300%;
    animation: gradient-shift 8s ease infinite;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

@keyframes gradient-shift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

/* Radio buttons */
[data-testid="stRadio"] label {
    color: #334155 !important;
}

/* Multiselect tags */
[data-testid="stMultiSelect"] span[data-baseweb="tag"] {
    background: linear-gradient(135deg, rgba(99,102,241,0.15), rgba(6,182,212,0.1)) !important;
    border: 1px solid rgba(99,102,241,0.25) !important;
}

/* Slider */
[data-testid="stSlider"] > div > div > div {
    background: linear-gradient(90deg, #6366f1, #06b6d4) !important;
}

/* Dark text for main content */
.stMarkdown p, .stMarkdown li, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
    color: #1e293b;
}
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING & CACHING ─────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data(filepath) -> pd.DataFrame:
    """Load and preprocess the kitchen PNL data."""
    df = pd.read_excel(filepath, header=1)

    df["GM%"]      = (df["GROSS MARGIN"]   / df["NET REVENUE"] * 100).round(2)
    df["EBITDA%"]  = (df["KITCHEN EBITDA"] / df["NET REVENUE"] * 100).round(2)
    df["VARIANCE%"]= (df["VARIANCE"]       / df["NET REVENUE"] * 100).round(4)

    df["CM"]  = df["GROSS MARGIN"]
    df["CM%"] = df["GM%"]

    month_order = ["Oct-2023","Nov-2023","Dec-2023","Jan-2024","Feb-2024","Mar-2024"]
    df["MONTH"] = pd.Categorical(df["MONTH"], categories=month_order, ordered=True)
    df = df.sort_values("MONTH")
    df["MONTH_STR"] = df["MONTH"].astype(str)

    def var_bucket(v):
        if v < 15000:   return "(a) Var < ₹15K"
        elif v < 20000: return "(b) Var ₹15K-20K"
        elif v < 25000: return "(c) Var ₹20K-25K"
        else:           return "(d) Var > ₹25K"
    df["VAR_BUCKET"] = df["VARIANCE"].apply(var_bucket)

    def var_pct_bucket(v):
        if v < 0.4:    return "(a) Var < 0.4%"
        elif v < 0.6:  return "(b) Var 0.4-0.6%"
        elif v < 0.8:  return "(c) Var 0.6-0.8%"
        else:          return "(d) Var > 0.8%"
    df["VAR_PCT_BUCKET"] = df["VARIANCE%"].apply(var_pct_bucket)

    def rev_bucket(r):
        lacs = r / 1e5
        if lacs < 15:   return "(a) Below INR 15 lacs"
        elif lacs < 25: return "(b) INR 15 to 25 lacs"
        elif lacs < 35: return "(c) INR 25 to 35 lacs"
        elif lacs < 45: return "(d) INR 35 to 45 lacs"
        else:           return "(e) Above INR 45 lacs"
    df["REV_BUCKET"] = df["NET REVENUE"].apply(rev_bucket)

    df["EBITDA_SIGN"] = df["KITCHEN EBITDA"].apply(lambda x: "▲ Positive" if x >= 0 else "▼ Negative")

    return df


# ─── CHART DEFAULTS (WHITE THEME) ───────────────────────────────────────────
CHART_BG   = "rgba(0,0,0,0)"
PAPER_BG   = "rgba(0,0,0,0)"
FONT_COLOR = "#1e293b"
GRID_COLOR = "rgba(99,102,241,0.1)"
PURPLE     = "#6366f1"
CYAN       = "#06b6d4"
EMERALD    = "#10b981"
AMBER      = "#f59e0b"
ROSE       = "#ef4444"
COLORS     = [PURPLE, CYAN, EMERALD, AMBER, ROSE, "#8b5cf6", "#22d3ee", "#34d399", "#f472b6", "#84cc16"]

def base_layout(**kwargs):
    return dict(
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=CHART_BG,
        font=dict(family="Inter", color=FONT_COLOR, size=12),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(
            bgcolor="rgba(255,255,255,0.8)", 
            bordercolor="rgba(99,102,241,0.2)", 
            borderwidth=1,
            font=dict(size=11, color="#1e293b")
        ),
        **kwargs
    )

def styled_axis(title="", fmt="", showgrid=True):
    return dict(
        title=title,
        gridcolor=GRID_COLOR if showgrid else "rgba(0,0,0,0)",
        gridwidth=1,
        linecolor="rgba(99,102,241,0.2)",
        tickfont=dict(color="#64748b", size=11),
        title_font=dict(color="#475569", size=12),
        tickformat=fmt,
        zeroline=False,
    )


# ─── SAFE STYLING HELPERS (NO MATPLOTLIB) ───────────────────────────────────
def safe_style_apply(styler, func, subset=None):
    """Safely apply styling function, handling pandas version differences."""
    try:
        return styler.map(func, subset=subset)
    except (AttributeError, TypeError):
        try:
            return styler.applymap(func, subset=subset)
        except Exception:
            return styler


def hex_to_rgba(hex_color, alpha=0.15):
    """Convert hex color to rgba string for Plotly fillcolor."""
    if not isinstance(hex_color, str):
        return f"rgba(99,102,241,{alpha})"
    hex_color = hex_color.lstrip('#')
    if len(hex_color) < 6:
        # Try to expand short hex
        if len(hex_color) == 3:
            hex_color = ''.join([c*2 for c in hex_color])
        else:
            return f"rgba(99,102,241,{alpha})"
    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return f"rgba({r},{g},{b},{alpha})"
    except (ValueError, IndexError):
        return f"rgba(99,102,241,{alpha})"


# ─── SIDEBAR ────────────────────────────────────────────────────────────────
def render_sidebar(df: pd.DataFrame):
    st.sidebar.markdown("""
    <div class='brand'>
        <div class='brand-title'>🍳 Kitchen Intelligence</div>
        <div class='brand-sub'>P&L Analytics Suite · v3.1</div>
    </div>
    """, unsafe_allow_html=True)
    st.sidebar.markdown("---")

    dashboard = st.sidebar.radio(
        "📊 Select Dashboard",
        ["🏪 Kitchen Level PNL", "📉 Variance Level PNL"],
        label_visibility="visible"
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
        st.sidebar.markdown("**📦 Cohort Filters**")

        filters["rev_cohort"] = st.sidebar.multiselect(
            "💰 Revenue Cohort",
            options=sorted(df["REVENUE COHORT"].unique().tolist()),
            default=sorted(df["REVENUE COHORT"].unique().tolist())
        )
        filters["cm_cohort"] = st.sidebar.multiselect(
            "📊 CM Cohort",
            options=sorted(df["CM COHORT"].unique().tolist()),
            default=sorted(df["CM COHORT"].unique().tolist())
        )
        filters["ebitda_cat"] = st.sidebar.multiselect(
            "📈 EBITDA Category",
            options=df["EBITDA CATEGORY"].unique().tolist(),
            default=df["EBITDA CATEGORY"].unique().tolist()
        )
        filters["ebitda_cohort"] = st.sidebar.multiselect(
            "🎯 EBITDA Cohort",
            options=sorted(df["EBITDA COHORT"].unique().tolist()),
            default=sorted(df["EBITDA COHORT"].unique().tolist())
        )

        st.sidebar.markdown("---")
        st.sidebar.markdown("**📐 Range Filters**")

        rev_min, rev_max = float(df["NET REVENUE"].min()), float(df["NET REVENUE"].max())
        filters["rev_range"] = st.sidebar.slider(
            "💵 Net Revenue (₹)",
            min_value=rev_min, max_value=rev_max,
            value=(rev_min, rev_max), format="₹%,.0f"
        )

        cm_min, cm_max = float(df["CM"].min()), float(df["CM"].max())
        filters["cm_range"] = st.sidebar.slider(
            "📊 Contribution Margin (₹)",
            min_value=cm_min, max_value=cm_max,
            value=(cm_min, cm_max), format="₹%,.0f"
        )

        ebitda_min, ebitda_max = float(df["KITCHEN EBITDA"].min()), float(df["KITCHEN EBITDA"].max())
        filters["ebitda_range"] = st.sidebar.slider(
            "📈 EBITDA (₹)",
            min_value=ebitda_min, max_value=ebitda_max,
            value=(ebitda_min, ebitda_max), format="₹%,.0f"
        )

    else:
        st.sidebar.markdown("### 🎛️ Variance Filters")

        filters["var_bucket"] = st.sidebar.multiselect(
            "🔍 Variance Category (₹ Amount)",
            options=["(a) Var < ₹15K", "(b) Var ₹15K-20K", "(c) Var ₹20K-25K", "(d) Var > ₹25K"],
            default=["(a) Var < ₹15K", "(b) Var ₹15K-20K", "(c) Var ₹20K-25K", "(d) Var > ₹25K"]
        )
        filters["city"] = st.sidebar.multiselect(
            "🏙️ City",
            options=sorted(df["CITY"].unique().tolist()),
            default=sorted(df["CITY"].unique().tolist())
        )
        filters["month"] = st.sidebar.multiselect(
            "📅 Month",
            options=df["MONTH_STR"].unique().tolist(),
            default=df["MONTH_STR"].unique().tolist()
        )

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        "<div style='text-align:center; color:#94a3b8; font-size:0.68rem;'>Data refreshes every 5 min<br>📊 Kitchen P&L Suite v3.1</div>",
        unsafe_allow_html=True
    )

    return dashboard, filters


# ─── FILTER APPLICATION ──────────────────────────────────────────────────────
def apply_pnl_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    mask = pd.Series([True] * len(df), index=df.index)
    if f.get("month"): mask &= df["MONTH_STR"].isin(f["month"])
    if f.get("store"): mask &= df["STORE"].isin(f["store"])
    if f.get("city"): mask &= df["CITY"].isin(f["city"])
    if f.get("zone"): mask &= df["ZONE MAPPING"].isin(f["zone"])
    if f.get("status"): mask &= df["STATUS"].isin(f["status"])
    if f.get("rev_cohort"): mask &= df["REVENUE COHORT"].isin(f["rev_cohort"])
    if f.get("cm_cohort"): mask &= df["CM COHORT"].isin(f["cm_cohort"])
    if f.get("ebitda_cat"): mask &= df["EBITDA CATEGORY"].isin(f["ebitda_cat"])
    if f.get("ebitda_cohort"): mask &= df["EBITDA COHORT"].isin(f["ebitda_cohort"])
    if f.get("rev_range"): mask &= df["NET REVENUE"].between(*f["rev_range"])
    if f.get("cm_range"): mask &= df["CM"].between(*f["cm_range"])
    if f.get("ebitda_range"): mask &= df["KITCHEN EBITDA"].between(*f["ebitda_range"])
    return df[mask]


def apply_var_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    mask = pd.Series([True] * len(df), index=df.index)
    if f.get("var_bucket"): mask &= df["VAR_BUCKET"].isin(f["var_bucket"])
    if f.get("city"): mask &= df["CITY"].isin(f["city"])
    if f.get("month"): mask &= df["MONTH_STR"].isin(f["month"])
    return df[mask]


# ═══════════════════════════════════════════════════════════════════════════════
#  KITCHEN LEVEL PNL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def render_kitchen_pnl(df: pd.DataFrame, filters: dict):
    """Render the Kitchen Level P&L dashboard."""
    fdf = apply_pnl_filters(df, filters)

    if fdf.empty:
        st.warning("No data matches the selected filters.")
        return

    # ── Header ──
    st.markdown("""
    <div style="text-align:center; margin-bottom: 24px;">
        <h1 class="header-glow" style="font-size: 2.2rem; font-weight: 900; margin-bottom: 4px;">
            🏪 Kitchen Level P&L
        </h1>
        <p style="color: #64748b; font-size: 0.9rem; margin-top: 0;">
            Profit & Loss analysis across all kitchen locations
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── KPI Cards ──
    total_rev = fdf["NET REVENUE"].sum()
    total_gm = fdf["GROSS MARGIN"].sum()
    total_ebitda = fdf["KITCHEN EBITDA"].sum()
    avg_gm_pct = (total_gm / total_rev * 100) if total_rev > 0 else 0
    avg_ebitda_pct = (total_ebitda / total_rev * 100) if total_rev > 0 else 0
    store_count = fdf["STORE"].nunique()
    active_count = fdf[fdf["STATUS"] == "Active"]["STORE"].nunique() if "Active" in fdf["STATUS"].values else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Total Revenue</div>
            <div class="metric-value">₹{total_rev/1e7:.2f}Cr</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Gross Margin</div>
            <div class="metric-value">₹{total_gm/1e7:.2f}Cr</div>
            <div class="metric-delta {'delta-pos' if avg_gm_pct >= 0 else 'delta-neg'}">{avg_gm_pct:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Kitchen EBITDA</div>
            <div class="metric-value">₹{total_ebitda/1e7:.2f}Cr</div>
            <div class="metric-delta {'delta-pos' if avg_ebitda_pct >= 0 else 'delta-neg'}">{avg_ebitda_pct:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Stores</div>
            <div class="metric-value">{store_count}</div>
            <div class="metric-delta delta-pos">{active_count} Active</div>
        </div>
        """, unsafe_allow_html=True)
    with c5:
        prof_pct = (fdf["KITCHEN EBITDA"] >= 0).mean() * 100
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Profitable</div>
            <div class="metric-value">{prof_pct:.0f}%</div>
            <div class="metric-delta">{(fdf["KITCHEN EBITDA"] >= 0).sum()}/{len(fdf)} stores</div>
        </div>
        """, unsafe_allow_html=True)

    # ── Tabs ──
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Overview", "📈 Trends", "🏪 Store Analysis", "📋 Data"])

    with tab1:
        render_overview(fdf)
    with tab2:
        render_trends(fdf)
    with tab3:
        render_store_analysis(fdf)
    with tab4:
        render_data_table(fdf)


def render_overview(fdf: pd.DataFrame):
    """Render overview charts."""
    st.markdown('<div class="section-header">Revenue & EBITDA by Month</div>', unsafe_allow_html=True)

    monthly = fdf.groupby("MONTH_STR").agg({
        "NET REVENUE": "sum",
        "GROSS MARGIN": "sum",
        "KITCHEN EBITDA": "sum"
    }).reset_index()

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Bar(
        x=monthly["MONTH_STR"], y=monthly["NET REVENUE"]/1e5,
        name="Net Revenue", marker_color=PURPLE, opacity=0.8,
        hovertemplate="Month: %{x}<br>Revenue: ₹%{y:.1f}L<extra></extra>"
    ), secondary_y=False)

    fig.add_trace(go.Bar(
        x=monthly["MONTH_STR"], y=monthly["GROSS MARGIN"]/1e5,
        name="Gross Margin", marker_color=CYAN, opacity=0.8,
        hovertemplate="Month: %{x}<br>GM: ₹%{y:.1f}L<extra></extra>"
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        x=monthly["MONTH_STR"], y=monthly["KITCHEN EBITDA"]/1e5,
        name="EBITDA", mode="lines+markers", line=dict(color=EMERALD, width=3),
        marker=dict(size=8), fill="tozeroy",
        fillcolor=hex_to_rgba(EMERALD, 0.15),
        hovertemplate="Month: %{x}<br>EBITDA: ₹%{y:.1f}L<extra></extra>"
    ), secondary_y=True)

    fig.update_layout(
        **base_layout(title="Monthly Revenue, GM & EBITDA", barmode="group"),
        height=450
    )
    fig.update_yaxes(title_text="Revenue / GM (₹ Lakhs)", secondary_y=False, **styled_axis())
    fig.update_yaxes(title_text="EBITDA (₹ Lakhs)", secondary_y=True, **styled_axis())
    fig.update_xaxes(**styled_axis(showgrid=False))
    st.plotly_chart(fig, use_container_width=True)

    # EBITDA Distribution
    st.markdown('<div class="section-header">EBITDA Distribution</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        ebitda_by_sign = fdf.groupby("EBITDA_SIGN").size().reset_index(name="Count")
        colors_pie = [EMERALD, ROSE]
        fig_pie = go.Figure(go.Pie(
            labels=ebitda_by_sign["EBITDA_SIGN"],
            values=ebitda_by_sign["Count"],
            hole=0.55,
            marker_colors=colors_pie,
            textinfo="label+percent",
            textfont=dict(size=12, color="#1e293b"),
            hovertemplate="%{label}<br>Stores: %{value}<br>Share: %{percent}<extra></extra>"
        ))
        fig_pie.update_layout(
            **base_layout(title="Positive vs Negative EBITDA"),
            height=350,
            annotations=[dict(text="EBITDA", x=0.5, y=0.5, font_size=16, showarrow=False, font_color="#1e293b")]
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col2:
        ebitda_hist = fdf["KITCHEN EBITDA"] / 1e5
        fig_hist = go.Figure(go.Histogram(
            x=ebitda_hist, nbinsx=30,
            marker_color=PURPLE, opacity=0.7,
            hovertemplate="EBITDA: ₹%{x:.1f}L<br>Stores: %{y}<extra></extra>"
        ))
        fig_hist.add_vline(x=0, line_dash="dash", line_color=ROSE, line_width=2)
        fig_hist.update_layout(
            **base_layout(title="EBITDA Distribution (₹ Lakhs)"),
            height=350,
            xaxis=styled_axis("EBITDA (₹ Lakhs)"),
            yaxis=styled_axis("Store Count")
        )
        st.plotly_chart(fig_hist, use_container_width=True)


def render_trends(fdf: pd.DataFrame):
    """Render trend analysis."""
    st.markdown('<div class="section-header">Monthly Trends</div>', unsafe_allow_html=True)

    monthly = fdf.groupby("MONTH_STR").agg({
        "NET REVENUE": "sum",
        "GROSS MARGIN": "sum",
        "KITCHEN EBITDA": "sum",
        "STORE": "nunique"
    }).reset_index()
    monthly["GM%"] = (monthly["GROSS MARGIN"] / monthly["NET REVENUE"] * 100).round(2)
    monthly["EBITDA%"] = (monthly["KITCHEN EBITDA"] / monthly["NET REVENUE"] * 100).round(2)

    fig = make_subplots(rows=2, cols=1, subplot_titles=("Revenue Trend", "Margin % Trend"),
                        vertical_spacing=0.15)

    fig.add_trace(go.Scatter(
        x=monthly["MONTH_STR"], y=monthly["NET REVENUE"]/1e5,
        mode="lines+markers", name="Revenue", line=dict(color=PURPLE, width=3),
        marker=dict(size=10), fill="tozeroy", fillcolor=hex_to_rgba(PURPLE, 0.1),
        hovertemplate="Month: %{x}<br>Revenue: ₹%{y:.1f}L<extra></extra>"
    ), row=1, col=1)

    fig.add_trace(go.Scatter(
        x=monthly["MONTH_STR"], y=monthly["GM%"],
        mode="lines+markers", name="GM%", line=dict(color=CYAN, width=3),
        marker=dict(size=8), hovertemplate="Month: %{x}<br>GM%: %{y:.1f}%<extra></extra>"
    ), row=2, col=1)

    fig.add_trace(go.Scatter(
        x=monthly["MONTH_STR"], y=monthly["EBITDA%"],
        mode="lines+markers", name="EBITDA%", line=dict(color=EMERALD, width=3),
        marker=dict(size=8), hovertemplate="Month: %{x}<br>EBITDA%: %{y:.1f}%<extra></extra>"
    ), row=2, col=1)

    fig.update_layout(**base_layout(height=650, showlegend=True))
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(**styled_axis())
    st.plotly_chart(fig, use_container_width=True)

    # City-wise comparison
    st.markdown('<div class="section-header">City-wise Performance</div>', unsafe_allow_html=True)
    city_perf = fdf.groupby("CITY").agg({
        "NET REVENUE": "sum",
        "KITCHEN EBITDA": "sum",
        "STORE": "nunique"
    }).reset_index()
    city_perf["EBITDA%"] = (city_perf["KITCHEN EBITDA"] / city_perf["NET REVENUE"] * 100).round(2)
    city_perf = city_perf.sort_values("NET REVENUE", ascending=True)

    fig_city = go.Figure()
    fig_city.add_trace(go.Bar(
        y=city_perf["CITY"], x=city_perf["NET REVENUE"]/1e5,
        orientation="h", name="Revenue", marker_color=PURPLE, opacity=0.8,
        hovertemplate="%{y}<br>Revenue: ₹%{x:.1f}L<extra></extra>"
    ))
    fig_city.add_trace(go.Scatter(
        y=city_perf["CITY"], x=city_perf["EBITDA%"],
        mode="markers", name="EBITDA%", marker=dict(color=EMERALD, size=12, symbol="diamond"),
        xaxis="x2", hovertemplate="%{y}<br>EBITDA%: %{x:.1f}%<extra></extra>"
    ))
    fig_city.update_layout(
        **base_layout(title="Revenue & EBITDA% by City", height=500),
        xaxis=styled_axis("Revenue (₹ Lakhs)"),
        xaxis2=dict(overlaying="x", side="top", title="EBITDA %", tickfont=dict(color="#64748b")),
        yaxis=dict(tickfont=dict(size=11, color="#1e293b"))
    )
    st.plotly_chart(fig_city, use_container_width=True)


def render_store_analysis(fdf: pd.DataFrame):
    """Render store-level analysis."""
    st.markdown('<div class="section-header">Store Performance Matrix</div>', unsafe_allow_html=True)

    store_perf = fdf.groupby("STORE").agg({
        "NET REVENUE": "mean",
        "KITCHEN EBITDA": "mean",
        "GM%": "mean",
        "CITY": "first",
        "STATUS": "first"
    }).reset_index()
    store_perf["EBITDA%"] = (store_perf["KITCHEN EBITDA"] / store_perf["NET REVENUE"] * 100).round(2)

    # Create safe size column for marker sizing
    store_perf_plot = store_perf.copy()
    store_perf_plot["GM_SIZE"] = store_perf_plot["GM%"].abs().fillna(0).replace([np.inf, -np.inf], 0)
    store_perf_plot["GM_SIZE"] = store_perf_plot["GM_SIZE"].clip(lower=1)

    fig = px.scatter(
        store_perf_plot, x="NET REVENUE", y="EBITDA%",
        color="CITY", size="GM_SIZE", hover_data=["STORE", "STATUS", "GM%"],
        color_discrete_sequence=COLORS,
        title="Revenue vs EBITDA% (Bubble = |GM%|)"
    )
    fig.update_layout(**base_layout(height=500))
    fig.update_xaxes(title="Avg Net Revenue (₹)", **styled_axis(fmt="₹,.0f"))
    fig.update_yaxes(title="EBITDA %", **styled_axis())
    fig.update_traces(marker=dict(opacity=0.7, line=dict(width=1, color="white")))
    st.plotly_chart(fig, use_container_width=True)

    # Top/Bottom performers
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### 🏆 Top 10 EBITDA Stores")
        top = store_perf.nlargest(10, "KITCHEN EBITDA")[["STORE", "CITY", "NET REVENUE", "EBITDA%"]]
        top["NET REVENUE"] = top["NET REVENUE"].apply(lambda x: f"₹{x/1e5:.1f}L")
        top["EBITDA%"] = top["EBITDA%"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(top, use_container_width=True, hide_index=True)

    with col2:
        st.markdown("##### ⚠️ Bottom 10 EBITDA Stores")
        bottom = store_perf.nsmallest(10, "KITCHEN EBITDA")[["STORE", "CITY", "NET REVENUE", "EBITDA%"]]
        bottom["NET REVENUE"] = bottom["NET REVENUE"].apply(lambda x: f"₹{x/1e5:.1f}L")
        bottom["EBITDA%"] = bottom["EBITDA%"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(bottom, use_container_width=True, hide_index=True)


def render_data_table(fdf: pd.DataFrame):
    """Render raw data table."""
    st.markdown('<div class="section-header">Raw Data</div>', unsafe_allow_html=True)

    display_cols = ["MONTH_STR", "STORE", "CITY", "ZONE MAPPING", "STATUS",
                    "NET REVENUE", "GROSS MARGIN", "GM%", "KITCHEN EBITDA", "EBITDA%", "VARIANCE"]
    display_df = fdf[display_cols].copy()
    display_df["NET REVENUE"] = display_df["NET REVENUE"].apply(lambda x: f"₹{x:,.0f}")
    display_df["GROSS MARGIN"] = display_df["GROSS MARGIN"].apply(lambda x: f"₹{x:,.0f}")
    display_df["KITCHEN EBITDA"] = display_df["KITCHEN EBITDA"].apply(lambda x: f"₹{x:,.0f}")
    display_df["VARIANCE"] = display_df["VARIANCE"].apply(lambda x: f"₹{x:,.0f}")
    display_df["GM%"] = display_df["GM%"].apply(lambda x: f"{x:.1f}%")
    display_df["EBITDA%"] = display_df["EBITDA%"].apply(lambda x: f"{x:.1f}%")

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = fdf.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Download CSV", csv, "kitchen_pnl_data.csv", "text/csv")


# ═══════════════════════════════════════════════════════════════════════════════
#  VARIANCE LEVEL PNL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def render_variance_pnl(df: pd.DataFrame, filters: dict):
    """Render the Variance Level P&L dashboard."""
    fdf = apply_var_filters(df, filters)

    if fdf.empty:
        st.warning("No data matches the selected filters.")
        return

    # ── Header ──
    st.markdown("""
    <div style="text-align:center; margin-bottom: 24px;">
        <h1 class="header-glow" style="font-size: 2.2rem; font-weight: 900; margin-bottom: 4px;">
            📉 Variance Level P&L
        </h1>
        <p style="color: #64748b; font-size: 0.9rem; margin-top: 0;">
            Variance analysis across kitchen locations
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Insights Summary (WHITE THEME) ──
    with st.expander("📊 View Full Insights Summary", expanded=False):
        total_stores = fdf["STORE"].nunique()
        active_stores = fdf[fdf["STATUS"] == "Active"]["STORE"].nunique() if "Active" in fdf["STATUS"].values else 0
        inactive_stores = total_stores - active_stores
        active_rate = (active_stores / total_stores * 100) if total_stores > 0 else 0
        avg_rev = fdf["NET REVENUE"].mean()

        monthly_ebitda = fdf.groupby("MONTH_STR")["KITCHEN EBITDA"].sum()
        best_month = monthly_ebitda.idxmax() if not monthly_ebitda.empty else "N/A"
        worst_month = monthly_ebitda.idxmin() if not monthly_ebitda.empty else "N/A"
        best_ebitda = monthly_ebitda.max() if not monthly_ebitda.empty else 0
        worst_ebitda = monthly_ebitda.min() if not monthly_ebitda.empty else 0

        rev_min = fdf["NET REVENUE"].min()
        rev_max = fdf["NET REVENUE"].max()
        gm_min = fdf["GM%"].min()
        gm_max = fdf["GM%"].max()
        ebitda_min = fdf["EBITDA%"].min()
        ebitda_max = fdf["EBITDA%"].max()

        prof_count = (fdf["KITCHEN EBITDA"] >= 0).sum()
        total_count = len(fdf)

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
            <div style="background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.15); border-radius: 12px; padding: 16px;">
                <h4 style="color: #4f46e5; margin: 0 0 8px 0; font-size: 0.95rem;">📅 Monthly Performance</h4>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Best Month</b>: <code style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px; color: #0f172a;">{best_month}</code> — ₹{best_ebitda/1e7:.2f} Cr total EBITDA</p>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Worst Month</b>: <code style="background: #f1f5f9; padding: 2px 6px; border-radius: 4px; color: #0f172a;">{worst_month}</code> — ₹{worst_ebitda/1e7:.2f} Cr total EBITDA</p>
            </div>
            <div style="background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.15); border-radius: 12px; padding: 16px;">
                <h4 style="color: #4f46e5; margin: 0 0 8px 0; font-size: 0.95rem;">🏪 Store Metrics</h4>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Active Rate</b>: {active_rate:.0f}% ({active_stores} active / {inactive_stores} inactive)</p>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Avg Revenue/Store</b>: ₹{avg_rev/1e5:.1f}L per month</p>
            </div>
            <div style="background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.15); border-radius: 12px; padding: 16px;">
                <h4 style="color: #4f46e5; margin: 0 0 8px 0; font-size: 0.95rem;">💰 Financial Range</h4>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Revenue Range</b>: ₹{rev_min/1e5:.1f}L – ₹{rev_max/1e5:.1f}L per store-month</p>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>GM% Range</b>: {gm_min:.1f}% – {gm_max:.1f}%</p>
            </div>
            <div style="background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.15); border-radius: 12px; padding: 16px;">
                <h4 style="color: #4f46e5; margin: 0 0 8px 0; font-size: 0.95rem;">📈 Profitability</h4>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>EBITDA% Range</b>: {ebitda_min:.1f}% – {ebitda_max:.1f}%</p>
                <p style="color: #334155; margin: 4px 0; font-size: 0.85rem;"><b>Profitable Stores</b>: {prof_count} / {total_count} store-months</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ── KPI Cards ──
    total_var = fdf["VARIANCE"].sum()
    avg_var = fdf["VARIANCE"].mean()
    var_buckets = fdf["VAR_BUCKET"].value_counts()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Total Variance</div>
            <div class="metric-value">₹{total_var/1e5:.1f}L</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Avg Variance</div>
            <div class="metric-value">₹{avg_var:,.0f}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        high_var = var_buckets.get("(d) Var > ₹25K", 0)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">High Variance</div>
            <div class="metric-value">{high_var}</div>
            <div class="metric-delta delta-neg">stores > ₹25K</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        low_var = var_buckets.get("(a) Var < ₹15K", 0)
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">Low Variance</div>
            <div class="metric-value">{low_var}</div>
            <div class="metric-delta delta-pos">stores < ₹15K</div>
        </div>
        """, unsafe_allow_html=True)

    # ── Tabs ──
    tab1, tab2, tab3 = st.tabs(["📉 Variance Distribution", "🔍 Variance Analysis", "📋 Data"])

    with tab1:
        render_variance_distribution(fdf)
    with tab2:
        render_variance_analysis(fdf)
    with tab3:
        render_variance_data(fdf)


def render_variance_distribution(fdf: pd.DataFrame):
    """Render variance distribution charts - FIXED fillcolor issue."""
    st.markdown('<div class="section-header">📉 Variance Distribution & Trends</div>', unsafe_allow_html=True)

    # Variance bucket distribution
    var_counts = fdf["VAR_BUCKET"].value_counts().sort_index()
    bucket_colors = {"(a) Var < ₹15K": EMERALD, "(b) Var ₹15K-20K": CYAN, 
                     "(c) Var ₹20K-25K": AMBER, "(d) Var > ₹25K": ROSE}

    col1, col2 = st.columns(2)

    with col1:
        fig1 = go.Figure(go.Bar(
            x=var_counts.index, y=var_counts.values,
            marker_color=[bucket_colors.get(b, PURPLE) for b in var_counts.index],
            opacity=0.85,
            hovertemplate="%{x}<br>Stores: %{y}<extra></extra>"
        ))
        fig1.update_layout(
            **base_layout(title="Variance Bucket Distribution"),
            height=380,
            xaxis=styled_axis("Variance Bucket", showgrid=False),
            yaxis=styled_axis("Store Count")
        )
        st.plotly_chart(fig1, use_container_width=True)

    with col2:
        # Monthly variance trend - FIXED: ensure valid fillcolor
        var_month = fdf.groupby(["MONTH_STR", "VAR_BUCKET"]).size().unstack(fill_value=0)
        months = var_month.index.tolist()

        fig2 = go.Figure()
        for i, bucket in enumerate(var_month.columns):
            color = bucket_colors.get(bucket, COLORS[i % len(COLORS)])
            # Ensure fillcolor is always a valid rgba string
            fill_color = hex_to_rgba(color, 0.15)

            fig2.add_trace(go.Scatter(
                x=months, y=var_month[bucket].values,
                mode="lines", name=bucket, line=dict(color=color, width=2.5),
                stackgroup="one", fillcolor=fill_color,
                hovertemplate=f"<b>{bucket}</b><br>Month: %{{x}}<br>Stores: %{{y}}<extra></extra>"
            ))

        fig2.update_layout(
            **base_layout(title="Monthly Variance Trend by Bucket"),
            height=380,
            xaxis=styled_axis(showgrid=False),
            yaxis=styled_axis("Store Count")
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Revenue bucket distribution
    st.markdown('<div class="section-header">🔍 Revenue Bucket Distribution (All Months)</div>', unsafe_allow_html=True)

    rev_counts = fdf["REV_BUCKET"].value_counts().sort_index()
    rev_colors = {"(a) Below INR 15 lacs": ROSE, "(b) INR 15 to 25 lacs": AMBER,
                  "(c) INR 25 to 35 lacs": CYAN, "(d) INR 35 to 45 lacs": PURPLE,
                  "(e) Above INR 45 lacs": EMERALD}

    fig3 = go.Figure(go.Bar(
        x=rev_counts.index, y=rev_counts.values,
        marker_color=[rev_colors.get(b, PURPLE) for b in rev_counts.index],
        opacity=0.85,
        hovertemplate="%{x}<br>Stores: %{y}<extra></extra>"
    ))
    fig3.update_layout(
        **base_layout(title="Revenue Bucket Distribution"),
        height=400,
        xaxis=styled_axis("Revenue Bucket", showgrid=False),
        yaxis=styled_axis("Store Count")
    )
    st.plotly_chart(fig3, use_container_width=True)

    # Variance % bucket heatmap
    with st.expander("🧊 View as Heatmap", expanded=False):
        heatmap_data = fdf.groupby(["CITY", "VAR_PCT_BUCKET"]).size().unstack(fill_value=0)
        if not heatmap_data.empty:
            fig_heat = px.imshow(
                heatmap_data.values,
                x=heatmap_data.columns.tolist(),
                y=heatmap_data.index.tolist(),
                color_continuous_scale="Purples",
                aspect="auto",
                title="Variance % by City"
            )
            fig_heat.update_layout(
                **base_layout(height=500),
                xaxis=styled_axis("Variance % Bucket", showgrid=False),
                yaxis=styled_axis("City", showgrid=False)
            )
            st.plotly_chart(fig_heat, use_container_width=True)
        else:
            st.info("No data available for heatmap.")


def render_variance_analysis(fdf: pd.DataFrame):
    """Render detailed variance analysis."""
    st.markdown('<div class="section-header">🔍 Variance vs Revenue Analysis</div>', unsafe_allow_html=True)

    # Create safe size column: EBITDA% must be non-negative for marker sizing
    plot_df = fdf.copy()
    plot_df["EBITDA_SIZE"] = plot_df["EBITDA%"].abs().fillna(0).replace([np.inf, -np.inf], 0)
    # Ensure minimum size so all points are visible
    plot_df["EBITDA_SIZE"] = plot_df["EBITDA_SIZE"].clip(lower=1)

    fig = px.scatter(
        plot_df, x="NET REVENUE", y="VARIANCE",
        color="VAR_BUCKET", size="EBITDA_SIZE", 
        hover_data=["STORE", "CITY", "MONTH_STR", "EBITDA%"],
        color_discrete_map={"(a) Var < ₹15K": EMERALD, "(b) Var ₹15K-20K": CYAN,
                           "(c) Var ₹20K-25K": AMBER, "(d) Var > ₹25K": ROSE},
        title="Variance vs Revenue (Color = Bucket, Size = |EBITDA%|)"
    )
    fig.update_layout(**base_layout(height=500))
    fig.update_xaxes(title="Net Revenue (₹)", **styled_axis(fmt="₹,.0f"))
    fig.update_yaxes(title="Variance (₹)", **styled_axis(fmt="₹,.0f"))
    fig.update_traces(marker=dict(opacity=0.7, line=dict(width=1, color="white")))
    st.plotly_chart(fig, use_container_width=True)

    # City-wise variance
    st.markdown('<div class="section-header">🏙️ City-wise Variance</div>', unsafe_allow_html=True)
    city_var = fdf.groupby("CITY").agg({
        "VARIANCE": ["mean", "sum", "count"],
        "NET REVENUE": "mean"
    }).reset_index()
    city_var.columns = ["CITY", "Avg Variance", "Total Variance", "Store Months", "Avg Revenue"]
    city_var = city_var.sort_values("Total Variance", ascending=True)

    fig_city = go.Figure()
    fig_city.add_trace(go.Bar(
        y=city_var["CITY"], x=city_var["Total Variance"]/1e5,
        orientation="h", marker_color=PURPLE, opacity=0.8,
        hovertemplate="%{y}<br>Total Var: ₹%{x:.1f}L<extra></extra>"
    ))
    fig_city.update_layout(
        **base_layout(title="Total Variance by City"),
        height=450,
        xaxis=styled_axis("Total Variance (₹ Lakhs)"),
        yaxis=dict(tickfont=dict(size=11, color="#1e293b"))
    )
    st.plotly_chart(fig_city, use_container_width=True)


def render_variance_data(fdf: pd.DataFrame):
    """Render variance data table."""
    st.markdown('<div class="section-header">📋 Variance Data</div>', unsafe_allow_html=True)

    display_cols = ["MONTH_STR", "STORE", "CITY", "NET REVENUE", "VARIANCE", "VARIANCE%", 
                    "VAR_BUCKET", "REV_BUCKET", "KITCHEN EBITDA", "EBITDA%"]
    display_df = fdf[display_cols].copy()
    display_df["NET REVENUE"] = display_df["NET REVENUE"].apply(lambda x: f"₹{x:,.0f}")
    display_df["VARIANCE"] = display_df["VARIANCE"].apply(lambda x: f"₹{x:,.0f}")
    display_df["VARIANCE%"] = display_df["VARIANCE%"].apply(lambda x: f"{x:.2f}%")
    display_df["KITCHEN EBITDA"] = display_df["KITCHEN EBITDA"].apply(lambda x: f"₹{x:,.0f}")
    display_df["EBITDA%"] = display_df["EBITDA%"].apply(lambda x: f"{x:.1f}%")

    st.dataframe(display_df, use_container_width=True, hide_index=True)

    csv = fdf.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Download CSV", csv, "variance_pnl_data.csv", "text/csv")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    """Main application entry point."""
    # Load data
    try:
        df = load_data(DATA_PATH)
    except Exception as e:
        st.error(f"❌ Error loading data: {e}")
        st.info("Please ensure 'Kittchen PNL Data.xlsx' exists in the app directory.")

        # Show sample data structure for reference
        st.markdown("""
        <div style="background: rgba(99,102,241,0.06); border: 1px solid rgba(99,102,241,0.15); border-radius: 12px; padding: 20px; margin-top: 20px;">
            <h4 style="color: #4f46e5; margin: 0 0 12px 0;">📋 Expected Data Structure</h4>
            <p style="color: #334155; font-size: 0.85rem; margin: 4px 0;">
                The Excel file should have columns: MONTH, STORE, CITY, ZONE MAPPING, STATUS, 
                NET REVENUE, GROSS MARGIN, KITCHEN EBITDA, VARIANCE, REVENUE COHORT, 
                CM COHORT, EBITDA CATEGORY, EBITDA COHORT
            </p>
        </div>
        """, unsafe_allow_html=True)
        return

    # Render sidebar
    dashboard, filters = render_sidebar(df)

    # Render selected dashboard
    if "Kitchen" in dashboard:
        render_kitchen_pnl(df, filters)
    else:
        render_variance_pnl(df, filters)

    # Footer
    st.markdown("""
    <div style="text-align: center; margin-top: 40px; padding: 20px 0; border-top: 1px solid rgba(99,102,241,0.15);">
        <p style="color: #94a3b8; font-size: 0.75rem; margin: 0;">
            Kitchen P&L Intelligence Suite v3.1 · Built with Streamlit + Plotly
        </p>
        <p style="color: #94a3b8; font-size: 0.7rem; margin: 4px 0 0 0;">
            Data refreshes every 5 minutes · White Theme
        </p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
