"""
╔══════════════════════════════════════════════════════════════╗
║         KITCHEN P&L INTELLIGENCE DASHBOARD                  ║
║         Cloud Kitchen Analytics Suite v3.0                  ║
╚══════════════════════════════════════════════════════════════╝

Python: 3.10+
Packages: streamlit>=1.32, plotly>=5.20, pandas>=2.0, openpyxl>=3.1, matplotlib>=3.8
Run: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np
from functools import lru_cache
import hashlib
from pathlib import Path

# Always resolve data file relative to this script
BASE_DIR  = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

# ─── PAGE CONFIG ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── THEME / CSS ────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp { 
    background: linear-gradient(135deg, #0a0a1a 0%, #12122e 40%, #0d1b3a 100%); 
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0a0a1a 0%, #120a2e 60%, #1a0a3a 100%);
    border-right: 1px solid rgba(139,92,246,0.25);
}
[data-testid="stSidebar"] label { 
    color: #c4b5fd !important; 
    font-weight: 600; 
    font-size: 0.78rem; 
    letter-spacing: 0.05em; 
}

.metric-card {
    background: linear-gradient(145deg, rgba(139,92,246,0.12) 0%, rgba(59,130,246,0.08) 50%, rgba(34,211,238,0.05) 100%);
    border: 1px solid rgba(139,92,246,0.35);
    border-radius: 20px;
    padding: 22px 18px;
    text-align: center;
    backdrop-filter: blur(16px);
    transition: all 0.35s cubic-bezier(0.4, 0, 0.2, 1);
    margin-bottom: 10px;
    position: relative;
    overflow: hidden;
}
.metric-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #8b5cf6, #22d3ee, #34d399);
    opacity: 0.6;
}
.metric-card:hover { 
    transform: translateY(-4px); 
    box-shadow: 0 12px 40px rgba(139,92,246,0.25), 0 0 60px rgba(139,92,246,0.08);
    border-color: rgba(139,92,246,0.6);
}
.metric-title { 
    color: #94a3b8; font-size: 0.68rem; font-weight: 700; 
    letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 8px; 
}
.metric-value { 
    color: #ffffff; font-size: 1.7rem; font-weight: 900; 
    text-shadow: 0 0 20px rgba(139,92,246,0.3);
}
.metric-delta { 
    font-size: 0.72rem; font-weight: 600; margin-top: 6px; letter-spacing: 0.03em;
}
.delta-pos { color: #34d399; }
.delta-neg { color: #f87171; }

.section-header {
    background: linear-gradient(90deg, rgba(139,92,246,0.18), rgba(34,211,238,0.05), transparent);
    border-left: 4px solid #8b5cf6;
    padding: 12px 18px;
    border-radius: 0 12px 12px 0;
    margin: 24px 0 14px 0;
    color: #e2e8f0;
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
    background: linear-gradient(90deg, #8b5cf6, transparent);
}

[data-testid="stTab"] { 
    color: #64748b; font-weight: 600; font-size: 0.9rem;
    padding: 10px 16px !important; transition: all 0.2s;
}
[data-testid="stTab"]:hover { color: #a78bfa; }
[data-testid="stTab"][aria-selected="true"] { 
    color: #c4b5fd; border-bottom: 3px solid #8b5cf6;
    background: linear-gradient(180deg, rgba(139,92,246,0.08), transparent);
}

.stDataFrame { border-radius: 14px; overflow: hidden; }

[data-testid="stExpander"] { 
    border: 1px solid rgba(139,92,246,0.25); border-radius: 12px; 
    background: rgba(139,92,246,0.03);
}

.filter-pill {
    display: inline-block;
    background: linear-gradient(135deg, rgba(139,92,246,0.2), rgba(59,130,246,0.15));
    border: 1px solid rgba(139,92,246,0.4);
    border-radius: 24px;
    padding: 4px 14px;
    font-size: 0.72rem; color: #c4b5fd;
    margin: 3px; font-weight: 500;
}

hr { border-color: rgba(139,92,246,0.2) !important; }

.brand {
    text-align: center; padding: 20px 0 28px 0; position: relative;
}
.brand::after {
    content: '';
    position: absolute; bottom: 10px; left: 25%; right: 25%;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(139,92,246,0.4), transparent);
}
.brand-title { 
    color: #ffffff; font-size: 1.15rem; font-weight: 900; 
    letter-spacing: 0.04em;
    text-shadow: 0 0 30px rgba(139,92,246,0.4);
}
.brand-sub { 
    color: #6d6d8a; font-size: 0.7rem; letter-spacing: 0.1em; margin-top: 4px;
}

::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: rgba(139,92,246,0.05); }
::-webkit-scrollbar-thumb { background: rgba(139,92,246,0.3); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: rgba(139,92,246,0.5); }

.stButton > button {
    background: linear-gradient(135deg, rgba(139,92,246,0.2), rgba(59,130,246,0.15)) !important;
    border: 1px solid rgba(139,92,246,0.4) !important;
    border-radius: 10px !important; color: #c4b5fd !important;
    font-weight: 600 !important; transition: all 0.3s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, rgba(139,92,246,0.35), rgba(59,130,246,0.25)) !important;
    box-shadow: 0 4px 20px rgba(139,92,246,0.3) !important;
    transform: translateY(-1px);
}

[data-testid="stAlert"] {
    border-radius: 12px !important;
    border: 1px solid rgba(139,92,246,0.2) !important;
}

@keyframes gradient-shift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}

.header-glow {
    background: linear-gradient(90deg, #8b5cf6, #22d3ee, #34d399, #8b5cf6);
    background-size: 300% 300%;
    animation: gradient-shift 8s ease infinite;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
</style>
""", unsafe_allow_html=True)


# ─── DATA LOADING & CACHING ─────────────────────────────────────────────────
@st.cache_data(ttl=300, show_spinner=False)
def load_data(filepath) -> pd.DataFrame:
    """Load and preprocess the kitchen PNL data."""
    df = pd.read_excel(filepath, header=1)

    # ── Derived columns ──
    df["GM%"]      = (df["GROSS MARGIN"]   / df["NET REVENUE"] * 100).round(2)
    df["EBITDA%"]  = (df["KITCHEN EBITDA"] / df["NET REVENUE"] * 100).round(2)
    df["VARIANCE%"]= (df["VARIANCE"]       / df["NET REVENUE"] * 100).round(4)

    df["CM"]  = df["GROSS MARGIN"]
    df["CM%"] = df["GM%"]

    # ── Month ordering ──
    month_order = ["Oct-2023","Nov-2023","Dec-2023","Jan-2024","Feb-2024","Mar-2024"]
    df["MONTH"] = pd.Categorical(df["MONTH"], categories=month_order, ordered=True)
    df = df.sort_values("MONTH")
    df["MONTH_STR"] = df["MONTH"].astype(str)

    # ── Variance buckets ──
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

    # ── Net Revenue buckets ──
    def rev_bucket(r):
        lacs = r / 1e5
        if lacs < 15:   return "(a) Below INR 15 lacs"
        elif lacs < 25: return "(b) INR 15 to 25 lacs"
        elif lacs < 35: return "(c) INR 25 to 35 lacs"
        elif lacs < 45: return "(d) INR 35 to 45 lacs"
        else:           return "(e) Above INR 45 lacs"
    df["REV_BUCKET"] = df["NET REVENUE"].apply(rev_bucket)

    # ── EBITDA Category label ──
    df["EBITDA_SIGN"] = df["KITCHEN EBITDA"].apply(lambda x: "▲ Positive" if x >= 0 else "▼ Negative")

    return df


def file_hash(path: str) -> str:
    """Quick hash for cache invalidation on file change."""
    with open(path, "rb") as f:
        return hashlib.md5(f.read()).hexdigest()


# ─── HELPER: CHART DEFAULTS ─────────────────────────────────────────────────
CHART_BG   = "rgba(0,0,0,0)"
PAPER_BG   = "rgba(0,0,0,0)"
FONT_COLOR = "#e2e8f0"
GRID_COLOR = "rgba(139,92,246,0.15)"
PURPLE     = "#8b5cf6"
CYAN       = "#22d3ee"
EMERALD    = "#34d399"
AMBER      = "#fbbf24"
ROSE       = "#f87171"
COLORS     = [PURPLE, CYAN, EMERALD, AMBER, ROSE, "#a78bfa", "#67e8f9", "#6ee7b7", "#f472b6", "#a3e635"]

def base_layout(**kwargs):
    return dict(
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=CHART_BG,
        font=dict(family="Inter", color=FONT_COLOR, size=12),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(
            bgcolor="rgba(0,0,0,0.3)", 
            bordercolor="rgba(139,92,246,0.3)", 
            borderwidth=1,
            font=dict(size=11)
        ),
        **kwargs
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


# ─── SAFE STYLING HELPERS (NO MATPLOTLIB DEPENDENCY) ────────────────────────
def safe_style_apply(styler, func, subset=None):
    """Safely apply styling function, handling pandas version differences."""
    try:
        # pandas >= 2.1
        return styler.map(func, subset=subset)
    except (AttributeError, TypeError):
        try:
            # pandas < 2.1
            return styler.applymap(func, subset=subset)
        except Exception:
            return styler


def color_scale(val, min_val, max_val, low_color="#f87171", mid_color="#fbbf24", high_color="#34d399"):
    """Generate color based on value position in range (no matplotlib)."""
    if pd.isna(val):
        return ""
    if max_val == min_val:
        return f"background-color: {mid_color}33"

    ratio = (val - min_val) / (max_val - min_val)
    ratio = max(0, min(1, ratio))

    # Simple 3-stop interpolation
    if ratio < 0.5:
        r = ratio * 2
        # Interpolate low to mid
        return f"background-color: rgba({int(248 + (251-248)*r)}, {int(113 + (191-113)*r)}, {int(113 + (24-113)*r)}, 0.25)"
    else:
        r = (ratio - 0.5) * 2
        # Interpolate mid to high
        return f"background-color: rgba({int(251 + (52-251)*r)}, {int(191 + (211-191)*r)}, {int(24 + (148-24)*r)}, 0.25)"


# ─── SIDEBAR ────────────────────────────────────────────────────────────────
def render_sidebar(df: pd.DataFrame):
    st.sidebar.markdown("""
    <div class='brand'>
        <div class='brand-title'>🍳 Kitchen Intelligence</div>
        <div class='brand-sub'>P&L Analytics Suite · v3.0</div>
    </div>
    """, unsafe_allow_html=True)
    st.sidebar.markdown("---")

    # ── Dashboard selector ──
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

    else:  # Variance dashboard
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
        "<div style='text-align:center; color:#4a4a6a; font-size:0.68rem;'>Data refreshes every 5 min<br>📊 Kitchen P&L Suite v3.0</div>",
        unsafe_allow_html=True
    )

    return dashboard, filters


# ─── FILTER APPLICATION ──────────────────────────────────────────────────────
def apply_pnl_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    mask = pd.Series([True] * len(df), index=df.index)
    if f.get("month"):
        mask &= df["MONTH_STR"].isin(f["month"])
    if f.get("store"):
        mask &= df["STORE"].isin(f["store"])
    if f.get("city"):
        mask &= df["CITY"].isin(f["city"])
    if f.get("zone"):
        mask &= df["ZONE MAPPING"].isin(f["zone"])
    if f.get("status"):
        mask &= df["STATUS"].isin(f["status"])
    if f.get("rev_cohort"):
        mask &= df["REVENUE COHORT"].isin(f["rev_cohort"])
    if f.get("cm_cohort"):
        mask &= df["CM COHORT"].isin(f["cm_cohort"])
    if f.get("ebitda_cat"):
        mask &= df["EBITDA CATEGORY"].isin(f["ebitda_cat"])
    if f.get("ebitda_cohort"):
        mask &= df["EBITDA COHORT"].isin(f["ebitda_cohort"])
    if f.get("rev_range"):
        mask &= df["NET REVENUE"].between(*f["rev_range"])
    if f.get("cm_range"):
        mask &= df["CM"].between(*f["cm_range"])
    if f.get("ebitda_range"):
        mask &= df["KITCHEN EBITDA"].between(*f["ebitda_range"])
    return df[mask]


def apply_var_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    mask = pd.Series([True] * len(df), index=df.index)
    if f.get("var_bucket"):
        mask &= df["VAR_BUCKET"].isin(f["var_bucket"])
    if f.get("city"):
        mask &= df["CITY"].isin(f["city"])
    if f.get("month"):
        mask &= df["MONTH_STR"].isin(f["month"])
    return df[mask]


# ═══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD 1 — KITCHEN LEVEL PNL
# ═══════════════════════════════════════════════════════════════════════════════
def render_kpi_strip(fdf: pd.DataFrame):
    """Top KPI strip: 6 metric cards with enhanced visuals."""
    total_rev    = fdf["NET REVENUE"].sum()
    avg_gm_pct   = fdf["GM%"].mean()
    avg_ebitda   = fdf["EBITDA%"].mean()
    total_ebitda = fdf["KITCHEN EBITDA"].sum()
    n_stores     = fdf["STORE"].nunique()
    n_active     = fdf[fdf["STATUS"] == "Active"]["STORE"].nunique()

    def card(title, value, delta="", delta_pos=True, icon=""):
        dclass = "delta-pos" if delta_pos else "delta-neg"
        return f"""
        <div class='metric-card'>
            <div class='metric-title'>{icon} {title}</div>
            <div class='metric-value'>{value}</div>
            {"<div class='metric-delta " + dclass + "'>" + delta + "</div>" if delta else ""}
        </div>"""

    cols = st.columns(6)
    data = [
        ("Net Revenue", f"₹{total_rev/1e7:.1f}Cr", "Total across selection", True, "💰"),
        ("Avg GM%",     f"{avg_gm_pct:.1f}%",       "Gross Margin %", avg_gm_pct > 50, "📊"),
        ("Avg EBITDA%", f"{avg_ebitda:.1f}%",        "Kitchen EBITDA %", avg_ebitda > 0, "📈"),
        ("Total EBITDA",f"₹{total_ebitda/1e7:.2f}Cr","Absolute EBITDA", total_ebitda > 0, "🏦"),
        ("Stores",      str(n_stores),               "Unique kitchens", True, "🏪"),
        ("Active",       str(n_active),               f"{n_active/max(n_stores,1)*100:.0f}% active", True, "✅"),
    ]
    for col, (title, val, delta, pos, icon) in zip(cols, data):
        with col:
            st.markdown(card(title, val, delta, pos, icon), unsafe_allow_html=True)


def render_pnl_snapshot(fdf: pd.DataFrame):
    """Pivoted PNL snapshot table — stores x months. FIXED: No matplotlib dependency."""
    st.markdown("<div class='section-header'>📋 Kitchen Snapshot — PNL Grid</div>", unsafe_allow_html=True)

    pivot = fdf.groupby(["STORE", "MONTH_STR"]).agg(
        NET_REVENUE=("NET REVENUE", "sum"),
        GM_PCT=("GM%", "mean"),
        CM_PCT=("CM%", "mean"),
        EBITDA=("KITCHEN EBITDA", "sum"),
        EBITDA_PCT=("EBITDA%", "mean"),
    ).reset_index()

    months = sorted(fdf["MONTH_STR"].unique(), 
                   key=lambda x: list(fdf["MONTH"].cat.categories).index(x) if x in fdf["MONTH"].cat.categories else 0)

    rows = []
    for store in sorted(pivot["STORE"].unique()):
        row = {"STORE": store}
        for m in months:
            sub = pivot[(pivot["STORE"] == store) & (pivot["MONTH_STR"] == m)]
            if len(sub):
                row[f"{m} | Net Rev (₹L)"] = round(sub["NET_REVENUE"].values[0] / 1e5, 1)
                row[f"{m} | GM%"]           = round(sub["GM_PCT"].values[0], 1)
                row[f"{m} | CM%"]           = round(sub["CM_PCT"].values[0], 1)
                row[f"{m} | EBITDA (₹L)"]  = round(sub["EBITDA"].values[0] / 1e5, 1)
                row[f"{m} | EBITDA%"]       = round(sub["EBITDA_PCT"].values[0], 1)
            else:
                for c in ["Net Rev (₹L)", "GM%", "CM%", "EBITDA (₹L)", "EBITDA%"]:
                    row[f"{m} | {c}"] = None
        rows.append(row)

    table_df = pd.DataFrame(rows)

    # Styling function for EBITDA columns
    def color_ebitda(val):
        if pd.isna(val): 
            return "color: #475569"
        if isinstance(val, (int, float)):
            if val < 0:  
                return "color: #f87171; font-weight: 700"
            elif val > 10: 
                return "color: #34d399; font-weight: 700"
            elif val > 5:
                return "color: #22d3ee; font-weight: 600"
        return "color: #e2e8f0"

    # Apply styling safely
    styler = table_df.style
    ebitda_cols = [c for c in table_df.columns if "EBITDA" in c]

    # Use our safe helper
    for col in ebitda_cols:
        styler = safe_style_apply(styler, color_ebitda, subset=[col])

    # Add number formatting
    styler = styler.format({c: "{:.1f}" for c in table_df.columns if c != "STORE" and "GM%" not in c and "CM%" not in c and "EBITDA%" not in c})
    styler = styler.format({c: "{:.1f}%" for c in table_df.columns if "GM%" in c or "CM%" in c or "EBITDA%" in c})

    st.dataframe(styler, use_container_width=True, height=420)


def render_trend_charts(fdf: pd.DataFrame):
    """Revenue & margin trend by month - ENHANCED."""
    st.markdown("<div class='section-header'>📈 Trend Analysis</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    monthly = fdf.groupby("MONTH_STR").agg(
        NET_REVENUE=("NET REVENUE", "sum"),
        GM_PCT=("GM%", "mean"),
        EBITDA_PCT=("EBITDA%", "mean"),
        CM_PCT=("CM%", "mean"),
        EBITDA=("KITCHEN EBITDA", "sum"),
    ).reset_index()

    cat_order = [m for m in fdf["MONTH"].cat.categories if m in monthly["MONTH_STR"].values]
    monthly["_sort"] = monthly["MONTH_STR"].map({m: i for i, m in enumerate(cat_order)})
    monthly = monthly.sort_values("_sort")

    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=monthly["MONTH_STR"], y=monthly["NET_REVENUE"] / 1e5,
            name="Net Revenue (₹L)", 
            marker=dict(
                color=monthly["NET_REVENUE"],
                colorscale=[[0, PURPLE], [0.5, CYAN], [1, EMERALD]],
                showscale=False,
                line=dict(color="rgba(139,92,246,0.5)", width=1)
            ),
            opacity=0.9,
            hovertemplate="<b>%{x}</b><br>Net Revenue: ₹%{y:.1f}L<extra></extra>"
        ))
        fig.add_trace(go.Scatter(
            x=monthly["MONTH_STR"], y=monthly["EBITDA"] / 1e5,
            name="EBITDA (₹L)", mode="lines+markers",
            line=dict(color=EMERALD, width=3),
            marker=dict(size=10, symbol="diamond", line=dict(color="#1a1a2e", width=2)),
            hovertemplate="<b>%{x}</b><br>EBITDA: ₹%{y:.1f}L<extra></extra>"
        ))
        fig.update_layout(
            title=dict(text="📊 Revenue & EBITDA Trend", font=dict(size=16, color="#e2e8f0")),
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("₹ (Lacs)"),
            **base_layout(barmode="overlay")
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure()
        for metric, color, name in [
            ("GM_PCT", CYAN, "GM%"), 
            ("CM_PCT", AMBER, "CM%"), 
            ("EBITDA_PCT", EMERALD, "EBITDA%")
        ]:
            fig2.add_trace(go.Scatter(
                x=monthly["MONTH_STR"], y=monthly[metric],
                name=name, mode="lines+markers",
                line=dict(color=color, width=3),
                marker=dict(size=8, line=dict(color="#1a1a2e", width=1.5)),
                hovertemplate=f"<b>%{{x}}</b><br>{name}: %{{y:.1f}}%<extra></extra>"
            ))
        fig2.update_layout(
            title=dict(text="📈 Margin % Trends", font=dict(size=16, color="#e2e8f0")),
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("Margin %", "%"),
            **base_layout()
        )
        st.plotly_chart(fig2, use_container_width=True)


def render_city_zone_charts(fdf: pd.DataFrame):
    """City and zone breakdown - ENHANCED."""
    st.markdown("<div class='section-header'>🗺️ City & Zone Breakdown</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    city_agg = fdf.groupby("CITY").agg(
        NET_REVENUE=("NET REVENUE", "sum"),
        EBITDA=("KITCHEN EBITDA", "sum"),
        GM_PCT=("GM%", "mean"),
        EBITDA_PCT=("EBITDA%", "mean"),
        STORES=("STORE", "nunique"),
    ).reset_index().sort_values("NET_REVENUE", ascending=False)

    with col1:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=city_agg["CITY"], y=city_agg["NET_REVENUE"] / 1e7,
            name="Net Revenue", 
            marker=dict(
                color=city_agg["NET_REVENUE"],
                colorscale=[[0, PURPLE], [1, CYAN]],
                showscale=False,
                line=dict(color="rgba(139,92,246,0.4)", width=1)
            ),
            text=[f"₹{v:.1f}Cr" for v in city_agg["NET_REVENUE"] / 1e7],
            textposition="outside",
            textfont=dict(color="#e2e8f0", size=11),
            hovertemplate="<b>%{x}</b><br>Net Revenue: ₹%{y:.2f}Cr<br>Stores: %{customdata}<extra></extra>",
            customdata=city_agg["STORES"]
        ))
        fig.update_layout(
            title=dict(text="💰 Net Revenue by City", font=dict(size=16, color="#e2e8f0")),
            xaxis=styled_axis("City", showgrid=False),
            yaxis=styled_axis("Revenue (₹ Cr)"),
            showlegend=False,
            **base_layout()
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = go.Figure(go.Pie(
            labels=city_agg["CITY"],
            values=city_agg["EBITDA"],
            hole=0.6,
            marker=dict(
                colors=COLORS[:len(city_agg)],
                line=dict(color="rgba(15,15,26,0.8)", width=2)
            ),
            textinfo="label+percent",
            textfont=dict(color="#e2e8f0", size=11),
            hovertemplate="<b>%{label}</b><br>EBITDA: ₹%{value:,.0f}<br>Share: %{percent}<extra></extra>"
        ))
        fig2.add_annotation(
            text=f"<b>₹{city_agg['EBITDA'].sum()/1e7:.1f}Cr</b><br><span style='font-size:11px;color:#94a3b8'>Total EBITDA</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=15, color="#ffffff")
        )
        fig2.update_layout(title=dict(text="🏦 EBITDA Share by City", font=dict(size=16, color="#e2e8f0")), **base_layout())
        st.plotly_chart(fig2, use_container_width=True)


def render_store_scatter(fdf: pd.DataFrame):
    """Bubble scatter: stores by Revenue vs EBITDA% - ENHANCED."""
    st.markdown("<div class='section-header'>🔮 Store Intelligence — Revenue vs EBITDA</div>", unsafe_allow_html=True)

    store_agg = fdf.groupby(["STORE", "CITY", "ZONE MAPPING", "STATUS"]).agg(
        NET_REVENUE=("NET REVENUE", "mean"),
        EBITDA_PCT=("EBITDA%", "mean"),
        GM_PCT=("GM%", "mean"),
        VARIANCE_PCT=("VARIANCE%", "mean"),
    ).reset_index()

    fig = px.scatter(
        store_agg,
        x="NET_REVENUE", y="EBITDA_PCT",
        color="CITY", size="NET_REVENUE",
        hover_name="STORE",
        hover_data={"NET_REVENUE": ":,.0f", "EBITDA_PCT": ":.1f", "GM_PCT": ":.1f", "CITY": True},
        size_max=35,
        color_discrete_sequence=COLORS,
        symbol="STATUS",
        labels={"NET_REVENUE": "Net Revenue (₹)", "EBITDA_PCT": "EBITDA %"},
        title="Store Portfolio: Revenue vs EBITDA%"
    )
    fig.add_hline(y=0, line_color=ROSE, line_dash="dash", line_width=2,
                  annotation_text="Break-even", annotation_font_color=ROSE,
                  annotation_font_size=12)
    fig.add_vline(x=store_agg["NET_REVENUE"].median(), line_color="rgba(139,92,246,0.3)", 
                  line_dash="dot", line_width=1,
                  annotation_text="Median Revenue", annotation_font_color="#8b5cf6",
                  annotation_font_size=10)
    fig.update_layout(**base_layout(
        xaxis=styled_axis("Net Revenue (₹)"),
        yaxis=styled_axis("EBITDA %"),
        height=520
    ))
    st.plotly_chart(fig, use_container_width=True)


def render_cohort_heatmap(fdf: pd.DataFrame):
    """Heatmap: Revenue Cohort x Month - ENHANCED."""
    st.markdown("<div class='section-header'>🧊 Cohort Performance Heatmap</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        pivot = fdf.groupby(["REVENUE COHORT", "MONTH_STR"])["EBITDA%"].mean().round(1).unstack("MONTH_STR")
        fig = go.Figure(go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale=[[0, "#7f1d1d"], [0.2, "#f87171"], [0.4, "#fbbf24"], [0.7, "#34d399"], [1, "#22d3ee"]],
            text=[[f"{v:.1f}%" if not np.isnan(v) else "" for v in row] for row in pivot.values],
            texttemplate="%{text}",
            textfont=dict(size=10, color="#ffffff"),
            hovertemplate="Revenue Cohort: %{y}<br>Month: %{x}<br>Avg EBITDA%: %{z:.1f}%<extra></extra>",
            colorbar=dict(
                title="EBITDA%", tickfont=dict(color=FONT_COLOR), 
                title_font=dict(color=FONT_COLOR),
                bgcolor="rgba(0,0,0,0.3)"
            )
        ))
        fig.update_layout(title=dict(text="Revenue Cohort x Month → Avg EBITDA%", font=dict(size=14)), **base_layout(height=340))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        pivot2 = fdf.groupby(["EBITDA CATEGORY", "MONTH_STR"])["STORE"].nunique().unstack("MONTH_STR").fillna(0)
        fig2 = go.Figure(go.Heatmap(
            z=pivot2.values,
            x=pivot2.columns.tolist(),
            y=pivot2.index.tolist(),
            colorscale=[[0, "#1e1b4b"], [0.5, PURPLE], [1, EMERALD]],
            text=[[str(int(v)) for v in row] for row in pivot2.values],
            texttemplate="%{text}",
            textfont=dict(size=10, color="#ffffff"),
            hovertemplate="EBITDA Cat: %{y}<br>Month: %{x}<br>Stores: %{z}<extra></extra>",
            colorbar=dict(
                title="# Stores", tickfont=dict(color=FONT_COLOR), 
                title_font=dict(color=FONT_COLOR),
                bgcolor="rgba(0,0,0,0.3)"
            )
        ))
        fig2.update_layout(title=dict(text="EBITDA Category x Month → Store Count", font=dict(size=14)), **base_layout(height=340))
        st.plotly_chart(fig2, use_container_width=True)


def render_top_stores(fdf: pd.DataFrame):
    """Top/bottom store rankings - ENHANCED."""
    st.markdown("<div class='section-header'>🏆 Store Rankings</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    store_perf = fdf.groupby(["STORE", "CITY"]).agg(
        AVG_EBITDA_PCT=("EBITDA%", "mean"),
        TOTAL_REVENUE=("NET REVENUE", "sum"),
        AVG_GM_PCT=("GM%", "mean"),
    ).reset_index().round(2)

    with col1:
        top10 = store_perf.nlargest(10, "AVG_EBITDA_PCT")
        fig = go.Figure(go.Bar(
            x=top10["AVG_EBITDA_PCT"], y=top10["STORE"],
            orientation="h",
            marker=dict(
                color=top10["AVG_EBITDA_PCT"],
                colorscale=[[0, EMERALD], [1, CYAN]],
                showscale=False,
                line=dict(color="rgba(52,211,153,0.5)", width=1)
            ),
            text=[f"{v:.1f}%" for v in top10["AVG_EBITDA_PCT"]],
            textposition="outside",
            textfont=dict(color="#e2e8f0", size=10),
            hovertemplate="<b>%{y}</b><br>Avg EBITDA%: %{x:.1f}%<extra></extra>"
        ))
        fig.update_layout(
            title=dict(text="🟢 Top 10 Stores by Avg EBITDA%", font=dict(size=14, color="#34d399")),
            xaxis=styled_axis("Avg EBITDA%"),
            yaxis=dict(title="", tickfont=dict(color="#94a3b8", size=10), autorange="reversed"),
            **base_layout(height=400)
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        bot10 = store_perf.nsmallest(10, "AVG_EBITDA_PCT")
        fig2 = go.Figure(go.Bar(
            x=bot10["AVG_EBITDA_PCT"], y=bot10["STORE"],
            orientation="h",
            marker=dict(
                color=bot10["AVG_EBITDA_PCT"],
                colorscale=[[0, ROSE], [1, AMBER]],
                showscale=False,
                line=dict(color="rgba(248,113,113,0.5)", width=1)
            ),
            text=[f"{v:.1f}%" for v in bot10["AVG_EBITDA_PCT"]],
            textposition="outside",
            textfont=dict(color="#e2e8f0", size=10),
            hovertemplate="<b>%{y}</b><br>Avg EBITDA%: %{x:.1f}%<extra></extra>"
        ))
        fig2.update_layout(
            title=dict(text="🔴 Bottom 10 Stores by Avg EBITDA%", font=dict(size=14, color="#f87171")),
            xaxis=styled_axis("Avg EBITDA%"),
            yaxis=dict(title="", tickfont=dict(color="#94a3b8", size=10), autorange="reversed"),
            **base_layout(height=400)
        )
        st.plotly_chart(fig2, use_container_width=True)


def render_kitchen_pnl(df: pd.DataFrame, filters: dict):
    fdf = apply_pnl_filters(df, filters)

    if fdf.empty:
        st.warning("⚠️ No data matches the current filters. Please adjust your selections.")
        return

    # Active filter pills
    active = []
    if filters.get("store"):      active.append(f"Stores: {len(filters['store'])}")
    if filters.get("city") and len(filters["city"]) < df["CITY"].nunique(): 
        active.append(f"Cities: {', '.join(filters['city'])}")
    if filters.get("month") and len(filters["month"]) < df["MONTH_STR"].nunique():
        active.append(f"Months: {len(filters['month'])}")
    if active:
        pills = "".join([f"<span class='filter-pill'>{a}</span>" for a in active])
        st.markdown(f"<div style='margin-bottom:12px;'>Active Filters: {pills}</div>", unsafe_allow_html=True)

    render_kpi_strip(fdf)
    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📋 Snapshot Table", "📈 Trends", "🗺️ City/Zone", "🔮 Store Map", "🏆 Rankings"
    ])
    with tab1:
        render_pnl_snapshot(fdf)
    with tab2:
        render_trend_charts(fdf)
        render_cohort_heatmap(fdf)
    with tab3:
        render_city_zone_charts(fdf)
    with tab4:
        render_store_scatter(fdf)
    with tab5:
        render_top_stores(fdf)


# ═══════════════════════════════════════════════════════════════════════════════
#  DASHBOARD 2 — VARIANCE LEVEL PNL
# ═══════════════════════════════════════════════════════════════════════════════
REV_BUCKETS_ORDER = [
    "(a) Below INR 15 lacs",
    "(b) INR 15 to 25 lacs",
    "(c) INR 25 to 35 lacs",
    "(d) INR 35 to 45 lacs",
    "(e) Above INR 45 lacs",
]
VAR_BUCKET_ORDER = [
    "(a) Var < ₹15K",
    "(b) Var ₹15K-20K",
    "(c) Var ₹20K-25K",
    "(d) Var > ₹25K",
]
VAR_COLORS = {
    "(a) Var < ₹15K":    EMERALD,
    "(b) Var ₹15K-20K":  AMBER,
    "(c) Var ₹20K-25K":  "#fb923c",
    "(d) Var > ₹25K":    ROSE,
}


def render_variance_kpis(fdf: pd.DataFrame):
    total_stores   = fdf["STORE"].nunique()
    avg_var_pct    = fdf["VARIANCE%"].mean()
    avg_var_abs    = fdf["VARIANCE"].mean()
    pct_high_var   = (fdf["VARIANCE%"] > 3).mean() * 100

    cols = st.columns(4)
    cards = [
        ("🏪 Stores (filtered)", str(total_stores), "", True),
        ("📉 Avg Variance %",    f"{avg_var_pct:.2f}%", "of Net Revenue", avg_var_pct < 3),
        ("💸 Avg Variance (₹)",  f"₹{avg_var_abs:,.0f}", "per store-month", True),
        ("⚠️ High Var Stores",   f"{pct_high_var:.1f}%", "variance > 3%", pct_high_var < 30),
    ]
    for col, (title, val, delta, pos) in zip(cols, cards):
        dclass = "delta-pos" if pos else "delta-neg"
        with col:
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-title'>{title}</div>
                <div class='metric-value'>{val}</div>
                {"<div class='metric-delta " + dclass + "'>" + delta + "</div>" if delta else ""}
            </div>""", unsafe_allow_html=True)


def render_sub1_avg_variance(fdf: pd.DataFrame):
    """Sub-dashboard 1: Avg Variance % by Revenue Category x Month."""
    st.markdown("<div class='section-header'>📊 Sub-Dashboard 1 · Average Variance % by Revenue Category</div>", unsafe_allow_html=True)

    pivot = fdf.groupby(["REV_BUCKET", "MONTH_STR"])["VARIANCE%"].mean().round(4).unstack("MONTH_STR")
    months = [m for m in fdf["MONTH"].cat.categories if m in pivot.columns]
    pivot  = pivot.reindex(columns=months, fill_value=np.nan)
    pivot  = pivot.reindex(REV_BUCKETS_ORDER)

    # Grand total row
    grand = fdf.groupby("MONTH_STR")["VARIANCE%"].mean().round(4)
    grand = grand.reindex(months)
    pivot.loc["Grand Total"] = grand

    # ── Formatted display table ──
    display = pivot.copy()
    for col in display.columns:
        display[col] = display[col].apply(
            lambda v: f"{v:.2f}%" if not pd.isna(v) else "—"
        )

    # Style for grand total row
    def highlight_grand(row):
        if row.name == "Grand Total":
            return ["background-color: rgba(139,92,246,0.25); font-weight: bold; color: #c4b5fd"] * len(row)
        return [""] * len(row)

    styled = display.style.apply(highlight_grand, axis=1)
    st.dataframe(styled, use_container_width=True)

    # ── Visual: grouped bar per revenue bucket ──
    fig = go.Figure()
    bucket_colors = [EMERALD, CYAN, AMBER, PURPLE, ROSE]
    for i, rb in enumerate(REV_BUCKETS_ORDER):
        if rb in pivot.index:
            row = pivot.loc[rb]
            fig.add_trace(go.Bar(
                name=rb,
                x=months,
                y=row.values,
                marker=dict(color=bucket_colors[i % len(bucket_colors)], line=dict(color="rgba(255,255,255,0.1)", width=1)),
                hovertemplate=f"<b>{rb}</b><br>Month: %{{x}}<br>Avg Variance: %{{y:.3f}}%<extra></extra>"
            ))
    fig.update_layout(
        title=dict(text="Avg Variance % by Revenue Category & Month", font=dict(size=15)),
        barmode="group",
        xaxis=styled_axis("Month", showgrid=False),
        yaxis=styled_axis("Avg Variance %"),
        **base_layout(height=440)
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── Heatmap view ──
    with st.expander("🧊 View as Heatmap"):
        numeric_pivot = pivot.drop(index="Grand Total", errors="ignore")
        fig2 = go.Figure(go.Heatmap(
            z=numeric_pivot.values.astype(float),
            x=numeric_pivot.columns.tolist(),
            y=numeric_pivot.index.tolist(),
            colorscale=[[0, EMERALD], [0.5, AMBER], [1, ROSE]],
            text=[[f"{v:.3f}%" if not np.isnan(v) else "" for v in row] for row in numeric_pivot.values.astype(float)],
            texttemplate="%{text}",
            textfont=dict(size=10),
            hovertemplate="Rev Bucket: %{y}<br>Month: %{x}<br>Avg Var%: %{z:.3f}%<extra></extra>",
            colorbar=dict(title="Var%", tickfont=dict(color=FONT_COLOR), title_font=dict(color=FONT_COLOR),
                         bgcolor="rgba(0,0,0,0.3)")
        ))
        fig2.update_layout(title=dict(text="Heatmap: Avg Variance % by Revenue Bucket x Month", font=dict(size=14)), **base_layout(height=340))
        st.plotly_chart(fig2, use_container_width=True)


def render_sub2_store_count(fdf: pd.DataFrame):
    """Sub-dashboard 2: Count of stores by Revenue Bucket x Month. 
    FIXED: Replaced matplotlib-dependent .background_gradient() with custom color scaling."""
    st.markdown("<div class='section-header'>🏪 Sub-Dashboard 2 · Kitchen Store Count by Revenue Range x Month</div>", unsafe_allow_html=True)

    pivot = fdf.groupby(["REV_BUCKET", "MONTH_STR"])["STORE"].count().unstack("MONTH_STR").fillna(0).astype(int)
    months = [m for m in fdf["MONTH"].cat.categories if m in pivot.columns]
    pivot  = pivot.reindex(columns=months, fill_value=0)
    pivot  = pivot.reindex(REV_BUCKETS_ORDER, fill_value=0)

    # Grand total row
    grand = fdf.groupby("MONTH_STR")["STORE"].count().reindex(months, fill_value=0)
    pivot.loc["Grand Total"] = grand

    # ── Custom color scaling (NO matplotlib dependency) ──
    numeric_pivot = pivot.copy()
    max_val = numeric_pivot.drop(index="Grand Total", errors="ignore").values.max()
    min_val = numeric_pivot.drop(index="Grand Total", errors="ignore").values.min()

    def color_scale_style(val):
        """Apply color intensity based on value."""
        if pd.isna(val) or val == 0:
            return ""
        if max_val == min_val:
            return "background-color: rgba(139,92,246,0.15)"
        ratio = (val - min_val) / (max_val - min_val)
        intensity = 0.1 + (ratio * 0.35)  # Scale from 0.1 to 0.45 opacity
        return f"background-color: rgba(139,92,246,{intensity:.2f}); color: #e2e8f0"

    # Style: highlight grand total row + color scale on data cells
    def highlight_grand2(row):
        if row.name == "Grand Total":
            return ["background-color: rgba(139,92,246,0.25); font-weight: bold; color: #c4b5fd"] * len(row)
        return [""] * len(row)

    styled_df = pivot.style.apply(highlight_grand2, axis=1)

    # Apply color scale to non-total rows, non-total columns
    data_cols = [c for c in pivot.columns if c != "Grand Total"]
    data_rows = [r for r in pivot.index if r != "Grand Total"]

    for col in data_cols:
        for idx in data_rows:
            val = pivot.loc[idx, col]
            if val > 0:
                styled_df = styled_df.map(
                    lambda x, v=val, mv=max_val, mn=min_val: color_scale_style(v) if x == v else "",
                    subset=pd.IndexSlice[[idx], [col]]
                )

    st.dataframe(styled_df, use_container_width=True)

    # ── Stacked bar chart ──
    fig = go.Figure()
    bucket_colors = [EMERALD, CYAN, AMBER, PURPLE, ROSE]
    numeric_pivot_chart = pivot.drop(index="Grand Total", errors="ignore")
    for i, rb in enumerate(REV_BUCKETS_ORDER):
        if rb in numeric_pivot_chart.index:
            fig.add_trace(go.Bar(
                name=rb,
                x=months,
                y=numeric_pivot_chart.loc[rb].values,
                marker=dict(color=bucket_colors[i % len(bucket_colors)], line=dict(color="rgba(255,255,255,0.1)", width=1)),
                hovertemplate=f"<b>{rb}</b><br>Month: %{{x}}<br>Store Count: %{{y}}<extra></extra>"
            ))
    fig.update_layout(
        title=dict(text="Kitchen Store Count by Revenue Range & Month", font=dict(size=15)),
        barmode="stack",
        xaxis=styled_axis("Month", showgrid=False),
        yaxis=styled_axis("Store Count"),
        **base_layout(height=440)
    )
    st.plotly_chart(fig, use_container_width=True)

    # ── Store distribution donut ──
    with st.expander("🔍 Revenue Bucket Distribution (All Months)"):
        bucket_counts = fdf["REV_BUCKET"].value_counts().reindex(REV_BUCKETS_ORDER).fillna(0)
        fig2 = go.Figure(go.Pie(
            labels=bucket_counts.index,
            values=bucket_counts.values,
            hole=0.6,
            marker=dict(colors=bucket_colors, line=dict(color="rgba(15,15,26,0.8)", width=2)),
            textinfo="label+percent+value",
            textfont=dict(color="#e2e8f0", size=11),
            hovertemplate="<b>%{label}</b><br>Records: %{value}<br>Share: %{percent}<extra></extra>"
        ))
        fig2.add_annotation(
            text=f"<b>{int(bucket_counts.sum())}</b><br><span style='font-size:11px;color:#94a3b8'>Total Records</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color="#ffffff")
        )
        fig2.update_layout(title=dict(text="Store-Month Records by Revenue Bucket", font=dict(size=14)), **base_layout())
        st.plotly_chart(fig2, use_container_width=True)


def render_variance_distribution(fdf: pd.DataFrame):
    """Variance bucket distribution & trend - ENHANCED."""
    st.markdown("<div class='section-header'>📉 Variance Distribution & Trends</div>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        bucket_dist = fdf["VAR_BUCKET"].value_counts().reindex(VAR_BUCKET_ORDER).fillna(0)
        fig = go.Figure(go.Bar(
            x=VAR_BUCKET_ORDER,
            y=bucket_dist.values,
            marker=dict(
                color=[VAR_COLORS[b] for b in VAR_BUCKET_ORDER],
                line=dict(color="rgba(255,255,255,0.1)", width=1)
            ),
            text=bucket_dist.values.astype(int),
            textposition="outside",
            textfont=dict(color="#e2e8f0", size=12),
            hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>"
        ))
        fig.update_layout(
            title=dict(text="Store-Month Distribution by Variance Bucket", font=dict(size=14)),
            xaxis=styled_axis("Variance Bucket", showgrid=False),
            yaxis=styled_axis("Count"),
            showlegend=False,
            **base_layout()
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        var_month = fdf.groupby(["MONTH_STR", "VAR_BUCKET"])["STORE"].count().unstack("VAR_BUCKET").fillna(0)
        months = [m for m in fdf["MONTH"].cat.categories if m in var_month.index]
        var_month = var_month.reindex(months)

        fig2 = go.Figure()
        for bucket in VAR_BUCKET_ORDER:
            if bucket in var_month.columns:
                color = VAR_COLORS[bucket]
                rgba_color = color.replace("rgb", "rgba").replace(")", ",0.15)") if "rgb" in color else color + "26"
                fig2.add_trace(go.Scatter(
                    x=months, y=var_month[bucket],
                    name=bucket, mode="lines+markers",
                    line=dict(color=color, width=2.5),
                    marker=dict(size=8, line=dict(color="#1a1a2e", width=1)),
                    fill="tozeroy",
                    fillcolor=rgba_color,
                    hovertemplate=f"<b>{bucket}</b><br>Month: %{{x}}<br>Stores: %{{y}}<extra></extra>"
                ))
        fig2.update_layout(
            title=dict(text="Variance Bucket Trend by Month", font=dict(size=14)),
            xaxis=styled_axis("Month"),
            yaxis=styled_axis("Store Count"),
            **base_layout()
        )
        st.plotly_chart(fig2, use_container_width=True)


def render_variance_city(fdf: pd.DataFrame):
    """Variance by city analysis - ENHANCED."""
    st.markdown("<div class='section-header'>🏙️ Variance by City</div>", unsafe_allow_html=True)

    city_var = fdf.groupby("CITY").agg(
        AVG_VAR_PCT=("VARIANCE%", "mean"),
        TOTAL_VAR=("VARIANCE", "sum"),
        STORES=("STORE", "nunique"),
    ).reset_index().sort_values("AVG_VAR_PCT", ascending=False)

    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(go.Bar(
        x=city_var["CITY"], y=city_var["AVG_VAR_PCT"],
        name="Avg Variance %", 
        marker=dict(
            color=city_var["AVG_VAR_PCT"],
            colorscale=[[0, EMERALD], [0.5, AMBER], [1, ROSE]],
            showscale=False,
            line=dict(color="rgba(255,255,255,0.1)", width=1)
        ),
        opacity=0.9
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=city_var["CITY"], y=city_var["TOTAL_VAR"] / 1e5,
        name="Total Variance (₹L)", mode="lines+markers",
        line=dict(color=AMBER, width=3), 
        marker=dict(size=9, symbol="diamond", line=dict(color="#1a1a2e", width=1.5))
    ), secondary_y=True)
    fig.update_layout(
        title=dict(text="Variance by City — % and Absolute", font=dict(size=15)),
        **base_layout(height=400)
    )
    fig.update_yaxes(title_text="Avg Variance %", **styled_axis(), secondary_y=False)
    fig.update_yaxes(title_text="Total Variance (₹ Lacs)", **styled_axis(), secondary_y=True)
    st.plotly_chart(fig, use_container_width=True)


def render_variance_pnl(df: pd.DataFrame, filters: dict):
    fdf = apply_var_filters(df, filters)

    if fdf.empty:
        st.warning("⚠️ No data matches the current filters.")
        return

    render_variance_kpis(fdf)
    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "📊 Avg Variance % Table", "🏪 Store Count Table", "📉 Distribution & City"
    ])
    with tab1:
        render_sub1_avg_variance(fdf)
    with tab2:
        render_sub2_store_count(fdf)
    with tab3:
        render_variance_distribution(fdf)
        render_variance_city(fdf)


# ═══════════════════════════════════════════════════════════════════════════════
#  BONUS: INSIGHTS SECTION
# ═══════════════════════════════════════════════════════════════════════════════
def render_insights(df: pd.DataFrame):
    """Auto-generated insights from data - ENHANCED."""
    st.markdown("---")
    st.markdown("<div class='section-header'>💡 Auto-Generated Insights</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    best_city = df.groupby("CITY")["EBITDA%"].mean().idxmax()
    best_ebitda = df.groupby("CITY")["EBITDA%"].mean().max()

    worst_city = df.groupby("CITY")["EBITDA%"].mean().idxmin()
    worst_ebitda = df.groupby("CITY")["EBITDA%"].mean().min()

    high_var = df[df["VARIANCE%"] > 3]["STORE"].nunique()
    total_s  = df["STORE"].nunique()

    with col1:
        st.info(f"🏆 **Best City**: {best_city} with avg EBITDA% of **{best_ebitda:.1f}%**")
    with col2:
        st.warning(f"⚠️ **Needs Attention**: {worst_city} with avg EBITDA% of **{worst_ebitda:.1f}%**")
    with col3:
        pct = high_var / total_s * 100
        if pct > 20:
            st.error(f"📉 **{high_var}** stores ({pct:.0f}%) have variance > 3%")
        elif pct > 10:
            st.warning(f"📉 **{high_var}** stores ({pct:.0f}%) have variance > 3%")
        else:
            st.success(f"📉 **{high_var}** stores ({pct:.0f}%) have variance > 3%")

    with st.expander("📊 View Full Insights Summary"):
        monthly_trend = df.groupby("MONTH_STR")["KITCHEN EBITDA"].sum()
        best_m = monthly_trend.idxmax()
        worst_m = monthly_trend.idxmin()

        # Additional insights
        active_pct = df[df['STATUS']=='Active']['STORE'].nunique() / df['STORE'].nunique() * 100
        avg_rev_per_store = df['NET REVENUE'].mean() / 1e5

        st.markdown(f"""
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px;">
        <div style="background: rgba(139,92,246,0.08); border: 1px solid rgba(139,92,246,0.2); border-radius: 12px; padding: 16px;">
            <h4 style="color: #c4b5fd; margin: 0 0 8px 0;">📅 Monthly Performance</h4>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Best Month</b>: <code>{best_m}</code> — ₹{monthly_trend[best_m]/1e7:.2f} Cr total EBITDA</p>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Worst Month</b>: <code>{worst_m}</code> — ₹{monthly_trend[worst_m]/1e7:.2f} Cr total EBITDA</p>
        </div>
        <div style="background: rgba(139,92,246,0.08); border: 1px solid rgba(139,92,246,0.2); border-radius: 12px; padding: 16px;">
            <h4 style="color: #c4b5fd; margin: 0 0 8px 0;">🏪 Store Metrics</h4>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Active Rate</b>: {active_pct:.0f}% ({df[df['STATUS']=='Active']['STORE'].nunique()} active / {df[df['STATUS']=='Inactive']['STORE'].nunique()} inactive)</p>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Avg Revenue/Store</b>: ₹{avg_rev_per_store:.1f}L per month</p>
        </div>
        <div style="background: rgba(139,92,246,0.08); border: 1px solid rgba(139,92,246,0.2); border-radius: 12px; padding: 16px;">
            <h4 style="color: #c4b5fd; margin: 0 0 8px 0;">💰 Financial Range</h4>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Revenue Range</b>: ₹{df['NET REVENUE'].min()/1e5:.1f}L – ₹{df['NET REVENUE'].max()/1e5:.1f}L per store-month</p>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>GM% Range</b>: {df['GM%'].min():.1f}% – {df['GM%'].max():.1f}%</p>
        </div>
        <div style="background: rgba(139,92,246,0.08); border: 1px solid rgba(139,92,246,0.2); border-radius: 12px; padding: 16px;">
            <h4 style="color: #c4b5fd; margin: 0 0 8px 0;">📈 Profitability</h4>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>EBITDA% Range</b>: {df['EBITDA%'].min():.1f}% – {df['EBITDA%'].max():.1f}%</p>
            <p style="color: #e2e8f0; margin: 4px 0;"><b>Profitable Stores</b>: {(df['KITCHEN EBITDA'] > 0).sum()} / {len(df)} store-months</p>
        </div>
        </div>
        """)


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    # ── Animated Header ──
    st.markdown("""
    <div style='text-align:center; padding: 12px 0 28px 0;'>
        <h1 class='header-glow' style='font-size:2.2rem; font-weight:900; margin:0; letter-spacing:0.02em;'>
            🍳 Kitchen P&L Intelligence Suite
        </h1>
        <p style='color:#6d6d8a; font-size:0.88rem; margin:8px 0 0 0; letter-spacing:0.08em; font-weight:500;'>
            CLOUD KITCHEN ANALYTICS · MULTI-CITY · REAL-TIME FILTERS
        </p>
        <div style="width:120px; height:3px; background: linear-gradient(90deg, #8b5cf6, #22d3ee); margin: 12px auto 0; border-radius: 2px;"></div>
    </div>
    """, unsafe_allow_html=True)

    # ── Load data ──
    with st.spinner("🔄 Loading kitchen data..."):
        try:
            df = load_data(DATA_PATH)
        except Exception as e:
            st.error(f"❌ Failed to load data: {e}")
            st.info("💡 Please ensure 'Kittchen PNL Data.xlsx' is in the same directory as app.py")
            return

    # ── Sidebar ──
    dashboard, filters = render_sidebar(df)

    # ── Route to correct dashboard ──
    if "Kitchen" in dashboard:
        render_kitchen_pnl(df, filters)
    else:
        render_variance_pnl(df, filters)

    # ── Auto insights ──
    render_insights(df)

    # ── Footer ──
    st.markdown("---")
    st.markdown("""
    <div style='text-align:center; padding: 8px 0; color: #475569; font-size: 0.7rem;'>
        <span style="color: #6d6d8a;">Kitchen P&L Intelligence Suite v3.0</span> · 
        Built with <span style="color: #8b5cf6;">Streamlit</span> + <span style="color: #22d3ee;">Plotly</span>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
