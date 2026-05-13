import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Kitchen P&L Intelligence",
    page_icon="🍳",
    layout="wide"
)

# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>

.stApp {
    background: linear-gradient(135deg,#0f172a,#111827);
    color: white;
}

.metric-card{
    background: rgba(255,255,255,0.05);
    border:1px solid rgba(255,255,255,0.08);
    padding:20px;
    border-radius:18px;
    text-align:center;
    backdrop-filter: blur(10px);
}

.metric-title{
    color:#94a3b8;
    font-size:14px;
}

.metric-value{
    font-size:32px;
    font-weight:700;
    color:white;
}

.section-header{
    font-size:24px;
    font-weight:700;
    margin-top:25px;
    margin-bottom:15px;
    color:#c4b5fd;
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# FILE PATH
# =========================================================
BASE_DIR = Path(__file__).parent
DATA_PATH = BASE_DIR / "Kittchen PNL Data.xlsx"

# =========================================================
# LOAD DATA
# =========================================================
@st.cache_data
def load_data():

    df = pd.read_excel(DATA_PATH, header=1)

    # Derived columns
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

    return df

df = load_data()

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.title("🍳 Kitchen Filters")

cities = st.sidebar.multiselect(
    "Select City",
    df["CITY"].unique(),
    default=df["CITY"].unique()
)

months = st.sidebar.multiselect(
    "Select Month",
    df["MONTH"].unique(),
    default=df["MONTH"].unique()
)

filtered_df = df[
    (df["CITY"].isin(cities)) &
    (df["MONTH"].isin(months))
]

# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div style='text-align:center;padding:10px;'>
    <h1 style='color:white;'>🍳 Kitchen P&L Intelligence Dashboard</h1>
    <p style='color:#94a3b8;'>
        Cloud Kitchen Analytics Suite
    </p>
</div>
""", unsafe_allow_html=True)

# =========================================================
# KPI SECTION
# =========================================================
total_revenue = filtered_df["NET REVENUE"].sum()
total_ebitda = filtered_df["KITCHEN EBITDA"].sum()
avg_gm = filtered_df["GM%"].mean()
avg_ebitda = filtered_df["EBITDA%"].mean()

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>TOTAL REVENUE</div>
        <div class='metric-value'>
            ₹ {total_revenue/10000000:.2f} Cr
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>TOTAL EBITDA</div>
        <div class='metric-value'>
            ₹ {total_ebitda/10000000:.2f} Cr
        </div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>AVG GM%</div>
        <div class='metric-value'>
            {avg_gm:.2f}%
        </div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-title'>AVG EBITDA%</div>
        <div class='metric-value'>
            {avg_ebitda:.2f}%
        </div>
    </div>
    """, unsafe_allow_html=True)

# =========================================================
# REVENUE TREND
# =========================================================
st.markdown(
    "<div class='section-header'>📈 Revenue Trend</div>",
    unsafe_allow_html=True
)

monthly = filtered_df.groupby("MONTH").agg({
    "NET REVENUE":"sum",
    "KITCHEN EBITDA":"sum"
}).reset_index()

fig = go.Figure()

fig.add_trace(
    go.Bar(
        x=monthly["MONTH"],
        y=monthly["NET REVENUE"],
        name="Revenue",
    )
)

fig.add_trace(
    go.Scatter(
        x=monthly["MONTH"],
        y=monthly["KITCHEN EBITDA"],
        mode="lines+markers",
        name="EBITDA"
    )
)

# =========================================================
# FIXED UPDATE_LAYOUT
# =========================================================
fig.update_layout(
    title="Revenue & EBITDA Trend",
    xaxis_title="Month",
    yaxis_title="Amount",
    template="plotly_dark",
    height=500,
    margin=dict(
        l=40,
        r=40,
        t=60,
        b=40
    ),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    )
)

st.plotly_chart(fig, use_container_width=True)

# =========================================================
# CITY PERFORMANCE
# =========================================================
st.markdown(
    "<div class='section-header'>🏙️ City Performance</div>",
    unsafe_allow_html=True
)

city_data = filtered_df.groupby("CITY").agg({
    "NET REVENUE":"sum",
    "KITCHEN EBITDA":"sum"
}).reset_index()

fig2 = px.bar(
    city_data,
    x="CITY",
    y="NET REVENUE",
    color="CITY",
    title="Revenue by City",
    template="plotly_dark"
)

fig2.update_layout(height=500)

st.plotly_chart(fig2, use_container_width=True)

# =========================================================
# STORE SCATTER
# =========================================================
st.markdown(
    "<div class='section-header'>🔮 Store Intelligence</div>",
    unsafe_allow_html=True
)

stores = filtered_df.groupby(
    ["STORE","CITY"]
).agg({
    "NET REVENUE":"mean",
    "EBITDA%":"mean"
}).reset_index()

fig3 = px.scatter(
    stores,
    x="NET REVENUE",
    y="EBITDA%",
    size="NET REVENUE",
    color="CITY",
    hover_name="STORE",
    size_max=40,
    template="plotly_dark"
)

# =========================================================
# FIXED SECTION
# =========================================================
fig3.update_layout(
    title="Revenue vs EBITDA%",
    height=600
)

st.plotly_chart(fig3, use_container_width=True)

# =========================================================
# DATA TABLE
# =========================================================
st.markdown(
    "<div class='section-header'>📋 Detailed Data</div>",
    unsafe_allow_html=True
)

st.dataframe(
    filtered_df,
    use_container_width=True,
    height=500
)

# =========================================================
# DOWNLOAD BUTTON
# =========================================================
csv = filtered_df.to_csv(index=False).encode("utf-8")

st.download_button(
    "⬇ Download CSV",
    csv,
    "kitchen_pnl.csv",
    "text/csv"
)

# =========================================================
# INSIGHTS
# =========================================================
st.markdown(
    "<div class='section-header'>💡 Insights</div>",
    unsafe_allow_html=True
)

best_city = (
    filtered_df.groupby("CITY")["EBITDA%"]
    .mean()
    .idxmax()
)

worst_city = (
    filtered_df.groupby("CITY")["EBITDA%"]
    .mean()
    .idxmin()
)

st.success(
    f"🏆 Best Performing City: {best_city}"
)

st.warning(
    f"⚠️ Lowest Performing City: {worst_city}"
)

st.info(
    f"📊 Total Stores Analysed: {filtered_df['STORE'].nunique()}"
)
