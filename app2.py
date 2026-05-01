import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
from pathlib import Path

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Salesforce Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp { background-color: #0f1117; color: #e8eaed; }

    section[data-testid="stSidebar"] {
        background-color: #1a1d27;
        border-right: 1px solid #2e3150;
    }

    .metric-card {
        background: linear-gradient(135deg, #1e2235 0%, #252a3d 100%);
        border: 1px solid #2e3150;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        margin-bottom: 12px;
    }
    .metric-card .label { font-size: 12px; color: #8b92a8; text-transform: uppercase; letter-spacing: 1px; }
    .metric-card .value { font-size: 26px; font-weight: 700; color: #e8eaed; margin-top: 4px; }
    .metric-card .sub   { font-size: 12px; color: #5b9bd5; margin-top: 4px; }

    h1, h2, h3 { color: #e8eaed !important; }

    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

    .section-header {
        font-size: 18px; font-weight: 600; color: #e8eaed;
        border-left: 4px solid #5b9bd5;
        padding-left: 12px; margin: 28px 0 16px 0;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CHART THEME
# ─────────────────────────────────────────────
COLORS = {
    "blue":   "#5b9bd5",
    "orange": "#f4a261",
    "green":  "#52b788",
    "red":    "#e76f51",
    "purple": "#9b72cf",
    "bg":     "#1e2235",
    "grid":   "#2e3150",
    "text":   "#e8eaed",
    "sub":    "#8b92a8",
}
CHART_THEME = dict(
    paper_bgcolor="#1e2235",
    plot_bgcolor="#1e2235",
    font=dict(family="Inter", color=COLORS["text"]),
    xaxis=dict(gridcolor=COLORS["grid"], zerolinecolor=COLORS["grid"]),
    yaxis=dict(gridcolor=COLORS["grid"], zerolinecolor=COLORS["grid"]),
    margin=dict(l=40, r=20, t=50, b=40),
)
TAB10 = px.colors.qualitative.Plotly

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data
def load_data(filepath: str = "salesforce.xlsx") -> dict:
    path = Path(filepath)
    if not path.exists():
        st.error(f"❌ File not found: `{filepath}`. Place `salesforce.xlsx` in the same folder as `app.py`.")
        st.stop()

    sheets = ["accounts", "contacts", "opportunities", "leads", "cases", "tasks", "products", "pricebook_entries"]
    data = {}
    xls = pd.ExcelFile(path)
    for sheet in sheets:
        if sheet in xls.sheet_names:
            data[sheet] = xls.parse(sheet)
        else:
            st.warning(f"Sheet `{sheet}` not found in the Excel file. Skipping.")
            data[sheet] = pd.DataFrame()
    return data

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.image("salesforce.png", width=280)
    st.markdown("##   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  Salesforce Analytics")
    st.markdown("---")

    data = load_data("salesforce.xlsx")

    accounts      = data["accounts"].copy()
    opportunities = data["opportunities"].copy()
    contacts      = data["contacts"].copy()

    # Coerce numeric columns
    for col in ["AnnualRevenue", "NumberOfEmployees"]:
        if col in accounts.columns:
            accounts[col] = pd.to_numeric(accounts[col], errors="coerce")

    for col in ["Amount"]:
        if col in opportunities.columns:
            opportunities[col] = pd.to_numeric(opportunities[col], errors="coerce")

    if "CloseDate" in opportunities.columns:
        opportunities["CloseDate"] = pd.to_datetime(opportunities["CloseDate"], errors="coerce")

    page = st.radio(
        "Navigate to",
        [
            "🏠 Overview",
            "📅 Monthly Revenue",
            "💎 Top Deals",
            "🔵 Revenue vs Employees",
            "🏆 Account Rating Analysis",
            "👥 Contacts per Account",
            "🔄 Open Pipeline by Account",
            "🔀 Revenue vs Pipeline Quadrant",
        ],
    )

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def fmt_usd(val):
    if val >= 1e9:
        return f"${val/1e9:.1f}B"
    elif val >= 1e6:
        return f"${val/1e6:.1f}M"
    elif val >= 1e3:
        return f"${val/1e3:.1f}K"
    return f"${val:,.0f}"

def section(title):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)

def metric_card(label, value, sub=""):
    st.markdown(
        f'<div class="metric-card"><div class="label">{label}</div>'
        f'<div class="value">{value}</div>'
        f'<div class="sub">{sub}</div></div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────
# OVERVIEW PAGE
# ─────────────────────────────────────────────
if page == "🏠 Overview":
    st.title("Salesforce Analytics Dashboard")
    st.markdown("High-level KPIs from your Salesforce data export.")

    opp = opportunities.dropna(subset=["Amount"])
    acc = accounts.dropna(subset=["AnnualRevenue"])

    won = opp[opp["StageName"] == "Closed Won"] if "StageName" in opp.columns else pd.DataFrame()
    open_pipe = opp[~opp["StageName"].isin(["Closed Won", "Closed Lost"])] if "StageName" in opp.columns else opp

    cols = st.columns(4)
    with cols[0]:
        metric_card("Total Accounts", f"{len(accounts):,}", "in accounts sheet")
    with cols[1]:
        metric_card("Total Opportunities", f"{len(opp):,}", "with amount")
    with cols[2]:
        metric_card("Closed Won Revenue", fmt_usd(won["Amount"].sum()) if not won.empty else "—", "all time")
    with cols[3]:
        metric_card("Open Pipeline", fmt_usd(open_pipe["Amount"].sum()) if not open_pipe.empty else "—", "excl. closed")

    #st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
            section("Pipeline Stage Distribution")
            if "StageName" in opp.columns:
                stage_counts = opp.groupby("StageName")["Amount"].agg(["sum", "count"]).reset_index()
                stage_counts.columns = ["Stage", "Total Value", "Count"]
                stage_counts = stage_counts.sort_values("Total Value", ascending=True)

                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=stage_counts["Total Value"],
                    y=stage_counts["Stage"],
                    name="Total Value",
                    orientation="h",
                    marker_color=COLORS["blue"],
                    hovertemplate="%{y}<br>$%{x:,.0f}<extra></extra>",
                    xaxis="x1",
                ))
                fig.add_trace(go.Scatter(
                    x=stage_counts["Count"],
                    y=stage_counts["Stage"],
                    name="Deal Count",
                    mode="lines+markers",
                    line=dict(color=COLORS["orange"], width=2),
                    marker=dict(size=8),
                    xaxis="x2",
                    hovertemplate="%{y}<br>%{x} deals<extra></extra>",
                ))
                fig.update_layout(
                    paper_bgcolor=CHART_THEME["paper_bgcolor"],
                    plot_bgcolor=CHART_THEME["plot_bgcolor"],
                    font=CHART_THEME["font"],
                    margin=CHART_THEME["margin"],
                    yaxis=dict(gridcolor=COLORS["grid"]),
                    xaxis=dict(title="Total Value (USD)", tickprefix="$", tickformat=",.0f",
                               gridcolor=COLORS["grid"]),
                    xaxis2=dict(title="Deal Count", overlaying="x", side="top",
                                gridcolor="rgba(0,0,0,0)"),
                    legend=dict(x=0.8, y=0),
                    height=340,
                )
                st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("""
            <div style="font-size:13px; color:#8b92a8; line-height:1.8; margin-top:8px;">
            💡 <b style="color:#e8eaed;">Closed Won</b> dominates at $3.6M across 18 deals, strong close rate.<br>
            📊 Entire open pipeline ($2.1M) is only 58% of already-won revenue.<br>
            ⚠️ <b style="color:#e8eaed;">Needs Analysis</b> holds $675K in a single deal, high value, high risk.<br>
            🔻 <b style="color:#e8eaed;">Qualification</b> nearly empty at $15K, top of funnel needs attention.<br>
            🏁 <b style="color:#e8eaed;">Negotiation & Proposal</b> stages close in value, near-term closures likely.
            </div>
            """, unsafe_allow_html=True)

    with col2:
        section("Revenue by Industry (Top 8)")
        acc_ind = acc.dropna(subset=["Industry"]) if "Industry" in acc.columns else pd.DataFrame()
        if not acc_ind.empty:
            top_ind = acc_ind.groupby("Industry")["AnnualRevenue"].sum().nlargest(8).reset_index()
            fig = px.pie(
                top_ind, names="Industry", values="AnnualRevenue",
                color_discrete_sequence=TAB10,
                hole=0.45,
            )
            fig.update_layout(**CHART_THEME, height=340, legend=dict(font=dict(size=10)))
            fig.update_traces(textinfo="percent+label", hovertemplate="%{label}<br>$%{value:,.0f}")
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("""
            <div style="font-size:13px; color:#8b92a8; line-height:1.8; margin-top:8px;">
            💡 <b style="color:#e8eaed;">Energy</b> dominates at $5.6B, nearly 6x the next industry.<br>
            📊 <b style="color:#e8eaed;">Construction & Transportation</b> are tied at $950M.<br>
            📉 Steep drop-off after top 3, bottom 4 combined don't match Transportation.<br>
            🌱 <b style="color:#e8eaed;">Biotech & Consulting</b> are small but worth watching for growth.
            </div>
            """, unsafe_allow_html=True)
            
    col3, col4 = st.columns(2)
    
    with col3:
        section("Pipeline Stage Distribution - Table")
        if "StageName" in opp.columns:
            out_pipe = opp.groupby("StageName")["Amount"].agg(["sum", "count"]).reset_index()
            out_pipe.columns = ["Stage", "Total Value", "Deal Count"]
            out_pipe = out_pipe.sort_values("Total Value", ascending=False)
            out_pipe["Total Value"] = out_pipe["Total Value"].apply(lambda x: f"${x:,.0f}")
            st.dataframe(out_pipe, use_container_width=True, hide_index=True)

    with col4:
        section("Revenue by Industry - Table")
        if not acc_ind.empty:
            out_ind = acc_ind.groupby("Industry")["AnnualRevenue"].sum().sort_values(ascending=False).reset_index()
            out_ind.columns = ["Industry", "Total Revenue"]
            out_ind["Total Revenue"] = out_ind["Total Revenue"].apply(lambda x: f"${x:,.0f}")
            st.dataframe(out_ind, use_container_width=True, hide_index=True)



# ─────────────────────────────────────────────
# MONTHLY REVENUE
# ─────────────────────────────────────────────
elif page == "📅 Monthly Revenue":
    st.title("Monthly Closed Won Revenue")

    opp = opportunities.dropna(subset=["Amount", "CloseDate"])
    if "StageName" not in opp.columns:
        st.warning("StageName column not found.")
        st.stop()

    df3 = opp[opp["StageName"] == "Closed Won"].copy()
    if df3.empty:
        st.info("No Closed Won opportunities found.")
        st.stop()

    # ✅ Keep as datetime (CRITICAL)
    df3["YearMonth"] = df3["CloseDate"].dt.to_period("M").dt.to_timestamp()
    df3_grouped = df3.groupby("YearMonth")["Amount"].sum().reset_index()

    # KPI row
    col1, col2, col3 = st.columns(3)
    with col1:
        metric_card("Total Closed Won", fmt_usd(df3_grouped["Amount"].sum()), "all months")
    with col2:
        peak_row = df3_grouped.loc[df3_grouped["Amount"].idxmax()]
        metric_card("Peak Month", fmt_usd(peak_row["Amount"]), peak_row["YearMonth"].strftime("%b %Y"))
    with col3:
        metric_card("Months Tracked", str(len(df3_grouped)), "in dataset")

    st.markdown("---")

    fig = go.Figure()

    # Gradient fill area
    fig.add_trace(go.Scatter(
        x=df3_grouped["YearMonth"],
        y=df3_grouped["Amount"],
        mode="none",
        fill="tozeroy",
        fillcolor="rgba(91,155,213,0.08)",
        showlegend=False,
        hoverinfo="skip",
    ))

    # Main line
    fig.add_trace(go.Scatter(
        x=df3_grouped["YearMonth"],
        y=df3_grouped["Amount"],
        mode="lines+markers+text",
        line=dict(color=COLORS["blue"], width=3, shape="spline", smoothing=0.6),
        marker=dict(
            size=10,
            color=df3_grouped["Amount"],
            colorscale=[[0, COLORS["blue"]], [1, COLORS["green"]]],
            line=dict(color="#1e2235", width=2),
        ),
        text=df3_grouped["Amount"].apply(fmt_usd),
        textposition="top center",
        textfont=dict(size=11, color=COLORS["text"]),
        hovertemplate="%{x|%b %Y}<br><b>$%{y:,.0f}</b><extra></extra>",
        name="Closed Won Revenue",
    ))

    # ✅ Peak month vertical line (robust method)
    peak_month = peak_row["YearMonth"].to_pydatetime()

    fig.add_shape(
        type="line",
        x0=peak_month,
        x1=peak_month,
        y0=0,
        y1=1,
        xref="x",
        yref="paper",
        line=dict(
            color=COLORS["green"],
            dash="dash",
            width=1.5,
        ),
    )

    fig.add_annotation(
        x=peak_month,
        y=1,
        yref="paper",
        text="Peak",
        showarrow=False,
        font=dict(color=COLORS["green"], size=11),
    )

    # ✅ Axis formatting (THIS is what fixes your labels)
    fig.update_xaxes(
        tickformat="%b %Y",
        dtick="M1"
    )

    fig.update_layout(
        paper_bgcolor=CHART_THEME["paper_bgcolor"],
        plot_bgcolor=CHART_THEME["plot_bgcolor"],
        font=CHART_THEME["font"],
        margin=dict(l=40, r=20, t=50, b=40),
        xaxis=dict(
            title="Month",
            gridcolor=COLORS["grid"],
            tickangle=-30,  # ✅ keep rotation only
        ),
        yaxis=dict(
            title="Revenue (USD)",
            tickprefix="$",
            tickformat=",.0f",
            gridcolor=COLORS["grid"],
        ),
        showlegend=False,
        height=480,
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown("""
    <div style="font-size:13px; color:#8b92a8; line-height:1.8; margin-top:8px; margin-bottom:16px;">
    💡 Rapid ramp-up from Dec 2025 through Mar 2026 — pipeline was just getting started.<br>
    🏆 <b style="color:#e8eaed;">March 2026</b> was the strongest month at $1.475M in closed revenue.<br>
    📉 <b style="color:#e8eaed;">April's drop to $345K</b> likely reflects incomplete data — month may not be finished.<br>
    📊 $3.645M total closed won revenue, almost entirely concentrated in <b style="color:#e8eaed;">Q1 2026</b>.
    </div>
    """, unsafe_allow_html=True)

    # ───────── Table ─────────
    section("Monthly Breakdown")

    df_table = df3_grouped.copy()
    df_table["YearMonth"] = df_table["YearMonth"].dt.strftime("%b %Y")  # ✅ FIX
    df_table["Amount"] = df_table["Amount"].apply(lambda x: f"${x:,.0f}")

    st.dataframe(
        df_table.rename(columns={"YearMonth": "Month", "Amount": "Revenue"}),
        use_container_width=True,
        hide_index=True
    )

# ─────────────────────────────────────────────
# TOP DEALS
# ─────────────────────────────────────────────
elif page == "💎 Top Deals":
    st.title("Top 10 Opportunities by Deal Size")

    opp = opportunities.dropna(subset=["Amount"])
    if opp.empty:
        st.warning("No opportunities with Amount found.")
        st.stop()

    # Resolve account name
    if "AccountName" not in opp.columns:
        if "Account.Name" in opp.columns:
            opp = opp.rename(columns={"Account.Name": "AccountName"})
        elif "AccountId" in opp.columns and not accounts.empty and "Name" in accounts.columns and "Id" in accounts.columns:
            opp = opp.merge(
                accounts[["Id", "Name"]].rename(columns={"Id": "AccountId", "Name": "AccountName"}),
                on="AccountId", how="left"
            )
        else:
            opp["AccountName"] = "Unknown"

    # ✅ Step 1: Get top 10 deals (by total amount per deal)
    deal_name_col = "Name" if "Name" in opp.columns else None
    if deal_name_col is None:
        opp["DealName"] = opp.index.astype(str)
        deal_name_col = "DealName"

    deal_totals = opp.groupby(deal_name_col)["Amount"].sum().nlargest(10)
    top_deals = opp[opp[deal_name_col].isin(deal_totals.index)].copy()

    # ✅ Step 2: Aggregate by Deal + Stage
    df5 = (
        top_deals
        .groupby([deal_name_col, "StageName"], as_index=False)["Amount"]
        .sum()
    )

    # ✅ Step 3: Compute row-based % (per deal)
    df5["TotalPerDeal"] = df5.groupby(deal_name_col)["Amount"].transform("sum")
    df5["Pct"] = df5["Amount"] / df5["TotalPerDeal"]

    # ✅ Stage color mapping
    def stage_color(stage):
        if stage == "Closed Won":
            return COLORS["green"]
        elif isinstance(stage, str) and "Negotiation" in stage:
            return COLORS["orange"]
        else:
            return COLORS["red"]

    stage_order = ["Closed Won", "Negotiation/Review", "Other"]
    stage_map = {
        "Closed Won": "Closed Won",
        "Negotiation/Review": "Negotiation/Review",
    }

    df5["StageGroup"] = df5["StageName"].apply(
        lambda s: "Closed Won" if s == "Closed Won"
        else ("Negotiation/Review" if "Negotiation" in str(s) else "Other")
    )

    # Pivot for stacked bars
    pivot_df = df5.pivot_table(
        index=deal_name_col,
        columns="StageGroup",
        values="Amount",
        fill_value=0
    )

    # Keep order of deals (largest on top)
    pivot_df = pivot_df.loc[deal_totals.index]

    fig = go.Figure()

    # ✅ Build stacked bars
    for stage in ["Closed Won", "Negotiation/Review", "Other"]:
        if stage in pivot_df.columns:
            amounts = pivot_df[stage]

            # Get matching % values
            pct_map = df5[df5["StageGroup"] == stage].set_index(deal_name_col)["Pct"]

            fig.add_trace(go.Bar(
                x=amounts,
                y=pivot_df.index,
                orientation="h",
                name=stage,
                marker_color=stage_color(stage),
                text=[
                    f"{fmt_usd(val)} ({pct_map.get(deal, 0):.0%})" if val > 0 else ""
                    for deal, val in zip(pivot_df.index, amounts)
                ],
                textposition="inside",
                insidetextanchor="middle",
                textfont=dict(color="white", size=11),
                hovertemplate="%{y}<br>$%{x:,.0f}<extra></extra>",
            ))

    fig.update_layout(
        **CHART_THEME,
        barmode="stack",  # ✅ TRUE stacked chart
        xaxis_title="Amount (USD)",
        xaxis_tickprefix="$",
        xaxis_tickformat=",.0f",
        yaxis_autorange="reversed",
        height=500,
    )

    st.plotly_chart(fig, use_container_width=True)

    # ───────── Table ─────────
    section("Top 10 Deal Details")

    cols_show = [c for c in ["Name", "AccountName", "StageName", "Amount", "CloseDate"] if c in top_deals.columns]
    out = top_deals[cols_show].copy()

    # ✅ Ensure datetime first
    if "CloseDate" in out.columns:
        out["CloseDate"] = pd.to_datetime(out["CloseDate"], errors="coerce")
        out["CloseDate"] = out["CloseDate"].dt.strftime("%b %d %Y")

    # Format amount
    if "Amount" in out.columns:
        out["Amount"] = out["Amount"].apply(lambda x: f"${x:,.0f}")

    st.dataframe(out, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# REVENUE VS EMPLOYEES
# ─────────────────────────────────────────────
elif page == "🔵 Revenue vs Employees":
    st.title("Annual Revenue vs Number of Employees")

    df6 = accounts.dropna(subset=["AnnualRevenue", "NumberOfEmployees"])
    if "Industry" in df6.columns:
        df6 = df6.dropna(subset=["Industry"])

    if df6.empty:
        st.warning("Need AnnualRevenue + NumberOfEmployees columns in accounts sheet.")
        st.stop()

    df6 = df6.copy()
    df6["RevenuePerEmployee"] = (df6["AnnualRevenue"] / df6["NumberOfEmployees"]).round(0)
    name_col = "Name" if "Name" in df6.columns else df6.columns[0]
    industry_col = "Industry" if "Industry" in df6.columns else None

    tab1, tab2 = st.tabs(["All Accounts", "Excluding Outlier"])

    for tab, exclude_outlier in [(tab1, False), (tab2, True)]:
        with tab:
            plot_df = df6.copy()
            outlier_name = None
            if exclude_outlier and not df6.empty:
                outlier_name = df6.loc[df6["AnnualRevenue"].idxmax(), name_col]
                plot_df = df6[df6[name_col] != outlier_name].copy()

            fig = px.scatter(
                plot_df,
                x="NumberOfEmployees",
                y="AnnualRevenue",
                color=industry_col if industry_col else None,
                size="RevenuePerEmployee",
                size_max=40,
                hover_name=name_col,
                hover_data={"AnnualRevenue": ":$,.0f", "NumberOfEmployees": ":,",
                            "RevenuePerEmployee": ":$,.0f"},
                color_discrete_sequence=TAB10,
                labels={"AnnualRevenue": "Annual Revenue (USD)", "NumberOfEmployees": "Employees"},
            )
            fig.update_layout(
                **CHART_THEME,
                yaxis_tickprefix="$", yaxis_tickformat=",.0f",
                height=520,
                title=f"Annual Revenue vs Employees{' (excl. ' + outlier_name + ')' if outlier_name else ''}",
            )
            st.plotly_chart(fig, use_container_width=True)
            
    st.markdown("""
    <div style="font-size:13px; color:#8b92a8; line-height:1.6; margin-top:8px; margin-bottom:16px;">
    💡 Two views are shown because <b style="color:#e8eaed;">United Oil & Gas Corp.</b> (~$5.6B revenue) is a major outlier that compresses all other data points.<br>
    📊 In the full chart, large enterprises dominate the scale, making mid-sized companies appear tightly clustered.<br>
    🔍 Removing the outlier reveals clearer comparisons across industries and highlights operational differences.<br>
    📈 Several smaller firms generate higher <b style="color:#e8eaed;">revenue per employee</b>, suggesting stronger efficiency despite lower total revenue.
    </div>
    """, unsafe_allow_html=True)

    section("Revenue per Employee Table")
    show_cols = [c for c in [name_col, industry_col, "AnnualRevenue", "NumberOfEmployees", "RevenuePerEmployee"]
                 if c and c in df6.columns]
    out = df6[show_cols].sort_values("RevenuePerEmployee", ascending=False).copy()
    for col in ["AnnualRevenue", "RevenuePerEmployee"]:
        if col in out.columns:
            out[col] = out[col].apply(lambda x: f"${x:,.0f}")
    st.dataframe(out, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# ACCOUNT RATING ANALYSIS
# ─────────────────────────────────────────────


elif page == "🏆 Account Rating Analysis":
    st.title("Account Rating Analysis")
    st.markdown(
        "<span style='color:#8b92a8; font-size:14px;'>Revenue concentration, account quality segmentation, "
        "and industry exposure by CRM rating tier</span>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    df9 = accounts.dropna(subset=["AnnualRevenue"])
    if "Rating" not in df9.columns:
        st.warning("Rating column not found in accounts sheet.")
        st.stop()
    df9 = df9.dropna(subset=["Rating"])

    rating_order  = ["Hot", "Warm", "Cold"]
    rating_colors = {"Hot": COLORS["red"], "Warm": COLORS["orange"], "Cold": COLORS["blue"]}

    df9_grouped = df9.groupby("Rating")["AnnualRevenue"].agg(
        Avg_Revenue="mean", Total_Revenue="sum", Account_Count="count"
    ).reindex([r for r in rating_order if r in df9["Rating"].unique()]).reset_index()

    total_rev  = df9["AnnualRevenue"].sum()
    hot_rev    = df9_grouped.loc[df9_grouped["Rating"] == "Hot",  "Total_Revenue"].sum()
    hot_count  = df9_grouped.loc[df9_grouped["Rating"] == "Hot",  "Account_Count"].sum()
    top_account = df9.loc[df9["AnnualRevenue"].idxmax(), "Name"] if "Name" in df9.columns else "—"
    hot_share  = (hot_rev / total_rev * 100) if total_rev > 0 else 0
    avg_rev_all = df9["AnnualRevenue"].mean()

    # ── KPI Row ──────────────────────────────────────────────
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Portfolio Revenue", fmt_usd(total_rev), f"{len(df9)} rated accounts")
    with k2:
        metric_card("Hot Tier Revenue Share", f"{hot_share:.0f}%", f"{int(hot_count)} accounts driving outsized value")
    with k3:
        metric_card("Portfolio Avg Revenue", fmt_usd(avg_rev_all), "per account")
    with k4:
        metric_card("Largest Account", top_account, fmt_usd(df9["AnnualRevenue"].max()))

    st.markdown("---")

    # ── Dodged bar: Revenue by Rating × Industry ──────────────
    col1, col2 = st.columns([3, 2])

    with col1:
        section("Revenue by Rating & Industry (Dodged)")
        if "Industry" in df9.columns:
            df9_dodge = df9.groupby(["Rating", "Industry"])["AnnualRevenue"].sum().reset_index()
            df9_dodge = df9_dodge[df9_dodge["Rating"].isin(rating_order)]

            fig_dodge = go.Figure()
            industries_present = df9_dodge["Industry"].unique().tolist()
            palette = TAB10

            for i, ind in enumerate(industries_present):
                sub = df9_dodge[df9_dodge["Industry"] == ind]
                # Reindex so every rating slot exists (fills gaps with 0)
                sub = sub.set_index("Rating").reindex(rating_order).reset_index().fillna(0)
                sub["Industry"] = ind
                fig_dodge.add_trace(go.Bar(
                    name=ind,
                    x=sub["Rating"],
                    y=sub["AnnualRevenue"],
                    marker_color=palette[i % len(palette)],
                    text=sub["AnnualRevenue"].apply(lambda v: fmt_usd(v) if v > 0 else ""),
                    textposition="outside",
                    hovertemplate=f"{ind} · %{{x}}<br>${{y:,.0f}}<extra></extra>",
                ))

            fig_dodge.update_layout(
                paper_bgcolor=CHART_THEME["paper_bgcolor"],
                plot_bgcolor=CHART_THEME["plot_bgcolor"],
                font=CHART_THEME["font"],
                margin=CHART_THEME["margin"],
                barmode="group",
                xaxis=dict(title="Rating Tier", gridcolor=COLORS["grid"]),
                yaxis=dict(title="Total Revenue (USD)", tickprefix="$",
                           tickformat=",.0f", gridcolor=COLORS["grid"]),
                legend=dict(title="Industry", font=dict(size=10),
                            bgcolor="rgba(0,0,0,0)"),
                height=420,
            )
            st.plotly_chart(fig_dodge, use_container_width=True)
        else:
            st.info("Industry column not found — cannot render dodged chart.")

    with col2:
        section("Revenue Concentration by Tier")

        fig_pie = go.Figure(go.Pie(
            labels=df9_grouped["Rating"],
            values=df9_grouped["Total_Revenue"],
            hole=0.55,
            marker=dict(colors=[rating_colors.get(r, COLORS["blue"]) for r in df9_grouped["Rating"]],
                        line=dict(color="#1e2235", width=3)),
            textinfo="label+percent",
            hovertemplate="%{label}<br>$%{value:,.0f}<br>%{percent}<extra></extra>",
        ))
        fig_pie.add_annotation(
            text=f"<b>{fmt_usd(total_rev)}</b><br><span style='font-size:10px'>Total</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color=COLORS["text"]),
        )
        fig_pie.update_layout(
            paper_bgcolor=CHART_THEME["paper_bgcolor"],
            font=CHART_THEME["font"],
            margin=dict(l=20, r=20, t=30, b=20),
            legend=dict(font=dict(size=11), bgcolor="rgba(0,0,0,0)"),
            height=420,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # ── Avg Revenue per Rating — horizontal bars ──────────────
    st.markdown("---")
    section("Average Revenue per Account by Rating Tier")

    df9_grouped["Revenue_per_Account"] = df9_grouped["Total_Revenue"] / df9_grouped["Account_Count"]
    max_val = df9_grouped["Avg_Revenue"].max()

    fig_avg = go.Figure()
    fig_avg.add_trace(go.Bar(
        x=df9_grouped["Avg_Revenue"],
        y=df9_grouped["Rating"],
        orientation="h",
        marker=dict(
            color=df9_grouped["Avg_Revenue"],
            colorscale=[[0, COLORS["blue"]], [0.5, COLORS["orange"]], [1, COLORS["red"]]],
            line=dict(color="#1e2235", width=1),
        ),
        text=df9_grouped["Avg_Revenue"].apply(lambda x: f"${x:,.0f}"),
        textposition="outside",
        hovertemplate="%{y} tier<br>Avg $%{x:,.0f}<extra></extra>",
    ))
    fig_avg.add_vline(
        x=avg_rev_all, line_dash="dash", line_color=COLORS["sub"],
        annotation_text=f"Portfolio avg {fmt_usd(avg_rev_all)}",
        annotation_position="top right",
        annotation_font_color=COLORS["sub"],
        annotation_font_size=11,
    )
    fig_avg.update_layout(
        paper_bgcolor=CHART_THEME["paper_bgcolor"],
        plot_bgcolor=CHART_THEME["plot_bgcolor"],
        font=CHART_THEME["font"],
        margin=CHART_THEME["margin"],
        xaxis=dict(title="Avg Annual Revenue (USD)", tickprefix="$",
                   tickformat=",.0f", gridcolor=COLORS["grid"]),
        yaxis=dict(gridcolor=COLORS["grid"], autorange="reversed"),
        height=240,
        showlegend=False,
    )
    st.plotly_chart(fig_avg, use_container_width=True)

    # ── Executive Insights ────────────────────────────────────
    st.markdown("""
    <div style="font-size:13px; color:#8b92a8; line-height:2; margin-top:4px; margin-bottom:20px;">
    💡 <b style="color:#e8eaed;">Hot accounts</b> generate disproportionate revenue — high retention risk if any single account churns.<br>
    📊 <b style="color:#e8eaed;">Cold accounts</b> show competitive avg revenue — re-engagement or upsell campaigns warranted.<br>
    ⚠️ Rating distribution is uniform (2 per tier) — CRM scoring criteria may need recalibration to improve segmentation signal.<br>
    🏆 <b style="color:#e8eaed;">Energy sector</b> dominates the Hot tier — portfolio is exposed to a single-industry concentration risk.
    </div>
    """, unsafe_allow_html=True)

    # ── Summary Table ─────────────────────────────────────────
    section("Summary Table")
    out = df9_grouped.copy()
    out["Avg_Revenue"]   = out["Avg_Revenue"].apply(lambda x: f"${x:,.0f}")
    out["Total_Revenue"] = out["Total_Revenue"].apply(lambda x: f"${x:,.0f}")
    st.dataframe(
        out.rename(columns={
            "Avg_Revenue": "Avg Revenue", "Total_Revenue": "Total Revenue", "Account_Count": "# Accounts"
        })[["Rating", "# Accounts", "Total Revenue", "Avg Revenue"]],
        use_container_width=True, hide_index=True,
    )

    # ── Detailed Breakdown ────────────────────────────────────
    if "Industry" in df9.columns and "Name" in df9.columns:
        section("Detailed Account Breakdown")
        det = df9[["Name", "Rating", "Industry", "AnnualRevenue"]].sort_values(
            "AnnualRevenue", ascending=False
        ).copy()
        det["AnnualRevenue"] = det["AnnualRevenue"].apply(lambda x: f"${x:,.0f}")
        det["Revenue Share"] = (
            df9.sort_values("AnnualRevenue", ascending=False)["AnnualRevenue"]
            / total_rev * 100
        ).apply(lambda x: f"{x:.1f}%").values
        st.dataframe(det.rename(columns={"AnnualRevenue": "Annual Revenue"}),
                     use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────
# CONTACTS PER ACCOUNT
# ─────────────────────────────────────────────

elif page == "👥 Contacts per Account":
    st.title("Relationship Coverage Analysis")
    st.markdown(
        "<span style='color:#8b92a8; font-size:14px;'>Account penetration depth, CRM contact risk scoring, "
        "and rating-tier coverage across the commercial portfolio</span>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    df11 = data["contacts"].copy()
    if df11.empty:
        st.warning("Contacts sheet is empty.")
        st.stop()

    # ── Resolve account name ──────────────────────────────────
    if "AccountName" not in df11.columns:
        if "Account.Name" in df11.columns:
            df11 = df11.rename(columns={"Account.Name": "AccountName"})
        elif "AccountId" in df11.columns and not accounts.empty and "Id" in accounts.columns:
            df11 = df11.merge(
                accounts[["Id", "Name"]].rename(columns={"Id": "AccountId", "Name": "AccountName"}),
                on="AccountId", how="left",
            )
        else:
            df11["AccountName"] = "Unknown"

    df11 = df11.dropna(subset=["AccountName"])

    # ── Build full name ───────────────────────────────────────
    if "FullName" not in df11.columns:
        if "FirstName" in df11.columns and "LastName" in df11.columns:
            df11["FullName"] = df11["FirstName"].fillna("") + " " + df11["LastName"].fillna("")
        elif "Name" in df11.columns:
            df11["FullName"] = df11["Name"]
        else:
            df11["FullName"] = df11.index.astype(str)

    # ── Base aggregation ──────────────────────────────────────
    df11_grouped = df11.groupby("AccountName").agg(
        Contact_Count=("FullName", "count"),
        Contacts=("FullName", lambda x: ", ".join(x.astype(str))),
    ).sort_values("Contact_Count", ascending=False).reset_index()

    def coverage_tier(n):
        if n == 1: return "Single Contact (High Risk)"
        if n == 2: return "Dual Contact"
        return "Multi-Contact (Healthy)"

    df11_grouped["Coverage_Tier"] = df11_grouped["Contact_Count"].apply(coverage_tier)

    # ── Merge Rating from accounts ────────────────────────────
    if "Rating" in accounts.columns and "Name" in accounts.columns:
        rating_map = accounts.set_index("Name")["Rating"].to_dict()
        df11_grouped["Rating"] = df11_grouped["AccountName"].map(rating_map).fillna("Unrated")
    else:
        df11_grouped["Rating"] = "Unrated"

    # ── KPI Metrics ───────────────────────────────────────────
    total_accounts  = len(df11_grouped)
    single_contact  = (df11_grouped["Contact_Count"] == 1).sum()
    healthy_accounts = (df11_grouped["Contact_Count"] >= 3).sum()
    avg_contacts    = df11_grouped["Contact_Count"].mean()
    at_risk_pct     = single_contact / total_accounts * 100 if total_accounts > 0 else 0
    total_contacts  = df11_grouped["Contact_Count"].sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("Total Accounts Covered", str(total_accounts), f"{total_contacts} total contacts")
    with k2:
        metric_card("Avg Contacts / Account", f"{avg_contacts:.1f}", "portfolio penetration depth")
    with k3:
        metric_card("Single-Contact Risk", f"{at_risk_pct:.0f}%", f"{single_contact} accounts exposed")
    with k4:
        metric_card("Healthy Coverage (3+)", f"{healthy_accounts}", f"{healthy_accounts/total_accounts*100:.0f}% of accounts")

    st.markdown("---")

    # ── Dodged bar: Contact Count × Rating ───────────────────
    col1, col2 = st.columns([3, 2])

    with col1:
        section("Contact Coverage by Account & Rating Tier")

        rating_order_dodge  = ["Hot", "Warm", "Cold", "Unrated"]
        rating_colors_dodge = {
            "Hot":     COLORS["red"],
            "Warm":    COLORS["orange"],
            "Cold":    COLORS["blue"],
            "Unrated": COLORS["sub"],
        }

        fig_dodge = go.Figure()
        ratings_present = [r for r in rating_order_dodge if r in df11_grouped["Rating"].unique()]

        for rating in ratings_present:
            sub = df11_grouped[df11_grouped["Rating"] == rating].sort_values(
                "Contact_Count", ascending=False
            )
            fig_dodge.add_trace(go.Bar(
                name=rating,
                x=sub["AccountName"],
                y=sub["Contact_Count"],
                marker_color=rating_colors_dodge[rating],
                text=sub["Contact_Count"],
                textposition="outside",
                hovertemplate=f"{rating}<br>%{{x}}<br>%{{y}} contacts<extra></extra>",
            ))

        fig_dodge.update_layout(
            paper_bgcolor=CHART_THEME["paper_bgcolor"],
            plot_bgcolor=CHART_THEME["plot_bgcolor"],
            font=CHART_THEME["font"],
            margin=CHART_THEME["margin"],
            barmode="group",
            xaxis=dict(title="Account", gridcolor=COLORS["grid"], tickangle=-35),
            yaxis=dict(title="Number of Contacts", gridcolor=COLORS["grid"], dtick=1),
            legend=dict(title="CRM Rating", font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
            height=420,
        )
        st.plotly_chart(fig_dodge, use_container_width=True)

    with col2:
        section("Coverage Risk Distribution")

        tier_summary = df11_grouped.groupby("Coverage_Tier")["Contact_Count"].count().reset_index()
        tier_summary.columns = ["Tier", "Count"]
        tier_color_map = {
            "Single Contact (High Risk)": COLORS["red"],
            "Dual Contact":               COLORS["orange"],
            "Multi-Contact (Healthy)":    COLORS["green"],
        }

        fig_donut = go.Figure(go.Pie(
            labels=tier_summary["Tier"],
            values=tier_summary["Count"],
            hole=0.55,
            marker=dict(
                colors=[tier_color_map.get(t, COLORS["blue"]) for t in tier_summary["Tier"]],
                line=dict(color="#1e2235", width=3),
            ),
            textinfo="label+percent",
            hovertemplate="%{label}<br>%{value} accounts<br>%{percent}<extra></extra>",
        ))
        fig_donut.add_annotation(
            text=f"<b>{total_accounts}</b><br><span style='font-size:10px'>Accounts</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=14, color=COLORS["text"]),
        )
        fig_donut.update_layout(
            paper_bgcolor=CHART_THEME["paper_bgcolor"],
            font=CHART_THEME["font"],
            margin=dict(l=20, r=20, t=30, b=20),
            legend=dict(font=dict(size=10), bgcolor="rgba(0,0,0,0)"),
            height=420,
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    # ── Horizontal risk bar ───────────────────────────────────
    st.markdown("---")
    section("Account Penetration Depth — Ranked by Contact Count")

    df11_sorted = df11_grouped.sort_values("Contact_Count", ascending=True)
    bar_colors  = [
        COLORS["red"]    if n == 1 else
        COLORS["orange"] if n == 2 else
        COLORS["green"]
        for n in df11_sorted["Contact_Count"]
    ]

    fig_h = go.Figure(go.Bar(
        x=df11_sorted["Contact_Count"],
        y=df11_sorted["AccountName"],
        orientation="h",
        marker_color=bar_colors,
        text=df11_sorted["Contact_Count"],
        textposition="outside",
        hovertemplate="%{y}<br>%{x} contacts<extra></extra>",
        showlegend=False,
    ))
    # Benchmark line at avg
    fig_h.add_vline(
        x=avg_contacts, line_dash="dash", line_color=COLORS["sub"],
        annotation_text=f"Avg {avg_contacts:.1f}",
        annotation_position="top right",
        annotation_font_color=COLORS["sub"],
        annotation_font_size=11,
    )
    fig_h.update_layout(
        paper_bgcolor=CHART_THEME["paper_bgcolor"],
        plot_bgcolor=CHART_THEME["plot_bgcolor"],
        font=CHART_THEME["font"],
        margin=CHART_THEME["margin"],
        xaxis=dict(title="Contact Count", gridcolor=COLORS["grid"], dtick=1),
        yaxis=dict(gridcolor=COLORS["grid"]),
        height=max(380, len(df11_grouped) * 28 + 100),
    )
    st.plotly_chart(fig_h, use_container_width=True)

    # ── Executive Insights ────────────────────────────────────
    st.markdown(f"""
    <div style="font-size:13px; color:#8b92a8; line-height:2; margin-top:4px; margin-bottom:20px;">
    ⚠️ <b style="color:#e8eaed;">{at_risk_pct:.0f}% of accounts</b> have a single contact — 
    any personnel change at those accounts creates immediate churn risk.<br>
    🏆 <b style="color:#e8eaed;">United Oil & Gas Corp.</b> leads with 4 contacts — 
    the only account with institutional-grade relationship coverage.<br>
    📊 Portfolio average of <b style="color:#e8eaed;">{avg_contacts:.1f} contacts per account</b> 
    is below the 3-contact benchmark for enterprise sales health.<br>
    🎯 Priority action: expand contact depth at <b style="color:#e8eaed;">single-contact accounts</b> 
    before next renewal cycle to reduce key-person dependency risk.
    </div>
    """, unsafe_allow_html=True)

    # ── Data Table ────────────────────────────────────────────
    section("Full Account Contact Register")
    table_out = df11_grouped[["AccountName", "Rating", "Coverage_Tier", "Contact_Count", "Contacts"]].copy()
    st.dataframe(
        table_out.rename(columns={
            "AccountName":   "Account",
            "Rating":        "CRM Rating",
            "Coverage_Tier": "Coverage Tier",
            "Contact_Count": "# Contacts",
        }),
        use_container_width=True, hide_index=True,
    )

# ─────────────────────────────────────────────
# OPEN PIPELINE BY ACCOUNT
# ─────────────────────────────────────────────

elif page == "🔄 Open Pipeline by Account":

    st.title("Open Pipeline Value by Account")

    st.markdown(
        "<span style='color:#8b92a8; font-size:14px;'>"
        "Pipeline distribution by account and opportunity stage"
        "</span>",
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # ── Prepare Opportunity Data ─────────────────────────────
    opp = opportunities.dropna(subset=["Amount"]).copy()

    # Ensure numeric
    opp["Amount"] = pd.to_numeric(
        opp["Amount"],
        errors="coerce"
    )

    # Clean stage names
    if "StageName" not in opp.columns:
        st.warning("StageName column not found.")
        st.stop()

    opp["StageName"] = (
        opp["StageName"]
        .astype(str)
        .str.strip()
    )

    # Keep only open opportunities
    df10 = opp[
        ~opp["StageName"].isin(["Closed Won", "Closed Lost"])
    ].copy()

    if df10.empty:
        st.info("No open opportunities found.")
        st.stop()

    # ── Resolve Account Names ────────────────────────────────
    if "AccountName" not in df10.columns:

        if "Account.Name" in df10.columns:

            df10 = df10.rename(
                columns={"Account.Name": "AccountName"}
            )

        elif (
            "AccountId" in df10.columns
            and not accounts.empty
            and "Id" in accounts.columns
        ):

            df10 = df10.merge(
                accounts[["Id", "Name"]].rename(
                    columns={
                        "Id": "AccountId",
                        "Name": "AccountName"
                    }
                ),
                on="AccountId",
                how="left",
            )

        else:
            df10["AccountName"] = "Unknown"

    # Remove blanks
    df10 = df10.dropna(subset=["AccountName"])

    # ── Build Pivot Table ────────────────────────────────────
    df10_pivot = (
        df10.groupby(
            ["AccountName", "StageName"]
        )["Amount"]
        .sum()
        .reset_index()
        .pivot_table(
            index="AccountName",
            columns="StageName",
            values="Amount",
            fill_value=0,
        )
    )

    # Total for sorting
    df10_pivot["_Total"] = df10_pivot.sum(axis=1)

    # Sort descending
    df10_pivot = df10_pivot.sort_values(
        "_Total",
        ascending=True
    )

    # Save totals
    total_values = df10_pivot["_Total"]

    # Remove helper column
    df10_pivot = df10_pivot.drop(columns="_Total")

    # ── KPI Metrics ──────────────────────────────────────────
    total_pipeline = df10["Amount"].sum()

    total_open_deals = len(df10)

    top_account = (
        total_values.idxmax()
        if not total_values.empty
        else "—"
    )

    top_account_value = (
        total_values.max()
        if not total_values.empty
        else 0
    )

    avg_pipeline = (
        total_pipeline / len(df10_pivot)
        if len(df10_pivot) > 0
        else 0
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        metric_card(
            "Total Open Pipeline",
            f"${total_pipeline:,.0f}",
            f"{total_open_deals} open deals"
        )

    with k2:
        metric_card(
            "Accounts with Pipeline",
            f"{len(df10_pivot)}",
            "active opportunity accounts"
        )

    with k3:
        metric_card(
            "Largest Pipeline Account",
            top_account,
            f"${top_account_value:,.0f}"
        )

    with k4:
        metric_card(
            "Avg Pipeline / Account",
            f"${avg_pipeline:,.0f}",
            "portfolio average"
        )

    st.markdown("---")

    # ── Horizontal Stacked Bar Chart ─────────────────────────
    section("Pipeline Distribution by Account & Stage")

    fig = go.Figure()

    for i, stage in enumerate(df10_pivot.columns):

        fig.add_trace(
            go.Bar(
                y=df10_pivot.index,
                x=df10_pivot[stage],

                name=stage,

                orientation="h",

                marker=dict(
                    color=TAB10[i % len(TAB10)],
                    line=dict(
                        color="#1e2235",
                        width=1,
                    ),
                ),

                hovertemplate=(
                    f"<b>{stage}</b>"
                    "<br>Account: %{y}"
                    "<br>Pipeline: $%{x:,.0f}"
                    "<extra></extra>"
                ),
            )
        )

    fig.update_layout(

        # Background
        paper_bgcolor=CHART_THEME["paper_bgcolor"],
        plot_bgcolor=CHART_THEME["plot_bgcolor"],

        # Font
        font=CHART_THEME["font"],

        # Margins
        margin=CHART_THEME["margin"],

        # Stack bars
        barmode="stack",

        # X Axis
        xaxis=dict(
            title="Open Pipeline Value (USD)",
            tickprefix="$",
            tickformat=",.0f",
            gridcolor=COLORS["grid"],
            zeroline=False,
        ),

        # Y Axis
        yaxis=dict(
            title="Account",
            gridcolor="rgba(0,0,0,0)",
            automargin=True,
        ),

        # Legend
        legend=dict(
            title="Opportunity Stage",
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
            bgcolor="rgba(0,0,0,0)",
        ),

        # Hover
        hovermode="y unified",

        # Dynamic Height
        height=max(
            500,
            len(df10_pivot) * 38
        ),
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    # ── Executive Insights ───────────────────────────────────
    st.markdown(f"""
    <div style="font-size:13px; color:#8b92a8; line-height:1.7; margin-top:4px; margin-bottom:20px;">
    📊 <b style="color:#e8eaed;">{top_account}</b>
    represents the largest concentration of open pipeline
    at <b style="color:#e8eaed;">${top_account_value:,.0f}</b>.<br>
    💰 Portfolio-wide open pipeline totals
    <b style="color:#e8eaed;">${total_pipeline:,.0f}</b>
    across <b style="color:#e8eaed;">{total_open_deals}</b> active opportunities.<br>
    ⚠️ Heavy pipeline concentration in a small number of accounts
    may increase forecast volatility and renewal dependency risk.<br>
    🎯 Opportunity stage diversification improves forecast resilience
    and reduces late-stage pipeline bottlenecks.
    </div>
    """, unsafe_allow_html=True)

    # ── Summary Table ────────────────────────────────────────
    section("Pipeline Summary by Account")

    summary = (
        df10.groupby("AccountName")["Amount"]
        .agg(
            Total_Open_Value="sum",
            Open_Deals="count",
        )
        .sort_values(
            "Total_Open_Value",
            ascending=False
        )
        .reset_index()
    )

    out = summary.copy()

    out["Total_Open_Value"] = out[
        "Total_Open_Value"
    ].apply(lambda x: f"${x:,.0f}")

    st.dataframe(
        out.rename(
            columns={
                "AccountName": "Account",
                "Total_Open_Value": "Total Open Value",
                "Open_Deals": "# Open Deals",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

# ─────────────────────────────────────────────
# REVENUE VS PIPELINE QUADRANT
# ─────────────────────────────────────────────

elif page == "🔀 Revenue vs Pipeline Quadrant":

    st.title("Annual Revenue vs Open Pipeline — Quadrant Analysis")

    st.markdown(
        "<span style='color:#8b92a8; font-size:14px;'>"
        "Relationship between current account revenue and active pipeline generation"
        "</span>",
        unsafe_allow_html=True,
    )

    st.markdown("---")

    # ── Prepare Data ─────────────────────────────────────────
    acc = accounts.dropna(subset=["AnnualRevenue"]).copy()
    opp = opportunities.dropna(subset=["Amount"]).copy()

    acc["AnnualRevenue"] = pd.to_numeric(
        acc["AnnualRevenue"],
        errors="coerce"
    )

    opp["Amount"] = pd.to_numeric(
        opp["Amount"],
        errors="coerce"
    )

    if "StageName" in opp.columns:

        opp["StageName"] = (
            opp["StageName"]
            .astype(str)
            .str.strip()
        )

        open_opp = opp[
            ~opp["StageName"].isin(
                ["Closed Won", "Closed Lost"]
            )
        ].copy()

    else:
        open_opp = opp.copy()

    # ── Resolve Account Names ────────────────────────────────
    if "AccountName" not in open_opp.columns:

        if "Account.Name" in open_opp.columns:

            open_opp = open_opp.rename(
                columns={"Account.Name": "AccountName"}
            )

        elif (
            "AccountId" in open_opp.columns
            and not acc.empty
            and "Id" in acc.columns
        ):

            open_opp = open_opp.merge(
                acc[["Id", "Name"]].rename(
                    columns={
                        "Id": "AccountId",
                        "Name": "AccountName"
                    }
                ),
                on="AccountId",
                how="left"
            )

        else:
            open_opp["AccountName"] = "Unknown"

    # ── Aggregate Pipeline ───────────────────────────────────
    pipe_summary = (
        open_opp.groupby("AccountName")["Amount"]
        .sum()
        .reset_index()
    )

    pipe_summary.columns = [
        "Name",
        "OpenPipeline"
    ]

    name_col = (
        "Name"
        if "Name" in acc.columns
        else acc.columns[0]
    )

    df12 = (
        acc[[name_col, "AnnualRevenue"]]
        .merge(
            pipe_summary,
            left_on=name_col,
            right_on="Name",
            how="left"
        )
    )

    df12 = df12.dropna(
        subset=[
            "AnnualRevenue",
            "OpenPipeline"
        ]
    )

    if "Name_x" in df12.columns:

        df12 = (
            df12.rename(
                columns={"Name_x": "Name"}
            )
            .drop(
                columns=["Name_y"],
                errors="ignore"
            )
        )

    df12.columns = [
        c.replace("_x", "")
        for c in df12.columns
    ]

    if df12.empty:
        st.warning(
            "Not enough joined data for quadrant analysis."
        )
        st.stop()

    # ── Median Thresholds ────────────────────────────────────
    med_rev = df12["AnnualRevenue"].median()
    med_pipe = df12["OpenPipeline"].median()

    # ── Quadrant Classification ──────────────────────────────
    def quadrant(row):

        if (
            row["AnnualRevenue"] >= med_rev
            and row["OpenPipeline"] >= med_pipe
        ):
            return "High Rev / High Pipeline"

        elif (
            row["AnnualRevenue"] < med_rev
            and row["OpenPipeline"] >= med_pipe
        ):
            return "Low Rev / High Pipeline"

        elif (
            row["AnnualRevenue"] >= med_rev
            and row["OpenPipeline"] < med_pipe
        ):
            return "High Rev / Low Pipeline"

        return "Low Rev / Low Pipeline"

    df12["Quadrant"] = df12.apply(
        quadrant,
        axis=1
    )

    # ── Colors ───────────────────────────────────────────────
    quad_colors = {
        "High Rev / High Pipeline": COLORS["green"],
        "Low Rev / High Pipeline":  COLORS["orange"],
        "High Rev / Low Pipeline":  COLORS["blue"],
        "Low Rev / Low Pipeline":   COLORS["red"],
    }

    # ── Outlier Detection ────────────────────────────────────
    outlier_name = df12.loc[
        df12["AnnualRevenue"].idxmax(),
        "Name"
    ]

    df12_main = df12[
        df12["Name"] != outlier_name
    ].copy()

    df12_outlier = df12[
        df12["Name"] == outlier_name
    ].copy()

    # ── KPI Metrics ──────────────────────────────────────────
    total_pipeline = df12["OpenPipeline"].sum()

    total_arr = df12["AnnualRevenue"].sum()

    high_high_count = len(
        df12[
            df12["Quadrant"] ==
            "High Rev / High Pipeline"
        ]
    )

    low_low_count = len(
        df12[
            df12["Quadrant"] ==
            "Low Rev / Low Pipeline"
        ]
    )

    pipeline_ratio = (
        total_pipeline / total_arr * 100
        if total_arr > 0
        else 0
    )

    k1, k2, k3, k4 = st.columns(4)

    with k1:
        metric_card(
            "Total Open Pipeline",
            f"${total_pipeline:,.0f}",
            f"{len(df12)} analyzed accounts"
        )

    with k2:
        metric_card(
            "Pipeline / Revenue Ratio",
            f"{pipeline_ratio:.2f}%",
            "open pipeline vs ARR"
        )

    with k3:
        metric_card(
            "High Rev / High Pipeline",
            str(high_high_count),
            "strategic growth accounts"
        )

    with k4:
        metric_card(
            "Low Rev / Low Pipeline",
            str(low_low_count),
            "underdeveloped accounts"
        )

    st.markdown("---")

    # ── Tabs ─────────────────────────────────────────────────
    tab1, tab2 = st.tabs([
        "Main View (excl. outlier)",
        "Outlier"
    ])

    # ── MAIN VIEW ────────────────────────────────────────────
    with tab1:

        fig = px.scatter(

            df12_main,

            x="AnnualRevenue",
            y="OpenPipeline",

            color="Quadrant",

            color_discrete_map=quad_colors,

            hover_name="Name",

            hover_data={
                "AnnualRevenue": ":$,.0f",
                "OpenPipeline": ":$,.0f",
            },

            labels={
                "AnnualRevenue": "Annual Revenue (USD)",
                "OpenPipeline": "Open Pipeline (USD)",
            },
        )

        # Median lines
        fig.add_vline(
            x=df12_main["AnnualRevenue"].median(),
            line_dash="dash",
            line_color="gray",
        )

        fig.add_hline(
            y=df12_main["OpenPipeline"].median(),
            line_dash="dash",
            line_color="gray",
        )

        # Layout
        fig.update_layout(

            paper_bgcolor=CHART_THEME["paper_bgcolor"],
            plot_bgcolor=CHART_THEME["plot_bgcolor"],
            font=CHART_THEME["font"],
            margin=CHART_THEME["margin"],

            xaxis_tickprefix="$",
            xaxis_tickformat=",.0f",

            yaxis_tickprefix="$",
            yaxis_tickformat=",.0f",

            height=520,

            legend=dict(
                title="Quadrant",
                bgcolor="rgba(0,0,0,0)",
            ),
        )

        fig.update_traces(
            marker=dict(
                size=12,
                opacity=0.85,
                line=dict(
                    width=1,
                    color="#1e2235"
                )
            )
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )

    # ── OUTLIER VIEW ─────────────────────────────────────────
    with tab2:

        fig2 = px.scatter(

            df12_outlier,

            x="AnnualRevenue",
            y="OpenPipeline",

            hover_name="Name",

            hover_data={
                "AnnualRevenue": ":$,.0f",
                "OpenPipeline": ":$,.0f",
            },
        )

        fig2.update_traces(
            marker=dict(
                size=18,
                color=COLORS["blue"],
                opacity=0.9,
                line=dict(
                    width=1,
                    color="#1e2235"
                )
            )
        )

        fig2.update_layout(

            paper_bgcolor=CHART_THEME["paper_bgcolor"],
            plot_bgcolor=CHART_THEME["plot_bgcolor"],
            font=CHART_THEME["font"],
            margin=CHART_THEME["margin"],

            xaxis_tickprefix="$",
            xaxis_tickformat=",.0f",

            yaxis_tickprefix="$",
            yaxis_tickformat=",.0f",

            title=f"Outlier Account: {outlier_name}",

            height=400,
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

    # ── Executive Insights ───────────────────────────────────
    st.markdown(f"""
    <div style="font-size:13px; color:#8b92a8; line-height:1.7; margin-top:4px; margin-bottom:20px;">
    🟢 <b style="color:#e8eaed;">High Revenue / High Pipeline accounts</b>
    represent the healthiest commercial segment — these accounts are simultaneously
    generating strong recurring revenue and maintaining future growth momentum.<br>
    🔴 <b style="color:#e8eaed;">Low Revenue / Low Pipeline accounts</b>
    show limited current value and weak forward-looking opportunity generation,
    indicating either low strategic priority or underdeveloped account penetration.<br>
    📊 The portfolio currently maintains an open pipeline equal to
    <b style="color:#e8eaed;">{pipeline_ratio:.2f}%</b> of total annual revenue,
    providing a benchmark for future sales coverage and forecasting stability.<br>
    ⚠️ The large outlier account significantly distorts portfolio averages,
    which suggests elevated concentration risk and dependence on a single enterprise customer.<br>
    🎯 Accounts positioned in the
    <b style="color:#e8eaed;">High Revenue / Low Pipeline</b> quadrant
    should become immediate pipeline generation priorities to reduce future revenue stagnation risk.<br>
    🚀 Accounts in the
    <b style="color:#e8eaed;">Low Revenue / High Pipeline</b> quadrant
    represent emerging expansion opportunities and may evolve into future strategic accounts if conversion rates remain healthy.
    </div>
    """, unsafe_allow_html=True)

    # ── Summary Table ────────────────────────────────────────
    section("Quadrant Summary Table")

    quad_summary = (
        df12.groupby("Quadrant")
        .agg(
            Accounts=("Name", "count"),
            Avg_Revenue=("AnnualRevenue", "mean"),
            Avg_Pipeline=("OpenPipeline", "mean"),
        )
        .reset_index()
    )

    quad_summary["Avg_Revenue"] = (
        quad_summary["Avg_Revenue"]
        .apply(lambda x: f"${x:,.0f}")
    )

    quad_summary["Avg_Pipeline"] = (
        quad_summary["Avg_Pipeline"]
        .apply(lambda x: f"${x:,.0f}")
    )

    st.dataframe(

        quad_summary.rename(
            columns={
                "Avg_Revenue": "Avg Annual Revenue",
                "Avg_Pipeline": "Avg Open Pipeline",
            }
        ),

        use_container_width=True,
        hide_index=True,
    )