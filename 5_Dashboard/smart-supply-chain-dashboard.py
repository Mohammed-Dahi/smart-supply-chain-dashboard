# ============================================================
# Smart Supply Chain Dashboard — DataCo Star Schema
# Sheets: Fact_Orders, Dim_Shipping, Dim_Date,
#         Dim_Department, Dim_Customer, Dim_Product
#
# pip install streamlit plotly pandas openpyxl
# streamlit run supply_chain_dashboard.py
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Smart Supply Chain Dashboard",
    page_icon="📦", layout="wide",
    initial_sidebar_state="expanded"
)

BG, CARD, RED = "#000000", "#1A1A1A", "#FF1F1F"
WHITE, GREY, BORDER = "#FFFFFF", "#AAAAAA", "#2A2A2A"
GREEN, GOLD = "#00C853", "#FFD700"

st.markdown(f"""
<style>
  html,body,[class*="css"]{{background:{BG};color:{WHITE};font-family:'Segoe UI',sans-serif;}}
  .stApp{{background:{BG};}}
  section[data-testid="stSidebar"]{{background:{CARD};border-right:1px solid {BORDER};}}
  section[data-testid="stSidebar"] *{{color:{WHITE}!important;}}
  .kpi-card{{background:{CARD};border:1px solid {BORDER};border-left:4px solid {RED};
             border-radius:10px;padding:18px 22px;margin-bottom:10px;}}
  .kpi-label{{font-size:11px;color:{GREY};text-transform:uppercase;letter-spacing:1px;}}
  .kpi-value{{font-size:28px;font-weight:700;color:{WHITE};margin:4px 0;}}
  .kpi-delta-pos{{font-size:12px;color:{GREEN};}}
  .kpi-delta-neg{{font-size:12px;color:{RED};}}
  .section-header{{font-size:17px;font-weight:600;color:{WHITE};
                   border-left:4px solid {RED};padding-left:10px;margin:22px 0 12px 0;}}
  div[data-baseweb="tab-list"]{{background:{CARD};border-radius:8px;padding:4px;border:1px solid {BORDER};}}
  div[data-baseweb="tab"]{{color:{GREY}!important;border-radius:6px;}}
  div[aria-selected="true"]{{background:{RED}!important;color:{WHITE}!important;}}
  ::-webkit-scrollbar{{width:5px;}}
  ::-webkit-scrollbar-track{{background:{BG};}}
  ::-webkit-scrollbar-thumb{{background:{RED};border-radius:3px;}}
  hr{{border-color:{BORDER};}}
</style>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def theme(fig, title="", h=380):
    fig.update_layout(
        title=dict(text=title, font=dict(color=WHITE, size=14), x=0.01),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor=CARD,
        height=h, font=dict(color=WHITE, size=11),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=WHITE),
                    bordercolor=BORDER, borderwidth=1),
        margin=dict(l=10, r=10, t=42, b=10),
        xaxis=dict(gridcolor=BORDER, zerolinecolor=BORDER,
                   tickfont=dict(color=GREY), title_font=dict(color=GREY)),
        yaxis=dict(gridcolor=BORDER, zerolinecolor=BORDER,
                   tickfont=dict(color=GREY), title_font=dict(color=GREY)),
    )
    return fig

def kpi(label, value, delta=None, pos_good=True):
    dhtml = ""
    if delta is not None:
        css  = "kpi-delta-pos" if (delta >= 0) == pos_good else "kpi-delta-neg"
        sign = "▲" if delta >= 0 else "▼"
        dhtml = f'<div class="{css}">{sign} {abs(delta):.1f}%</div>'
    return f"""<div class="kpi-card">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value">{value}</div>{dhtml}
    </div>"""

def sh(text):
    st.markdown(f'<div class="section-header">{text}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# DATA LOADING — exact column names from schema
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading DataCo Star Schema…")
def load(file):
    xl = pd.ExcelFile(file, engine="openpyxl")

    # ── Fact_Orders ──────────────────────────────
    fo = xl.parse("Fact_Orders")
    fo.columns = fo.columns.str.strip()
    # Exact cols visible in Image 1:
    # Order Id, Order Date, Customer Id, Product Card Id, Department Id,
    # Order Item Quantity, Order Status, Order Item Product Price, Sales,
    # Order Item Discount, Order Item Total, Order Profit Per Order,
    # Benefit per order, Days for shipping (real), Days for shipment (scheduled),
    # Late_delivery_risk
    fo["Order Date"] = pd.to_datetime(fo["Order Date"], errors="coerce")
    fo["Year"]    = fo["Order Date"].dt.year
    fo["Month"]   = fo["Order Date"].dt.to_period("M").astype(str)
    fo["Quarter"] = fo["Order Date"].dt.to_period("Q").astype(str)

    # Normalise key numeric cols
    num_map = {
        "Sales":                            "Sales",
        "Order Profit Per Order":           "Profit",
        "Benefit per order":                "Benefit",
        "Order Item Quantity":              "Quantity",
        "Days for shipping (real)":         "Actual Ship Days",
        "Days for shipment (scheduled)":    "Scheduled Ship Days",
        "Late_delivery_risk":               "Late Delivery Risk",
        "Order Item Discount":              "Discount",
        "Order Item Total":                 "Order Total",
    }
    for orig, new in num_map.items():
        if orig in fo.columns:
            fo[new] = pd.to_numeric(fo[orig], errors="coerce").fillna(0)
        elif new not in fo.columns:
            fo[new] = 0

    fo["Is Fraud"] = fo["Order Status"].str.strip().str.upper().str.contains("FRAUD").astype(int)
    fo.dropna(subset=["Order Date"], inplace=True)

    # ── Dim_Shipping ─────────────────────────────
    # Cols: Order Id, Shipping Mode, Delivery Status, Shipping Date
    ds = xl.parse("Dim_Shipping")
    ds.columns = ds.columns.str.strip()
    ds["Shipping Date"] = pd.to_datetime(ds["Shipping Date"], errors="coerce")

    # ── Dim_Date ─────────────────────────────────
    # Cols: YearMonth, Month, Year
    dd = xl.parse("Dim_Date")
    dd.columns = dd.columns.str.strip()

    # ── Dim_Department ───────────────────────────
    # Cols: Department Id, Department Name
    dept = xl.parse("Dim_Department")
    dept.columns = dept.columns.str.strip()

    # ── Dim_Customer ─────────────────────────────
    # Cols: Customer Id, Customer Name, Customer Segment,
    #       Customer City, Customer State, Customer Country
    dc = xl.parse("Dim_Customer")
    dc.columns = dc.columns.str.strip()

    # ── Dim_Product ──────────────────────────────
    # Cols: Product Card Id, Product Name, Category Name,
    #       Category Id, Product Category Id, Product Price
    dp = xl.parse("Dim_Product")
    dp.columns = dp.columns.str.strip()

    # ── JOIN all dims onto Fact ───────────────────
    df = fo.copy()

    # Shipping
    if "Order Id" in ds.columns and "Shipping Mode" in ds.columns:
        df = df.merge(
            ds[["Order Id","Shipping Mode","Delivery Status","Shipping Date"]],
            on="Order Id", how="left"
        )
    else:
        df["Shipping Mode"] = "Unknown"
        df["Delivery Status"] = "Unknown"

    # Customer
    if "Customer Id" in dc.columns:
        df = df.merge(dc, on="Customer Id", how="left")

    # Product
    if "Product Card Id" in dp.columns:
        df = df.merge(dp, on="Product Card Id", how="left")

    # Department
    if "Department Id" in dept.columns:
        df = df.merge(dept, on="Department Id", how="left")

    # Fill missing str cols
    for c in ["Shipping Mode","Delivery Status","Customer Segment",
              "Customer Name","Customer City","Customer Country",
              "Customer State","Product Name","Category Name","Department Name"]:
        if c not in df.columns:
            df[c] = "Unknown"
        else:
            df[c].fillna("Unknown", inplace=True)

    return df, ds, dd, dept, dc, dp

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"<h2 style='color:{RED};margin-bottom:2px'>📦 Supply Chain</h2>", unsafe_allow_html=True)
    st.markdown(f"<p style='color:{GREY};font-size:11px'>DataCo Smart Analytics</p>", unsafe_allow_html=True)
    st.divider()

    uploaded = st.file_uploader("Upload DataCo_StarSchema.xlsx", type=["xlsx"],
                                 help="Must contain all 6 sheets")
    if not uploaded:
        st.info("⬆️ Upload your Excel file to begin")
        st.stop()

    df, ds, dd, dept, dc, dp = load(uploaded)

    st.markdown("### 🎛️ Filters")

    mn, mx = df["Order Date"].min().date(), df["Order Date"].max().date()
    dr = st.date_input("📅 Date Range", (mn, mx), min_value=mn, max_value=mx)

    segs  = ["All"] + sorted(df["Customer Segment"].dropna().unique())
    seg   = st.selectbox("👥 Segment",       segs)

    modes = ["All"] + sorted(df["Shipping Mode"].dropna().unique())
    mode  = st.selectbox("🚚 Shipping Mode", modes)

    stats = ["All"] + sorted(df["Order Status"].dropna().unique())
    stat  = st.selectbox("📋 Order Status",  stats)

    depts = ["All"] + sorted(df["Department Name"].dropna().unique())
    dept_f= st.selectbox("🏢 Department",    depts)

    cats  = ["All"] + sorted(df["Category Name"].dropna().unique())
    cat   = st.selectbox("📦 Category",      cats)

    st.divider()
    st.markdown(f"<p style='color:{GREY};font-size:10px'>DataCo Smart Supply Chain</p>",
                unsafe_allow_html=True)

# ─────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────
d = df.copy()
if len(dr) == 2:
    d = d[(d["Order Date"] >= pd.Timestamp(dr[0])) &
          (d["Order Date"] <= pd.Timestamp(dr[1]))]
if seg    != "All": d = d[d["Customer Segment"] == seg]
if mode   != "All": d = d[d["Shipping Mode"]    == mode]
if stat   != "All": d = d[d["Order Status"]     == stat]
if dept_f != "All": d = d[d["Department Name"]  == dept_f]
if cat    != "All": d = d[d["Category Name"]    == cat]

if d.empty:
    st.warning("⚠️ No data for selected filters."); st.stop()

# ─────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────
tot_orders   = d["Order Id"].nunique()
tot_sales    = d["Sales"].sum()
tot_profit   = d["Profit"].sum()
tot_benefit  = d["Benefit"].sum()
tot_cust     = d["Customer Id"].nunique()
tot_qty      = d["Quantity"].sum()
fraud_rate   = d["Is Fraud"].mean() * 100
on_time      = (1 - d["Late Delivery Risk"].mean()) * 100
avg_actual   = d["Actual Ship Days"].mean()
avg_sched    = d["Scheduled Ship Days"].mean()
avg_profit   = tot_profit / tot_orders if tot_orders else 0
avg_sales_c  = tot_sales  / tot_cust   if tot_cust   else 0
fraud_sales  = d[d["Is Fraud"]==1]["Sales"].sum()
fraud_cnt    = int(d["Is Fraud"].sum())
top_seg      = d.groupby("Customer Segment")["Sales"].sum().idxmax() if tot_sales else "N/A"
top_prod     = d.groupby("Product Name")["Quantity"].sum().idxmax()  if tot_qty   else "N/A"

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown(f"""
<div style="background:{CARD};border-bottom:2px solid {RED};padding:14px 20px;
            border-radius:10px;margin-bottom:18px;display:flex;align-items:center;gap:12px">
  <span style="font-size:30px">📦</span>
  <div>
    <div style="font-size:21px;font-weight:700;color:{WHITE}">Smart Supply Chain Dashboard</div>
    <div style="font-size:12px;color:{GREY}">
      DataCo Analytics · {tot_orders:,} orders · {len(d):,} rows after filters
    </div>
  </div>
</div>""", unsafe_allow_html=True)

tabs = st.tabs(["📊 Fulfillment","👥 Customers","💰 Profitability","🚚 Shipping","📈 Trends"])

# ══════════════════════════════════════════════
# TAB 1 — FULFILLMENT
# ══════════════════════════════════════════════
with tabs[0]:
    sh("Order & Fulfillment Performance")
    c1,c2,c3,c4 = st.columns(4)
    c1.markdown(kpi("Total Orders",      f"{tot_orders:,}"),          unsafe_allow_html=True)
    c2.markdown(kpi("Total Sales",       f"${tot_sales/1e6:.2f}M"),   unsafe_allow_html=True)
    c3.markdown(kpi("Fraud Rate",        f"{fraud_rate:.2f}%", pos_good=False), unsafe_allow_html=True)
    c4.markdown(kpi("On-Time Delivery",  f"{on_time:.1f}%"),          unsafe_allow_html=True)
    st.markdown("---")

    # Orders over time (monthly)
    mo = d.groupby("Month").agg(Orders=("Order Id","count"), Sales=("Sales","sum")).reset_index().sort_values("Month")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=mo["Month"], y=mo["Orders"], mode="lines",
                             line=dict(color=RED,width=2.5), fill="tozeroy",
                             fillcolor="rgba(255,31,31,0.10)", name="Orders"))
    fig.add_trace(go.Scatter(x=mo["Month"], y=mo["Sales"]/1e3, mode="lines",
                             line=dict(color=GOLD,width=1.8,dash="dot"),
                             name="Sales ($K)", yaxis="y2"))
    fig.update_layout(yaxis2=dict(overlaying="y", side="right",
                                  tickfont=dict(color=GOLD), title_font=dict(color=GOLD),
                                  gridcolor="rgba(0,0,0,0)"))
    fig = theme(fig, "Orders & Sales Over Time", 320)
    st.plotly_chart(fig, use_container_width=True)

    ca, cb = st.columns(2)
    with ca:
        # Revenue by Order Status
        s_df = d.groupby("Order Status")["Sales"].sum().reset_index().sort_values("Sales")
        fig2 = px.bar(s_df, x="Sales", y="Order Status", orientation="h",
                      color_discrete_sequence=[RED])
        fig2 = theme(fig2, "Revenue by Order Status", 340)
        st.plotly_chart(fig2, use_container_width=True)

    with cb:
        # Shipping Mode Donut — from Dim_Shipping
        sm = d.groupby("Shipping Mode")["Order Id"].count().reset_index(name="Orders")
        fig3 = px.pie(sm, names="Shipping Mode", values="Orders", hole=0.55,
                      color_discrete_sequence=[RED,"#FF6B6B","#FF9999","#FFCCCC"])
        fig3.update_traces(textfont_color=WHITE)
        fig3 = theme(fig3, "Shipping Mode Distribution", 340)
        st.plotly_chart(fig3, use_container_width=True)

    # Delivery Status breakdown (from Dim_Shipping join)
    sh("Delivery Status Breakdown")
    dl = d.groupby("Delivery Status")["Order Id"].count().reset_index(name="Orders").sort_values("Orders", ascending=False)
    fig4 = px.bar(dl, x="Delivery Status", y="Orders",
                  color="Delivery Status",
                  color_discrete_sequence=[RED,"#FF6B6B","#FF9999","#FFCCCC","#FFE0E0"])
    fig4 = theme(fig4, "", 280)
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════
# TAB 2 — CUSTOMERS
# ══════════════════════════════════════════════
with tabs[1]:
    sh("Customer & Market Insights")
    c1,c2,c3 = st.columns(3)
    c1.markdown(kpi("Total Customers",      f"{tot_cust:,}"),            unsafe_allow_html=True)
    c2.markdown(kpi("Top Segment",          top_seg),                    unsafe_allow_html=True)
    c3.markdown(kpi("Avg Sales / Customer", f"${avg_sales_c:,.0f}"),     unsafe_allow_html=True)
    st.markdown("---")

    ca, cb = st.columns([1.3, 0.7])
    with ca:
        # Top 10 Customers
        tc = (d.groupby("Customer Name")["Sales"]
               .sum().reset_index().sort_values("Sales", ascending=False)
               .head(10).sort_values("Sales"))
        fig = px.bar(tc, x="Sales", y="Customer Name", orientation="h",
                     color_discrete_sequence=[RED])
        fig = theme(fig, "Top 10 Customers by Sales", 420)
        st.plotly_chart(fig, use_container_width=True)

    with cb:
        # Segment column
        sg = d.groupby("Customer Segment")["Sales"].sum().reset_index().sort_values("Sales", ascending=False)
        fig2 = px.bar(sg, x="Customer Segment", y="Sales",
                      color_discrete_sequence=[RED,"#FF6B6B","#FF9999"])
        fig2 = theme(fig2, "Sales by Segment", 420)
        st.plotly_chart(fig2, use_container_width=True)

    sh("Sales by Country / City")
    # Country bar (no lat/lon in schema)
    cc = d.groupby("Customer Country")["Sales"].sum().reset_index().sort_values("Sales", ascending=False).head(20)
    fig3 = px.bar(cc, x="Customer Country", y="Sales", color_discrete_sequence=[RED])
    fig3.update_xaxes(tickangle=-30)
    fig3 = theme(fig3, "Top Countries by Sales", 320)
    st.plotly_chart(fig3, use_container_width=True)

    # City treemap
    sh("City Sales Treemap")
    ct = d.groupby(["Customer Country","Customer State","Customer City"])["Sales"].sum().reset_index()
    fig4 = px.treemap(ct, path=["Customer Country","Customer State","Customer City"],
                      values="Sales",
                      color="Sales",
                      color_continuous_scale=["#1A1A1A","#FF6B6B",RED])
    fig4.update_layout(paper_bgcolor="rgba(0,0,0,0)", height=420,
                       margin=dict(l=0,r=0,t=10,b=0),
                       coloraxis_colorbar=dict(tickfont=dict(color=WHITE),
                                               title=dict(text="Sales",font=dict(color=WHITE))))
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════
# TAB 3 — PROFITABILITY
# ══════════════════════════════════════════════
with tabs[2]:
    sh("Product & Profitability")
    c1,c2,c3,c4 = st.columns(4)
    tp_label = (top_prod[:20]+"…") if len(top_prod)>20 else top_prod
    c1.markdown(kpi("Units Sold",         f"{int(tot_qty):,}"),        unsafe_allow_html=True)
    c2.markdown(kpi("Top Product",        tp_label),                   unsafe_allow_html=True)
    c3.markdown(kpi("Total Profit",       f"${tot_profit/1e6:.2f}M"),  unsafe_allow_html=True)
    c4.markdown(kpi("Avg Profit / Order", f"${avg_profit:.1f}"),       unsafe_allow_html=True)
    st.markdown("---")

    ca, cb = st.columns(2)
    with ca:
        # Quarterly Profit
        qp = d.groupby("Quarter")["Profit"].sum().reset_index().sort_values("Quarter")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=qp["Quarter"], y=qp["Profit"],
                                 mode="lines+markers",
                                 line=dict(color=RED,width=2.5),
                                 marker=dict(size=7,color=RED,line=dict(color=WHITE,width=1))))
        fig = theme(fig, "Quarterly Profit Trend", 320)
        st.plotly_chart(fig, use_container_width=True)

    with cb:
        # Top 5 Products by Qty
        t5 = (d.groupby("Product Name")["Quantity"]
               .sum().reset_index().sort_values("Quantity",ascending=False).head(5))
        fig2 = px.bar(t5, x="Quantity", y="Product Name", orientation="h",
                      color_discrete_sequence=[RED])
        fig2 = theme(fig2, "Top 5 Products by Units Sold", 320)
        st.plotly_chart(fig2, use_container_width=True)

    # Department Profit
    sh("Profit by Department")
    dp2 = d.groupby("Department Name")["Profit"].sum().reset_index().sort_values("Profit", ascending=False)
    fig3 = px.bar(dp2, x="Department Name", y="Profit",
                  color="Profit",
                  color_continuous_scale=["#4D0000","#FF6B6B",RED])
    fig3.update_xaxes(tickangle=-30)
    fig3 = theme(fig3, "", 300)
    st.plotly_chart(fig3, use_container_width=True)

    # Discount Sunburst → Segment > Shipping Mode > Delivery Status
    sh("Discount Patterns — Segment › Shipping › Delivery Status")
    sun = (d.groupby(["Customer Segment","Shipping Mode","Delivery Status"])
            ["Discount"].mean().reset_index())
    sun["Discount %"] = (sun["Discount"] * 100).round(2)
    sun = sun[sun["Discount %"] > 0]
    fig4 = px.sunburst(sun, path=["Customer Segment","Shipping Mode","Delivery Status"],
                       values="Discount %", color="Discount %",
                       color_continuous_scale=["#1A1A1A","#FF6B6B",RED])
    fig4.update_traces(textfont=dict(color=WHITE, size=11),
                       insidetextorientation="radial",
                       marker=dict(line=dict(color=BG, width=1.5)))
    fig4.update_layout(paper_bgcolor="rgba(0,0,0,0)", height=500,
                       margin=dict(l=0,r=0,t=10,b=0),
                       coloraxis_colorbar=dict(tickfont=dict(color=WHITE),
                                               title=dict(text="Avg Disc %",font=dict(color=WHITE))))
    st.plotly_chart(fig4, use_container_width=True)

    # Category profitability
    sh("Profit by Category")
    cat_p = (d.groupby("Category Name").agg(Profit=("Profit","sum"), Sales=("Sales","sum"))
              .reset_index().sort_values("Profit", ascending=False).head(15))
    fig5 = px.scatter(cat_p, x="Sales", y="Profit", text="Category Name",
                      color="Profit", size="Sales",
                      color_continuous_scale=["#4D0000","#FF6B6B",RED])
    fig5.update_traces(textfont=dict(color=WHITE, size=9))
    fig5 = theme(fig5, "Category: Sales vs Profit Bubble", 380)
    st.plotly_chart(fig5, use_container_width=True)


# ══════════════════════════════════════════════
# TAB 4 — SHIPPING
# ══════════════════════════════════════════════
with tabs[3]:
    sh("Shipping & Delivery Performance")
    c1,c2,c3,c4 = st.columns(4)
    c1.markdown(kpi("Avg Actual Ship Days",  f"{avg_actual:.1f}d"),                unsafe_allow_html=True)
    c2.markdown(kpi("Avg Scheduled Days",    f"{avg_sched:.1f}d"),                 unsafe_allow_html=True)
    c3.markdown(kpi("Fraud Sales $",         f"${fraud_sales:,.0f}", pos_good=False), unsafe_allow_html=True)
    c4.markdown(kpi("Fraud Orders",          f"{fraud_cnt:,}",       pos_good=False), unsafe_allow_html=True)
    st.markdown("---")

    ca, cb = st.columns([1.5, 1])
    with ca:
        # Grouped bar: Actual vs Scheduled by Shipping Mode
        sh_df = d.groupby("Shipping Mode").agg(
            Actual=("Actual Ship Days","mean"),
            Scheduled=("Scheduled Ship Days","mean")
        ).reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Scheduled", x=sh_df["Shipping Mode"],
                             y=sh_df["Scheduled"], marker_color=WHITE, opacity=0.55))
        fig.add_trace(go.Bar(name="Actual",    x=sh_df["Shipping Mode"],
                             y=sh_df["Actual"],    marker_color=RED))
        fig.update_layout(barmode="group")
        fig = theme(fig, "Shipping Efficiency: Actual vs Scheduled Days", 380)
        fig.add_hline(y=sh_df["Scheduled"].mean(), line_dash="dot",
                      line_color=GOLD, line_width=1.5,
                      annotation_text="Avg SLA", annotation_font_color=GOLD)
        st.plotly_chart(fig, use_container_width=True)

    with cb:
        # Gauge
        fig2 = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=on_time,
            delta={"reference":85,"suffix":"%",
                   "increasing":{"color":GREEN},"decreasing":{"color":RED}},
            number={"suffix":"%","font":{"color":WHITE,"size":34}},
            gauge={"axis":{"range":[0,100],"tickcolor":GREY,"tickfont":{"color":GREY}},
                   "bar":{"color":RED},
                   "bgcolor":CARD, "bordercolor":BORDER,
                   "steps":[{"range":[0,60],"color":"#2D0000"},
                             {"range":[60,85],"color":"#4D1010"},
                             {"range":[85,100],"color":"#0D2D0D"}],
                   "threshold":{"line":{"color":GOLD,"width":3},
                                "thickness":0.8,"value":85}},
            title={"text":"On-Time Delivery Rate","font":{"color":WHITE,"size":13}}
        ))
        fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)", height=380,
                           font=dict(color=WHITE), margin=dict(l=30,r=30,t=40,b=10))
        st.plotly_chart(fig2, use_container_width=True)

    # Delivery Status by Shipping Mode heatmap
    sh("Delivery Status Heatmap by Shipping Mode")
    hm = d.groupby(["Shipping Mode","Delivery Status"])["Order Id"].count().reset_index(name="Count")
    hm_piv = hm.pivot(index="Shipping Mode", columns="Delivery Status", values="Count").fillna(0)
    fig3 = go.Figure(go.Heatmap(
        z=hm_piv.values,
        x=hm_piv.columns.tolist(),
        y=hm_piv.index.tolist(),
        colorscale=[[0,"#1A1A1A"],[0.5,"#FF6B6B"],[1,RED]],
        text=hm_piv.values.astype(int),
        texttemplate="%{text}",
        textfont={"color":WHITE,"size":11}
    ))
    fig3 = theme(fig3, "", 300)
    st.plotly_chart(fig3, use_container_width=True)

    # Fraud stacked bar
    sh("Fraud vs Non-Fraud by Segment")
    fs = d.groupby("Customer Segment").agg(Fraud=("Is Fraud","sum"),
                                            Total=("Is Fraud","count")).reset_index()
    fs["Non-Fraud"] = fs["Total"] - fs["Fraud"]
    fig4 = go.Figure()
    fig4.add_trace(go.Bar(name="Non-Fraud", x=fs["Customer Segment"],
                          y=fs["Non-Fraud"], marker_color="#333333"))
    fig4.add_trace(go.Bar(name="Fraud",     x=fs["Customer Segment"],
                          y=fs["Fraud"],     marker_color=RED))
    fig4.update_layout(barmode="stack")
    fig4 = theme(fig4, "", 300)
    st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════
# TAB 5 — TRENDS
# ══════════════════════════════════════════════
with tabs[4]:
    sh("Time Intelligence & Trends")

    ts = (d.groupby("Order Date")
           .agg(Orders=("Order Id","count"), Sales=("Sales","sum"),
                Profit=("Profit","sum"), Fraud=("Is Fraud","sum"))
           .reset_index().sort_values("Order Date"))
    ts.set_index("Order Date", inplace=True)
    ts["Roll3M_Orders"]   = ts["Orders"].rolling("90D").sum()
    ts["Roll12M_Revenue"] = ts["Sales"].rolling("365D").sum()
    ts.reset_index(inplace=True)

    ca, cb = st.columns(2)
    with ca:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=ts["Order Date"], y=ts["Roll3M_Orders"],
                                 mode="lines", line=dict(color=RED,width=2),
                                 fill="tozeroy", fillcolor="rgba(255,31,31,0.09)",
                                 name="3M Rolling"))
        fig = theme(fig, "Rolling 3-Month Orders", 300)
        st.plotly_chart(fig, use_container_width=True)

    with cb:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(x=ts["Order Date"], y=ts["Roll12M_Revenue"],
                                  mode="lines", line=dict(color=GOLD,width=2),
                                  fill="tozeroy", fillcolor="rgba(255,215,0,0.07)",
                                  name="12M Rolling"))
        fig2 = theme(fig2, "Rolling 12-Month Revenue", 300)
        st.plotly_chart(fig2, use_container_width=True)

    # YoY
    sh("Year-over-Year Sales vs Profit")
    yoy = d.groupby("Year").agg(Sales=("Sales","sum"), Profit=("Profit","sum")).reset_index()
    fig3 = go.Figure()
    fig3.add_trace(go.Bar(name="Sales",  x=yoy["Year"].astype(str), y=yoy["Sales"],  marker_color=RED))
    fig3.add_trace(go.Bar(name="Profit", x=yoy["Year"].astype(str), y=yoy["Profit"], marker_color=GOLD))
    fig3.update_layout(barmode="group")
    fig3 = theme(fig3, "", 300)
    st.plotly_chart(fig3, use_container_width=True)

    # Monthly Profit by Dept (Year filter)
    sh("Department Profit Trend")
    yr_opts = sorted(d["Year"].dropna().unique(), reverse=True)
    yr_pick = st.selectbox("Select Year", yr_opts)
    dept_ts = (d[d["Year"]==yr_pick]
               .groupby(["Month","Department Name"])["Profit"]
               .sum().reset_index().sort_values("Month"))
    fig4 = px.line(dept_ts, x="Month", y="Profit", color="Department Name",
                   color_discrete_sequence=px.colors.sequential.Reds_r)
    fig4 = theme(fig4, f"Monthly Profit by Department — {yr_pick}", 360)
    st.plotly_chart(fig4, use_container_width=True)

    # Fraud trend
    sh("Fraud Order Trend by Segment")
    ft = (d[d["Is Fraud"]==1]
           .groupby(["Month","Customer Segment"])["Is Fraud"]
           .sum().reset_index().sort_values("Month"))
    fig5 = px.line(ft, x="Month", y="Is Fraud", color="Customer Segment",
                   color_discrete_sequence=[RED,"#FF6B6B","#FFAAAA"])
    fig5 = theme(fig5, "", 300)
    st.plotly_chart(fig5, use_container_width=True)

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown(f"""
<div style="margin-top:36px;padding:12px;background:{CARD};border-radius:8px;
            border-top:2px solid {RED};text-align:center;color:{GREY};font-size:11px">
  Smart Supply Chain Dashboard · Streamlit + Plotly ·
  Schema: Fact_Orders · Dim_Shipping · Dim_Date · Dim_Department · Dim_Customer · Dim_Product
</div>""", unsafe_allow_html=True)