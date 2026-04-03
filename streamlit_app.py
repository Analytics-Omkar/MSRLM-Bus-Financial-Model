"""
MSRLM Tourist Bus Fleet — Interactive Financial Dashboard
Zilla Parishad, Dharashiv District, Maharashtra
Run: streamlit run streamlit_app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import io, os

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MSRLM Bus Fleet | Dharashiv ZP",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── STYLES ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main-header {
    background: linear-gradient(135deg, #0d2137 0%, #0e6b82 100%);
    padding: 2rem 2.5rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    color: white;
}
.main-header h1 { font-size: 1.9rem; font-weight: 700; margin: 0; color: white; }
.main-header p  { font-size: 0.95rem; color: #aed6e8; margin: 0.3rem 0 0 0; }

.kpi-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 1.1rem 1.2rem;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.kpi-value { font-size: 1.55rem; font-weight: 700; margin: 0; }
.kpi-label { font-size: 0.75rem; color: #6b7c93; margin: 0.2rem 0 0 0; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em; }

.section-header {
    font-size: 1.1rem; font-weight: 700;
    color: #0d2137;
    border-left: 4px solid #0e6b82;
    padding-left: 0.75rem;
    margin: 1.5rem 0 0.75rem 0;
}
.insight-box {
    background: #f0f7ff;
    border: 1px solid #bee3f8;
    border-radius: 8px;
    padding: 0.85rem 1.1rem;
    margin: 0.4rem 0;
    font-size: 0.9rem;
    color: #2d4a6e;
}
.rec-box {
    border-radius: 8px;
    padding: 1rem 1.2rem;
    margin: 0.5rem 0;
}
.status-viable  { color: #27ae60; font-weight: 700; }
.status-loss    { color: #c0392b; font-weight: 700; }
.footer {
    text-align: center; color: #9aa5b4;
    font-size: 0.78rem; padding: 1.5rem 0 0.5rem 0;
    border-top: 1px solid #e2e8f0; margin-top: 2rem;
}
</style>
""", unsafe_allow_html=True)

# ─── COLOUR PALETTE ──────────────────────────────────────────────────────────
C_NAVY   = "#0d2137"
C_TEAL   = "#0e6b82"
C_GOLD   = "#c8992a"
C_GREEN  = "#27ae60"
C_RED    = "#c0392b"
C_ORANGE = "#e67e22"
C_PURPLE = "#8e44ad"

def styled_ax(ax, title=""):
    ax.set_facecolor("#f9fbfd")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#d0d7e2")
    ax.spines["bottom"].set_color("#d0d7e2")
    ax.tick_params(colors="#555", labelsize=9)
    ax.grid(axis="y", linestyle="--", alpha=0.4, color="#e0e0e0")
    if title:
        ax.set_title(title, color=C_NAVY, fontsize=10, fontweight="bold", pad=10)

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf

# ─── DATA LOADING & COMPUTATION ──────────────────────────────────────────────
@st.cache_data
def load_and_compute(csv_path="routes_data.csv"):
    df = pd.read_csv(csv_path)

    df["annual_km"]                   = (df["route_length_km"] * 2
                                         * df["trips_per_day"]
                                         * df["operating_days_per_year"])
    df["annual_fuel_cost_inr"]        = df["annual_km"] * df["fuel_cost_per_km_inr"]
    df["annual_driver_salary_inr"]    = df["driver_salary_per_month_inr"] * 12
    df["annual_conductor_salary_inr"] = df["conductor_salary_per_month_inr"] * 12
    df["annual_maintenance_inr"]      = df["maintenance_per_month_inr"] * 12
    df["annual_misc_inr"]             = df["misc_cost_per_month_inr"] * 12
    df["annual_depreciation_inr"]     = ((df["bus_cost_inr"] - df["salvage_value_inr"])
                                         / df["useful_life_years"])
    df["depreciation_per_day_inr"]    = df["annual_depreciation_inr"] / df["operating_days_per_year"]

    # Clamp occupancy to seat capacity
    df["avg_occupancy_seats_clamped"] = df[["avg_occupancy_seats","total_seats"]].min(axis=1)
    df["occupancy_pct"]               = (df["avg_occupancy_seats_clamped"] / df["total_seats"]) * 100

    df["daily_revenue_inr"]           = (df["avg_occupancy_seats_clamped"]
                                         * df["ticket_price_inr"]
                                         * df["trips_per_day"] * 2)
    df["annual_revenue_inr"]          = df["daily_revenue_inr"] * df["operating_days_per_year"]
    df["annual_operating_cost_inr"]   = (df["annual_fuel_cost_inr"]
                                         + df["annual_driver_salary_inr"]
                                         + df["annual_conductor_salary_inr"]
                                         + df["annual_maintenance_inr"]
                                         + df["annual_misc_inr"])
    df["annual_total_cost_inr"]       = df["annual_operating_cost_inr"] + df["annual_depreciation_inr"]
    df["contribution_margin_inr"]     = df["annual_revenue_inr"] - df["annual_operating_cost_inr"]
    df["net_profit_loss_inr"]         = df["annual_revenue_inr"] - df["annual_total_cost_inr"]
    df["monthly_revenue_inr"]         = df["annual_revenue_inr"] / 12
    df["monthly_cost_inr"]            = df["annual_total_cost_inr"] / 12
    df["monthly_net_inr"]             = df["net_profit_loss_inr"] / 12

    df["be_occupancy_seats"]          = (df["annual_total_cost_inr"]
                                         / (df["ticket_price_inr"] * df["trips_per_day"] * 2
                                            * df["operating_days_per_year"]))
    df["be_occupancy_pct"]            = (df["be_occupancy_seats"] / df["total_seats"]) * 100
    df["be_trips_per_day"]            = (df["annual_total_cost_inr"]
                                         / (df["avg_occupancy_seats_clamped"]
                                            * df["ticket_price_inr"] * 2
                                            * df["operating_days_per_year"]))
    df["gap_seats"]                   = df["avg_occupancy_seats_clamped"] - df["be_occupancy_seats"]
    df["book_value_inr"]              = df["bus_cost_inr"] - df["annual_depreciation_inr"]
    df["lease_annual_inr"]            = df["annual_revenue_inr"] * 0.35
    df["ppp_annual_share_inr"]        = df["annual_revenue_inr"] * 0.25
    df["scrap_value_inr"]             = df["salvage_value_inr"] * 0.85

    return df

# ─── SIDEBAR ─────────────────────────────────────────────────────────────────
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/f/f9/Maharashtra_Seal.svg/200px-Maharashtra_Seal.svg.png", width=60)
st.sidebar.markdown("## 🚌 Fleet Parameters")
st.sidebar.markdown("Adjust assumptions to explore scenarios")

uploaded = st.sidebar.file_uploader("📂 Upload routes_data.csv", type=["csv"])
if uploaded:
    with open("data/routes_data.csv", "wb") as f:
        f.write(uploaded.read())
    st.cache_data.clear()

df_base = load_and_compute()

st.sidebar.markdown("---")
st.sidebar.markdown("### Override Assumptions")

# ✅ UPDATED TICKET PRICE CONTROL
st.sidebar.markdown("#### 🎫 Ticket Pricing Policy Range")
ticket_override  = st.sidebar.slider(
    "Fare per Seat (₹)",
    min_value=500,
    max_value=2000,
    value=int(np.clip(df_base["ticket_price_inr"].iloc[0], 500, 2000)),
    step=50
)

fuel_override    = st.sidebar.number_input("Fuel Cost (₹/km)", 40, 200,
                                            int(df_base["fuel_cost_per_km_inr"].iloc[0]), step=5)

occ_override_r1  = st.sidebar.slider("R001 Occupancy (seats)", 1, int(df_base["total_seats"].iloc[0]),
                                      int(df_base["avg_occupancy_seats_clamped"].iloc[0]))

occ_override_r2  = st.sidebar.slider("R002 Occupancy (seats)", 1, int(df_base["total_seats"].iloc[1]),
                                      int(df_base["avg_occupancy_seats_clamped"].iloc[1]))

current_age      = st.sidebar.slider("Bus Age (years)", 0, 15, 1)
lease_pct        = st.sidebar.slider("Lease Fee (% of revenue)", 10, 60, 35)
ppp_pct          = st.sidebar.slider("PPP Revenue Share to ZP (%)", 10, 50, 25)

st.sidebar.markdown("---")
st.sidebar.markdown("### Navigation")
page = st.sidebar.radio("Go to", [
    "📊 Dashboard",
    "💰 Financial Model",
    "📈 Break-even Analysis",
    "🏦 Asset Valuation",
    "🔮 Scenario Simulation",
    "📋 Data Table",
])

# ─── RECOMPUTE WITH OVERRIDES ─────────────────────────────────────────────────
def recompute(df_in, ticket, fuel, occ_list, age, lp, pp):
    df = df_in.copy()
    df["ticket_price_inr"]            = ticket
    df["fuel_cost_per_km_inr"]        = fuel
    df["avg_occupancy_seats_clamped"] = occ_list
    df["occupancy_pct"]               = (df["avg_occupancy_seats_clamped"] / df["total_seats"]) * 100
    df["annual_fuel_cost_inr"]        = df["annual_km"] * fuel
    df["daily_revenue_inr"]           = df["avg_occupancy_seats_clamped"] * ticket * df["trips_per_day"] * 2
    df["annual_revenue_inr"]          = df["daily_revenue_inr"] * df["operating_days_per_year"]
    df["annual_operating_cost_inr"]   = (df["annual_fuel_cost_inr"]
                                         + df["annual_driver_salary_inr"]
                                         + df["annual_conductor_salary_inr"]
                                         + df["annual_maintenance_inr"]
                                         + df["annual_misc_inr"])
    df["annual_total_cost_inr"]       = df["annual_operating_cost_inr"] + df["annual_depreciation_inr"]
    df["contribution_margin_inr"]     = df["annual_revenue_inr"] - df["annual_operating_cost_inr"]
    df["net_profit_loss_inr"]         = df["annual_revenue_inr"] - df["annual_total_cost_inr"]
    df["monthly_revenue_inr"]         = df["annual_revenue_inr"] / 12
    df["monthly_cost_inr"]            = df["annual_total_cost_inr"] / 12
    df["monthly_net_inr"]             = df["net_profit_loss_inr"] / 12
    df["be_occupancy_seats"]          = (df["annual_total_cost_inr"]
                                         / (ticket * df["trips_per_day"] * 2 * df["operating_days_per_year"]))
    df["be_occupancy_pct"]            = df["be_occupancy_seats"] / df["total_seats"] * 100
    df["be_trips_per_day"]            = (df["annual_total_cost_inr"]
                                         / (df["avg_occupancy_seats_clamped"] * ticket * 2
                                            * df["operating_days_per_year"]))
    df["gap_seats"]                   = df["avg_occupancy_seats_clamped"] - df["be_occupancy_seats"]
    df["book_value_inr"]              = df["bus_cost_inr"] - (df["annual_depreciation_inr"] * age)
    df["book_value_inr"]              = df["book_value_inr"].clip(lower=df["scrap_value_inr"])
    df["lease_annual_inr"]            = df["annual_revenue_inr"] * (lp/100)
    df["ppp_annual_share_inr"]        = df["annual_revenue_inr"] * (pp/100)
    return df

df = recompute(df_base,
               ticket_override, fuel_override,
               [occ_override_r1, occ_override_r2],
               current_age, lease_pct, ppp_pct)

# Fleet aggregates
fleet_rev      = df["annual_revenue_inr"].sum()
fleet_op_cost  = df["annual_operating_cost_inr"].sum()
fleet_tot_cost = df["annual_total_cost_inr"].sum()
fleet_net      = df["net_profit_loss_inr"].sum()
fleet_cm       = df["contribution_margin_inr"].sum()
fleet_book     = df["book_value_inr"].sum()
fleet_lease    = df["lease_annual_inr"].sum()

# ─── HEADER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🚌 MSRLM Tourist Bus Fleet — Financial Model</h1>
  <p>Zilla Parishad, Dharashiv District, Maharashtra &nbsp;|&nbsp;
     17-Seater Tourist Buses &nbsp;|&nbsp; 2 Routes &nbsp;|&nbsp; Interactive Policy Dashboard</p>
</div>
""", unsafe_allow_html=True)

# ─── PAGE: DASHBOARD ─────────────────────────────────────────────────────────
if page == "📊 Dashboard":
    st.markdown('<div class="section-header">Fleet KPIs</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    def kpi(col, val, label, color):
        col.markdown(f"""
        <div class="kpi-card">
          <p class="kpi-value" style="color:{color}">{val}</p>
          <p class="kpi-label">{label}</p>
        </div>""", unsafe_allow_html=True)

    kpi(c1, f"₹{fleet_rev/1e5:.1f}L",     "Annual Revenue",       C_TEAL)
    kpi(c2, f"₹{fleet_tot_cost/1e5:.1f}L", "Annual Total Cost",    C_RED)
    kpi(c3, f"₹{fleet_net/1e5:.1f}L",      "Net Profit / Loss",    C_GREEN if fleet_net>=0 else C_RED)
    kpi(c4, f"₹{fleet_cm/1e5:.1f}L",       "Contribution Margin",  C_ORANGE)
    kpi(c5, f"₹{fleet_book/1e5:.1f}L",     "Fleet Book Value",     C_PURPLE)
    kpi(c6, f"₹{fleet_lease/1e5:.1f}L",    "Lease Income/yr",      C_GOLD)

    st.markdown('<div class="section-header">Financial Overview Charts</div>', unsafe_allow_html=True)
    col_l, col_r = st.columns(2)

    # Chart: Revenue vs Cost stacked
    with col_l:
        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor("white")
        x = np.arange(len(df))
        fuel_v = df["annual_fuel_cost_inr"].values/1e5
        sal_v  = (df["annual_driver_salary_inr"]+df["annual_conductor_salary_inr"]).values/1e5
        mnt_v  = (df["annual_maintenance_inr"]+df["annual_misc_inr"]).values/1e5
        dep_v  = df["annual_depreciation_inr"].values/1e5
        ax.bar(x, fuel_v,                     color="#e74c3c", label="Fuel", width=0.45)
        ax.bar(x, sal_v,   bottom=fuel_v,     color="#3498db", label="Salaries", width=0.45)
        ax.bar(x, mnt_v,   bottom=fuel_v+sal_v, color="#f39c12", label="Maint & Misc", width=0.45)
        ax.bar(x, dep_v,   bottom=fuel_v+sal_v+mnt_v, color="#9b59b6", label="Depreciation", width=0.45)
        rev_v = df["annual_revenue_inr"].values/1e5
        for i, rv in enumerate(rev_v):
            ax.plot([i-0.28, i+0.28], [rv, rv], color=C_GREEN, lw=2.5, zorder=5)
            ax.text(i, rv+1, f"Rev ₹{rv:.0f}L", ha="center", fontsize=8, color=C_GREEN, fontweight="bold")
        ax.set_xticks(x)
        ax.set_xticklabels(["R001: Tuljapur","R002: Naldurg"], fontsize=9)
        ax.set_ylabel("₹ Lakhs")
        ax.legend(fontsize=8, loc="upper right")
        styled_ax(ax, "Annual Cost Breakdown vs Revenue")
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # Chart: Net P/L and Contribution Margin
    with col_r:
        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor("white")
        x = np.arange(len(df))
        w = 0.33
        cm_v  = df["contribution_margin_inr"].values/1e5
        net_v = df["net_profit_loss_inr"].values/1e5
        ax.bar(x-w/2, cm_v,  width=w, color=[C_GREEN if v>=0 else C_RED for v in cm_v],  label="Contribution Margin", alpha=0.85)
        ax.bar(x+w/2, net_v, width=w, color=[C_TEAL if v>=0 else "#8B0000" for v in net_v], label="Net P/L", alpha=0.85)
        ax.axhline(0, color="#555", lw=1, ls="--")
        for i, (cm, net) in enumerate(zip(cm_v, net_v)):
            ax.text(i-w/2, cm+(1 if cm>=0 else -3), f"₹{cm:.0f}L", ha="center", fontsize=8, color="#333")
            ax.text(i+w/2, net+(1 if net>=0 else -3), f"₹{net:.0f}L", ha="center", fontsize=8, color="#333")
        ax.set_xticks(x); ax.set_xticklabels(["R001: Tuljapur","R002: Naldurg"], fontsize=9)
        ax.set_ylabel("₹ Lakhs"); ax.legend(fontsize=8)
        styled_ax(ax, "Contribution Margin & Net Profit/Loss")
        st.pyplot(fig, use_container_width=True)
        plt.close(fig)

    # Insights
    st.markdown('<div class="section-header">Key Insights</div>', unsafe_allow_html=True)
    viable_count = (df["net_profit_loss_inr"] > 0).sum()
    insights = [
        f"<b>{viable_count}/2 routes</b> are profitable at current settings.",
        f"Break-even occupancy: <b>{df['be_occupancy_pct'].min():.0f}%–{df['be_occupancy_pct'].max():.0f}%</b> vs actual avg <b>{df['occupancy_pct'].mean():.0f}%</b>.",
        f"Fuel dominates costs at <b>>96%</b> of operating cost (₹{df['annual_fuel_cost_inr'].mean()/1e5:.1f}L/bus/yr).",
        f"High ticket price (₹{ticket_override:,}/seat) — full bus generates <b>₹{df['total_seats'].iloc[0]*ticket_override*2:,}/trip</b>.",
        f"Leasing out both buses generates <b>₹{fleet_lease/1e5:.1f}L/yr</b> to ZP with zero management.",
        f"Fleet book value: <b>₹{fleet_book/1e5:.1f}L</b> at Year {current_age} — preserve through monetisation.",
    ]
    for ins in insights:
        st.markdown(f'<div class="insight-box">💡 {ins}</div>', unsafe_allow_html=True)

# ─── PAGE: FINANCIAL MODEL ────────────────────────────────────────────────────
elif page == "💰 Financial Model":
    st.markdown('<div class="section-header">Route-Level Financial Summary</div>', unsafe_allow_html=True)

    for _, row in df.iterrows():
        status_html = ('<span class="status-viable">✅ VIABLE</span>'
                       if row["net_profit_loss_inr"] >= 0
                       else '<span class="status-loss">⚠️ LOSS-MAKING</span>')
        st.markdown(f"#### {row['route_name']} &nbsp; {status_html}", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Annual Revenue",      f"₹{row['annual_revenue_inr']/1e5:.2f}L")
        c2.metric("Annual Total Cost",   f"₹{row['annual_total_cost_inr']/1e5:.2f}L")
        c3.metric("Contribution Margin", f"₹{row['contribution_margin_inr']/1e5:.2f}L",
                  delta=f"₹{row['contribution_margin_inr']/1e5:.2f}L")
        c4.metric("Net Profit / Loss",   f"₹{row['net_profit_loss_inr']/1e5:.2f}L",
                  delta=f"₹{row['net_profit_loss_inr']/1e5:.2f}L")
        st.markdown("---")

    st.markdown('<div class="section-header">Detailed Cost Breakdown Table</div>', unsafe_allow_html=True)
    cost_data = []
    for _, row in df.iterrows():
        cost_data.append({
            "Route": row["route_id"],
            "Fuel (₹L)": f"{row['annual_fuel_cost_inr']/1e5:.2f}",
            "Driver Salary (₹L)": f"{row['annual_driver_salary_inr']/1e5:.2f}",
            "Conductor Salary (₹L)": f"{row['annual_conductor_salary_inr']/1e5:.2f}",
            "Maintenance (₹L)": f"{row['annual_maintenance_inr']/1e5:.2f}",
            "Misc (₹L)": f"{row['annual_misc_inr']/1e5:.2f}",
            "Depreciation (₹L)": f"{row['annual_depreciation_inr']/1e5:.2f}",
            "Total Cost (₹L)": f"{row['annual_total_cost_inr']/1e5:.2f}",
            "Revenue (₹L)": f"{row['annual_revenue_inr']/1e5:.2f}",
            "Net P/L (₹L)": f"{row['net_profit_loss_inr']/1e5:.2f}",
        })
    st.dataframe(pd.DataFrame(cost_data), use_container_width=True)

    st.markdown('<div class="section-header">Monthly Cashflow Summary</div>', unsafe_allow_html=True)
    monthly_data = []
    for _, row in df.iterrows():
        monthly_data.append({
            "Route": row["route_id"],
            "Monthly Revenue (₹)": f"₹{row['monthly_revenue_inr']:,.0f}",
            "Monthly Cost (₹)": f"₹{row['monthly_cost_inr']:,.0f}",
            "Monthly Net (₹)": f"₹{row['monthly_net_inr']:,.0f}",
            "Daily Revenue (₹)": f"₹{row['daily_revenue_inr']:,.0f}",
            "Depreciation/Day (₹)": f"₹{row['depreciation_per_day_inr']:,.0f}",
        })
    st.dataframe(pd.DataFrame(monthly_data), use_container_width=True)

# ─── PAGE: BREAK-EVEN ─────────────────────────────────────────────────────────
elif page == "📈 Break-even Analysis":
    st.markdown('<div class="section-header">Break-even Summary</div>', unsafe_allow_html=True)

    for _, row in df.iterrows():
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(f"{row['route_id']} — BE Seats", f"{row['be_occupancy_seats']:.1f}")
        c2.metric("BE Occupancy %",   f"{row['be_occupancy_pct']:.1f}%")
        c3.metric("Actual Occupancy", f"{row['occupancy_pct']:.1f}%",
                  delta=f"{row['occupancy_pct']-row['be_occupancy_pct']:.1f}%")
        c4.metric("Gap (Actual − BE)", f"{row['gap_seats']:+.1f} seats")

    st.markdown('<div class="section-header">Sensitivity Curves</div>', unsafe_allow_html=True)
    fig, axes = plt.subplots(1, 2, figsize=(14, 5))
    fig.patch.set_facecolor("white")
    occ_range = np.linspace(5, 100, 300)
    colors_s  = [C_TEAL, C_ORANGE]
    for idx, (_, row) in enumerate(df.iterrows()):
        ax = axes[idx]
        nets = []
        for occ in occ_range:
            seats = (occ/100) * row["total_seats"]
            rev   = seats * ticket_override * row["trips_per_day"] * 2 * row["operating_days_per_year"]
            nets.append((rev - row["annual_total_cost_inr"]) / 1e5)
        nets_arr = np.array(nets)
        ax.plot(occ_range, nets_arr, color=colors_s[idx], lw=2.5, label="Net P/L")
        ax.fill_between(occ_range, nets_arr, 0, where=(nets_arr>=0), alpha=0.15, color=C_GREEN)
        ax.fill_between(occ_range, nets_arr, 0, where=(nets_arr<0),  alpha=0.10, color=C_RED)
        ax.axhline(0, color="#555", lw=1)
        be = row["be_occupancy_pct"]
        act = row["occupancy_pct"]
        ax.axvline(be,  color=C_RED,   lw=1.5, ls="--", label=f"BE: {be:.0f}%")
        ax.axvline(act, color=C_GREEN, lw=1.5, ls="--", label=f"Actual: {act:.0f}%")
        ax.set_xlabel("Occupancy %"); ax.set_ylabel("Net P/L (₹ Lakhs)")
        ax.legend(fontsize=8.5)
        styled_ax(ax, f"{row['route_id']}: {row['route_name'][:35]}")
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    st.markdown('<div class="section-header">Sensitivity Table</div>', unsafe_allow_html=True)
    occ_levels = list(range(10, 105, 5))
    sens_rows = []
    for occ in occ_levels:
        row_d = {"Occupancy %": f"{occ}%"}
        for _, row in df.iterrows():
            seats = (occ/100) * row["total_seats"]
            rev   = seats * ticket_override * row["trips_per_day"] * 2 * row["operating_days_per_year"]
            net   = (rev - row["annual_total_cost_inr"]) / 1e5
            row_d[f"{row['route_id']} Net (₹L)"] = f"{net:+.2f}"
        sens_rows.append(row_d)
    sens_df = pd.DataFrame(sens_rows)

    def color_net(val):
        try:
            v = float(val)
            color = "#d4edda" if v >= 0 else "#f8d7da"
            return f"background-color: {color}"
        except:
            return ""

    st.dataframe(
        sens_df.style.applymap(color_net, subset=[c for c in sens_df.columns if "Net" in c]),
        use_container_width=True, height=500
    )

# ─── PAGE: ASSET VALUATION ────────────────────────────────────────────────────
elif page == "🏦 Asset Valuation":
    st.markdown('<div class="section-header">Fleet Asset Valuation</div>', unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Fleet Book Value",       f"₹{fleet_book/1e5:.2f}L", help=f"Year {current_age}")
    c2.metric("Annual Lease Income",    f"₹{fleet_lease/1e5:.2f}L", help=f"{lease_pct}% of revenue")
    c3.metric("Annual PPP Share",       f"₹{df['ppp_annual_share_inr'].sum()/1e5:.2f}L", help=f"{ppp_pct}% of revenue")
    c4.metric("Fleet Scrap Value",      f"₹{df['scrap_value_inr'].sum()/1e5:.2f}L")

    val_data = []
    for _, row in df.iterrows():
        val_data.append({
            "Route": row["route_id"],
            "Acquisition Cost (₹)": f"₹{row['bus_cost_inr']:,.0f}",
            f"Book Value Yr{current_age} (₹)": f"₹{row['book_value_inr']:,.0f}",
            f"Lease Income/yr @ {lease_pct}% (₹)": f"₹{row['lease_annual_inr']:,.0f}",
            f"PPP Share/yr @ {ppp_pct}% (₹)": f"₹{row['ppp_annual_share_inr']:,.0f}",
            "Scrap Value (₹)": f"₹{row['scrap_value_inr']:,.0f}",
            "Annual Depreciation (₹)": f"₹{row['annual_depreciation_inr']:,.0f}",
            "Dep/km (₹)": f"₹{row['annual_depreciation_inr']/row['annual_km']:.2f}",
            "Dep/day (₹)": f"₹{row['depreciation_per_day_inr']:.0f}",
        })
    st.dataframe(pd.DataFrame(val_data), use_container_width=True)

    # Book value depreciation over time
    st.markdown('<div class="section-header">Book Value Depreciation Over Fleet Life</div>', unsafe_allow_html=True)
    fig, ax = plt.subplots(figsize=(10, 4))
    fig.patch.set_facecolor("white")
    years = np.arange(0, 16)
    for idx, (_, row) in enumerate(df.iterrows()):
        bvs = [max(row["bus_cost_inr"] - row["annual_depreciation_inr"]*y, row["scrap_value_inr"]) / 1e5
               for y in years]
        ax.plot(years, bvs, marker="o", markersize=4, label=row["route_id"],
                color=[C_TEAL, C_ORANGE][idx], lw=2)
    ax.axhline(df["scrap_value_inr"].mean()/1e5, color=C_RED, ls="--", lw=1.2, label="Salvage Floor")
    ax.set_xlabel("Year"); ax.set_ylabel("Book Value (₹ Lakhs)")
    ax.legend(fontsize=9)
    styled_ax(ax, "Bus Book Value Over 15-Year Life")
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

# ─── PAGE: SCENARIO SIMULATION ────────────────────────────────────────────────
elif page == "🔮 Scenario Simulation":
    st.markdown('<div class="section-header">Policy Scenario Comparison</div>', unsafe_allow_html=True)
    st.info("Scenario B (Lease) and C (PPP) returns are ZP's net gain after bearing depreciation cost.")

    scen_rows = []
    for _, row in df.iterrows():
        scen_a = row["net_profit_loss_inr"]
        scen_b = row["lease_annual_inr"] - row["annual_depreciation_inr"]
        scen_c = row["ppp_annual_share_inr"] - row["annual_depreciation_inr"]
        scen_d = row["scrap_value_inr"]
        best   = max(scen_a, scen_b, scen_c, scen_d)
        scen_rows.append({
            "Route": f"{row['route_id']} — {row['route_name'][:35]}",
            "A: Direct Ops (₹L)": f"{scen_a/1e5:+.2f}",
            "B: Lease Out (₹L)": f"{scen_b/1e5:+.2f}",
            "C: PPP Share (₹L)": f"{scen_c/1e5:+.2f}",
            "D: Scrap ₹L (once)": f"{scen_d/1e5:.2f}",
            "Best Option": ("A" if best==scen_a else "B" if best==scen_b else "C" if best==scen_c else "D"),
        })

    scen_df = pd.DataFrame(scen_rows)

    def color_scen(val):
        try:
            v = float(val)
            return "background-color: #d4edda" if v >= 0 else "background-color: #f8d7da"
        except:
            return ""

    st.dataframe(
        scen_df.style.map(color_scen, subset=["A: Direct Ops (₹L)","B: Lease Out (₹L)","C: PPP Share (₹L)"]),
        use_container_width=True
    )

    # Scenario chart
    fig, ax = plt.subplots(figsize=(11, 4.5))
    fig.patch.set_facecolor("white")
    scen_labels = ["A: Direct\nOps", "B: Lease\nOut", f"C: PPP\n{ppp_pct}%", "D: Scrap\n(once)"]
    clrs = [C_RED, C_GREEN, C_TEAL, C_ORANGE]
    x_s  = np.arange(4)
    w_s  = 0.32
    for idx, (_, row) in enumerate(df.iterrows()):
        scen_a = row["net_profit_loss_inr"]/1e5
        scen_b = (row["lease_annual_inr"] - row["annual_depreciation_inr"])/1e5
        scen_c = (row["ppp_annual_share_inr"] - row["annual_depreciation_inr"])/1e5
        scen_d = row["scrap_value_inr"]/1e5
        vals   = [scen_a, scen_b, scen_c, scen_d]
        offset = (idx - 0.5) * w_s
        bars = ax.bar(x_s + offset, vals, width=w_s, color=clrs,
                      alpha=0.9 if idx==0 else 0.6, edgecolor="white",
                      label=row["route_id"])
        for bar, v in zip(bars, vals):
            ax.text(bar.get_x()+bar.get_width()/2,
                    v+(0.5 if v>=0 else -2),
                    f"₹{v:.1f}L", ha="center", fontsize=7.5, color="#333")
    ax.axhline(0, color="#555", lw=1, ls="--")
    ax.set_xticks(x_s); ax.set_xticklabels(scen_labels, fontsize=10)
    ax.set_ylabel("Annual Net Gain to ZP (₹ Lakhs)")
    ax.legend(fontsize=9)
    styled_ax(ax, "Policy Scenario Comparison — ZP Net Annual Position")
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

    # Recommendations
    st.markdown('<div class="section-header">Policy Recommendations</div>', unsafe_allow_html=True)
    recs = [
        ("🟦 Immediate Action", C_TEAL,
         "Route R002 (Naldurg) should be explored for private lease arrangement. "
         f"At {lease_pct}% revenue lease, ZP earns ₹{df.iloc[1]['lease_annual_inr']/1e5:.1f}L/yr with zero operational risk."),
        ("🟩 Short-Term", C_GREEN,
         "Route R001 (Tuljapur) is the stronger performer. Push occupancy to full 17 seats through "
         "temple trust tie-ups and school excursion empanelment. Full occupancy yields ₹" +
         f"{df.iloc[0]['total_seats']*ticket_override*2*df.iloc[0]['trips_per_day']*df.iloc[0]['operating_days_per_year']/1e5:.1f}L/yr."),
        ("🟨 Medium-Term", C_GOLD,
         f"Structure a PPP arrangement for both routes. At {ppp_pct}% revenue share, "
         f"ZP earns ₹{df['ppp_annual_share_inr'].sum()/1e5:.1f}L/yr with no management burden. "
         "Include minimum guaranteed revenue clause."),
        ("🟧 Cost Side", C_ORANGE,
         "Fuel is >96% of operating cost. Evaluate CNG retrofitting (~₹2.5L one-time/bus) "
         "to cut fuel cost by 30–40%, saving ₹13–18L over the bus life."),
    ]
    for title, color, text in recs:
        st.markdown(f"""
        <div class="rec-box" style="background:{color}18; border-left:4px solid {color};">
          <strong style="color:{color}">{title}</strong><br/>
          <span style="font-size:0.9rem; color:#2c3e50">{text}</span>
        </div>""", unsafe_allow_html=True)

# ─── PAGE: DATA TABLE ─────────────────────────────────────────────────────────
elif page == "📋 Data Table":
    st.markdown('<div class="section-header">Full Computed Dataset</div>', unsafe_allow_html=True)
    show_cols = [
        "route_id","route_name","total_seats","ticket_price_inr",
        "fuel_cost_per_km_inr","avg_occupancy_seats_clamped","occupancy_pct",
        "annual_km","annual_revenue_inr","annual_operating_cost_inr",
        "annual_depreciation_inr","annual_total_cost_inr",
        "contribution_margin_inr","net_profit_loss_inr",
        "be_occupancy_seats","be_occupancy_pct","gap_seats",
        "book_value_inr","lease_annual_inr","ppp_annual_share_inr","scrap_value_inr",
    ]
    st.dataframe(df[show_cols], use_container_width=True)

    st.markdown('<div class="section-header">Download Outputs</div>', unsafe_allow_html=True)
    csv_buf = df[show_cols].to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download Dataset (CSV)", csv_buf,
                       "MSRLM_Fleet_Data.csv", "text/csv")

# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  MSRLM Tourist Bus Fleet Financial Model &nbsp;|&nbsp;
  Zilla Parishad, Dharashiv District &nbsp;|&nbsp;
  Built with Python &amp; Streamlit
</div>
""", unsafe_allow_html=True)
