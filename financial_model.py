"""
=============================================================================
MSRLM TOURIST BUS FLEET — FINANCIAL MODEL  (v2)
Zilla Parishad, Dharashiv District, Maharashtra
=============================================================================
17-seater tourist buses | ₹1,450 ticket | ₹90/km fuel
=============================================================================
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
import os, warnings
warnings.filterwarnings("ignore")

os.makedirs("outputs", exist_ok=True)

# ─── COLOUR PALETTE ──────────────────────────────────────────────────────────
C_NAVY   = "#0d2137"
C_TEAL   = "#0e6b82"
C_GOLD   = "#c8992a"
C_GREEN  = "#27ae60"
C_RED    = "#c0392b"
C_ORANGE = "#e67e22"
C_PURPLE = "#8e44ad"
C_LIGHT  = "#f5f7fa"
C_MID    = "#6b7c93"

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 1: PROBLEM FRAMING
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 1 — PROBLEM FRAMING")
print("=" * 70)
print("""
PROBLEM:
  Zilla Parishad, Dharashiv acquired 2 MSRLM tourist mini-buses (17-seater,
  ₹27.7 lakh each via GoM funds) for rural pilgrimage-tourism circuits.
  Buses are underutilised; revenue model needs validation.

KEY QUESTIONS:
  Q1. Which routes are financially viable (revenue > full cost)?
  Q2. What is the break-even occupancy and trips/day needed?
  Q3. What is the recoverable / monetisable asset value of the fleet?
  Q4. Which policy scenario maximises ZP's net financial position?
""")

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 2: SYSTEM UNDERSTANDING
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 2 — SYSTEM UNDERSTANDING")
print("=" * 70)
SYSTEM = {
    "Assets"         : "2 × 17-seat tourist mini-buses (GoM funded, ₹27.7L each)",
    "Routes"         : "Tuljapur Temple Circuit (215 km) | Naldurg Fort Circuit (216 km)",
    "Revenue Streams": "Ticketing (₹1,450/seat) | Charter leasing | PPP revenue share",
    "Demand Drivers" : "Tuljapur temple pilgrims | Naldurg fort tourists | School groups",
    "Key Costs"      : "Fuel ₹90/km | Driver ₹18k/mo | Conductor ₹14k/mo | Maintenance ₹8k/mo",
    "Stakeholders"   : "Zilla Parishad (owner) | MSRLM (scheme) | Drivers | Tourists",
}
for k, v in SYSTEM.items():
    print(f"  {k:20s}: {v}")
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 3: DATA DESIGN — SCHEMA
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 3 — DATA SCHEMA")
print("=" * 70)
SCHEMA = {
    "route_id"                      : "Unique route identifier",
    "route_name"                    : "Name / description of the route",
    "bus_cost_inr"                  : "Purchase cost of bus (₹)",
    "salvage_value_inr"             : "End-of-life salvage value (₹)",
    "useful_life_years"             : "Expected bus life (years)",
    "route_length_km"               : "One-way route distance (km)",
    "operating_days_per_year"       : "Days the bus operates per year",
    "trips_per_day"                 : "Return trips per operating day",
    "avg_occupancy_seats"           : "Average filled seats per trip",
    "total_seats"                   : "Total seating capacity (17)",
    "ticket_price_inr"              : "Ticket price per passenger (₹1,450)",
    "fuel_cost_per_km_inr"          : "Fuel cost per km (₹90)",
    "driver_salary_per_month_inr"   : "Monthly driver salary (₹)",
    "conductor_salary_per_month_inr": "Monthly conductor salary (₹)",
    "maintenance_per_month_inr"     : "Monthly maintenance cost (₹)",
    "misc_cost_per_month_inr"       : "Misc: permits, insurance, admin (₹)",
}
for f, d in SCHEMA.items():
    print(f"  {f:38s}: {d}")
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 4: DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 4 — DATA LOADING")
print("=" * 70)
df = pd.read_csv("routes_data.csv")
print(f"  Loaded {len(df)} routes.")
print(df[["route_id","route_name","total_seats","ticket_price_inr",
          "fuel_cost_per_km_inr","avg_occupancy_seats"]].to_string(index=False))
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 5: DATA CLEANING & PREPARATION
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 5 — DATA CLEANING & PREPARATION")
print("=" * 70)

# Annualise all monthly figures
df["annual_km"]                   = (df["route_length_km"] * 2
                                     * df["trips_per_day"]
                                     * df["operating_days_per_year"])
df["annual_fuel_cost_inr"]        = df["annual_km"] * df["fuel_cost_per_km_inr"]
df["annual_driver_salary_inr"]    = df["driver_salary_per_month_inr"] * 12
df["annual_conductor_salary_inr"] = df["conductor_salary_per_month_inr"] * 12
df["annual_maintenance_inr"]      = df["maintenance_per_month_inr"] * 12
df["annual_misc_inr"]             = df["misc_cost_per_month_inr"] * 12

# Straight-line depreciation
df["annual_depreciation_inr"]     = ((df["bus_cost_inr"] - df["salvage_value_inr"])
                                     / df["useful_life_years"])
df["depreciation_per_km_inr"]     = df["annual_depreciation_inr"] / df["annual_km"]
df["depreciation_per_day_inr"]    = df["annual_depreciation_inr"] / df["operating_days_per_year"]

# Daily revenue
df["daily_revenue_inr"]           = (df["avg_occupancy_seats"]
                                     * df["ticket_price_inr"]
                                     * df["trips_per_day"] * 2)
df["occupancy_pct"]               = (df["avg_occupancy_seats"] / df["total_seats"]) * 100

print("  Calculated columns added successfully.")
print(f"  Annual km per bus: {df['annual_km'].values}")
print(f"  Annual fuel cost:  {df['annual_fuel_cost_inr'].values}")
print(f"  Annual depreciation: {df['annual_depreciation_inr'].values}")
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 6: CORE FINANCIAL MODEL
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 6 — CORE FINANCIAL MODEL")
print("=" * 70)

df["annual_revenue_inr"]      = df["daily_revenue_inr"] * df["operating_days_per_year"]
df["annual_operating_cost_inr"] = (df["annual_fuel_cost_inr"]
                                   + df["annual_driver_salary_inr"]
                                   + df["annual_conductor_salary_inr"]
                                   + df["annual_maintenance_inr"]
                                   + df["annual_misc_inr"])
df["annual_total_cost_inr"]   = df["annual_operating_cost_inr"] + df["annual_depreciation_inr"]
df["contribution_margin_inr"] = df["annual_revenue_inr"] - df["annual_operating_cost_inr"]
df["net_profit_loss_inr"]     = df["annual_revenue_inr"] - df["annual_total_cost_inr"]
df["monthly_revenue_inr"]     = df["annual_revenue_inr"] / 12
df["monthly_cost_inr"]        = df["annual_total_cost_inr"] / 12
df["monthly_net_inr"]         = df["net_profit_loss_inr"] / 12

# Revenue per km and cost per km
df["revenue_per_km_inr"]      = df["annual_revenue_inr"] / df["annual_km"]
df["cost_per_km_inr"]         = df["annual_total_cost_inr"] / df["annual_km"]

for _, row in df.iterrows():
    status = "✅ VIABLE" if row["net_profit_loss_inr"] > 0 else "⚠️  LOSS-MAKING"
    print(f"\n  Route : {row['route_name']}")
    print(f"    Seats / Ticket Price   : {int(row['total_seats'])} seats | ₹{row['ticket_price_inr']:,.0f}/seat")
    print(f"    Avg Occupancy          : {row['avg_occupancy_seats']} seats ({row['occupancy_pct']:.1f}%)")
    print(f"    Annual Revenue         : ₹{row['annual_revenue_inr']:>12,.0f}")
    print(f"    Annual Operating Cost  : ₹{row['annual_operating_cost_inr']:>12,.0f}")
    print(f"    Annual Depreciation    : ₹{row['annual_depreciation_inr']:>12,.0f}")
    print(f"    Annual Total Cost      : ₹{row['annual_total_cost_inr']:>12,.0f}")
    print(f"    Contribution Margin    : ₹{row['contribution_margin_inr']:>12,.0f}")
    print(f"    Net Profit / Loss      : ₹{row['net_profit_loss_inr']:>12,.0f}")
    print(f"    Monthly Net            : ₹{row['monthly_net_inr']:>12,.0f}")
    print(f"    Revenue per km         : ₹{row['revenue_per_km_inr']:>8.2f}")
    print(f"    Cost per km            : ₹{row['cost_per_km_inr']:>8.2f}")
    print(f"    Status                 : {status}")
print()

# Fleet totals
fleet_rev      = df["annual_revenue_inr"].sum()
fleet_op_cost  = df["annual_operating_cost_inr"].sum()
fleet_tot_cost = df["annual_total_cost_inr"].sum()
fleet_net      = df["net_profit_loss_inr"].sum()
fleet_cm       = df["contribution_margin_inr"].sum()
print(f"  ── FLEET TOTALS ──────────────────────────────────────")
print(f"  Annual Revenue      : ₹{fleet_rev:>12,.0f}")
print(f"  Annual Total Cost   : ₹{fleet_tot_cost:>12,.0f}")
print(f"  Fleet Net P/L       : ₹{fleet_net:>12,.0f}")
print(f"  ─────────────────────────────────────────────────────\n")

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 7: BREAK-EVEN ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 7 — BREAK-EVEN ANALYSIS")
print("=" * 70)

df["be_occupancy_seats"] = (df["annual_total_cost_inr"]
                            / (df["ticket_price_inr"] * df["trips_per_day"] * 2
                               * df["operating_days_per_year"]))
df["be_occupancy_pct"]   = (df["be_occupancy_seats"] / df["total_seats"]) * 100
df["be_trips_per_day"]   = (df["annual_total_cost_inr"]
                            / (df["avg_occupancy_seats"] * df["ticket_price_inr"] * 2
                               * df["operating_days_per_year"]))
df["gap_seats"]          = df["avg_occupancy_seats"] - df["be_occupancy_seats"]
df["revenue_shortfall_inr"] = df["annual_total_cost_inr"] - df["annual_revenue_inr"]
df["revenue_shortfall_inr"] = df["revenue_shortfall_inr"].clip(lower=0)

for _, row in df.iterrows():
    print(f"\n  Route: {row['route_name']}")
    print(f"    Break-even Occupancy  : {row['be_occupancy_seats']:.1f} seats  ({row['be_occupancy_pct']:.1f}%)")
    print(f"    Actual Occupancy      : {row['avg_occupancy_seats']} seats  ({row['occupancy_pct']:.1f}%)")
    print(f"    Gap (actual - BE)     : {row['gap_seats']:+.1f} seats")
    print(f"    Break-even Trips/Day  : {row['be_trips_per_day']:.2f}")
    print(f"    Revenue Shortfall/Yr  : ₹{row['revenue_shortfall_inr']:,.0f}")

print("\n  ── Sensitivity Table: Occupancy % vs Net Profit (₹) ──")
occ_levels = [20, 30, 40, 50, 60, 70, 80, 90, 100]
header = f"  {'Occ%':>5} | " + " | ".join([f"{r['route_id']:>14}" for _, r in df.iterrows()])
print(header)
print("  " + "-" * len(header))
for occ in occ_levels:
    row_str = f"  {occ:>4}% | "
    for _, row in df.iterrows():
        seats = (occ / 100) * row["total_seats"]
        rev   = seats * row["ticket_price_inr"] * row["trips_per_day"] * 2 * row["operating_days_per_year"]
        net   = rev - row["annual_total_cost_inr"]
        row_str += f"₹{net:>13,.0f} | "
    print(row_str)
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 8: ASSET VALUATION
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 8 — ASSET VALUATION")
print("=" * 70)

current_age = 1
df["book_value_inr"]          = df["bus_cost_inr"] - (df["annual_depreciation_inr"] * current_age)
df["lease_annual_inr"]        = df["annual_revenue_inr"] * 0.35   # 35% of revenue as lease fee
df["ppp_annual_share_inr"]    = df["annual_revenue_inr"] * 0.25   # 25% rev share to ZP
df["scrap_value_inr"]         = df["salvage_value_inr"] * 0.85
df["recoverable_value_inr"]   = df["book_value_inr"]  # conservative = book value

fleet_book       = df["book_value_inr"].sum()
fleet_lease_pa   = df["lease_annual_inr"].sum()
fleet_ppp_pa     = df["ppp_annual_share_inr"].sum()
fleet_scrap      = df["scrap_value_inr"].sum()

print(f"  Fleet Book Value (Yr {current_age})     : ₹{fleet_book:>12,.0f}")
print(f"  Fleet Annual Lease Income         : ₹{fleet_lease_pa:>12,.0f}")
print(f"  Fleet PPP Revenue Share (25%/yr)  : ₹{fleet_ppp_pa:>12,.0f}")
print(f"  Fleet Scrap Value                 : ₹{fleet_scrap:>12,.0f}")
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 9: SCENARIO SIMULATION
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 9 — SCENARIO SIMULATION")
print("=" * 70)

scenarios_out = {}
for _, row in df.iterrows():
    rid = row["route_id"]
    scen_a = row["net_profit_loss_inr"]
    scen_b = row["lease_annual_inr"] - row["annual_depreciation_inr"]
    scen_c = row["ppp_annual_share_inr"] - row["annual_depreciation_inr"]
    scen_d = row["scrap_value_inr"]
    scenarios_out[rid] = {
        "Route"                       : row["route_name"][:42],
        "A: Direct Ops (₹/yr)"        : scen_a,
        "B: Lease Out (₹/yr)"         : scen_b,
        "C: PPP 25% Share (₹/yr)"     : scen_c,
        "D: Scrap (₹ one-time)"        : scen_d,
    }
    best_k = max({k: v for k, v in scenarios_out[rid].items() if k != "Route"}, key=lambda k: scenarios_out[rid][k])
    print(f"  {rid}: {row['route_name']}")
    for k, v in scenarios_out[rid].items():
        if k != "Route":
            marker = " ← BEST" if k == best_k else ""
            print(f"    {k:30s}: ₹{v:>12,.0f}{marker}")
    print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 10: INSIGHTS
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 10 — KEY INSIGHTS")
print("=" * 70)

viable = df[df["net_profit_loss_inr"] > 0]
insights = [
    f"{len(viable)}/{len(df)} routes are profitable under direct ZP operations at current occupancy.",
    f"Break-even occupancy: {df['be_occupancy_pct'].min():.0f}%–{df['be_occupancy_pct'].max():.0f}% load vs actual {df['occupancy_pct'].mean():.0f}% average.",
    f"Fuel dominates costs at ₹{df['annual_fuel_cost_inr'].mean()/1e5:.1f}L/yr per bus (>{df['annual_fuel_cost_inr'].mean()/df['annual_total_cost_inr'].mean()*100:.0f}% of total cost).",
    f"High ticket price (₹1,450) means even partial occupancy generates strong revenue per trip — ₹{df['daily_revenue_inr'].mean():,.0f}/day.",
    f"Fleet book value ₹{fleet_book/1e5:.1f}L; lease income alone covers depreciation with surplus.",
    f"Leasing out generates ₹{fleet_lease_pa/1e5:.1f}L/yr to ZP with zero operational burden.",
]
for i, ins in enumerate(insights, 1):
    print(f"  {i}. {ins}")
print()

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 11: VISUALISATIONS
# ─────────────────────────────────────────────────────────────────────────────
print("=" * 70)
print("STAGE 11 — GENERATING CHARTS")
print("=" * 70)

def styled_ax(ax, title=""):
    ax.set_facecolor("#f9fbfd")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#cccccc")
    ax.spines["bottom"].set_color("#cccccc")
    ax.tick_params(colors="#444", labelsize=9)
    ax.yaxis.label.set_color("#444")
    ax.xaxis.label.set_color("#444")
    ax.grid(axis="y", linestyle="--", alpha=0.45, color="#dddddd")
    if title:
        ax.set_title(title, color=C_NAVY, fontsize=10, fontweight="bold", pad=10)

route_short = ["R001: Tuljapur", "R002: Naldurg"]
x = np.arange(len(df))

# ── CHART 1: 6-panel master dashboard ─────────────────────────────────────
fig = plt.figure(figsize=(18, 11))
fig.patch.set_facecolor("#ffffff")
gs  = fig.add_gridspec(2, 3, hspace=0.38, wspace=0.32)

# Panel 1: Cost breakdown stacked bar
ax1 = fig.add_subplot(gs[0, 0])
fuel_v  = df["annual_fuel_cost_inr"].values / 1e5
sal_v   = (df["annual_driver_salary_inr"] + df["annual_conductor_salary_inr"]).values / 1e5
maint_v = (df["annual_maintenance_inr"] + df["annual_misc_inr"]).values / 1e5
depr_v  = df["annual_depreciation_inr"].values / 1e5
rev_v   = df["annual_revenue_inr"].values / 1e5
b1 = ax1.bar(x, fuel_v, color="#e74c3c", label="Fuel", width=0.5)
b2 = ax1.bar(x, sal_v,  bottom=fuel_v, color="#3498db", label="Salaries", width=0.5)
b3 = ax1.bar(x, maint_v, bottom=fuel_v+sal_v, color="#f39c12", label="Maint & Misc", width=0.5)
b4 = ax1.bar(x, depr_v,  bottom=fuel_v+sal_v+maint_v, color="#9b59b6", label="Depreciation", width=0.5)
for i, rv in enumerate(rev_v):
    ax1.plot([i-0.3, i+0.3], [rv, rv], color=C_GREEN, linewidth=2.5, zorder=5)
    ax1.text(i, rv+0.5, f"₹{rv:.1f}L\nRevenue", ha="center", fontsize=7.5, color=C_GREEN, fontweight="bold")
ax1.set_xticks(x); ax1.set_xticklabels(route_short, fontsize=8.5)
ax1.set_ylabel("₹ Lakhs"); ax1.legend(fontsize=7.5, loc="upper right")
styled_ax(ax1, "Cost Stack vs Revenue")

# Panel 2: Contribution Margin & Net P/L
ax2 = fig.add_subplot(gs[0, 1])
cm_v  = df["contribution_margin_inr"].values / 1e5
net_v = df["net_profit_loss_inr"].values / 1e5
w = 0.32
bars_cm  = ax2.bar(x - w/2, cm_v,  width=w, color=[C_GREEN if v>=0 else C_RED for v in cm_v],  label="Contrib. Margin", alpha=0.85)
bars_net = ax2.bar(x + w/2, net_v, width=w, color=[C_TEAL if v>=0 else "#8B0000" for v in net_v], label="Net P/L", alpha=0.85)
ax2.axhline(0, color="#555", linewidth=1, linestyle="--")
for bar in list(bars_cm) + list(bars_net):
    h = bar.get_height()
    ax2.text(bar.get_x()+bar.get_width()/2, h + (0.3 if h >= 0 else -1.5),
             f"₹{h:.1f}L", ha="center", fontsize=7.5, color="#333")
ax2.set_xticks(x); ax2.set_xticklabels(route_short, fontsize=8.5)
ax2.set_ylabel("₹ Lakhs"); ax2.legend(fontsize=8)
styled_ax(ax2, "Contribution Margin & Net P/L")

# Panel 3: Break-even gauge
ax3 = fig.add_subplot(gs[0, 2])
be_pct  = df["be_occupancy_pct"].values
act_pct = df["occupancy_pct"].values
w = 0.32
ax3.bar(x - w/2, be_pct,  width=w, color=C_RED,   label="BE Occupancy %", alpha=0.85)
ax3.bar(x + w/2, act_pct, width=w, color=C_GREEN, label="Actual Occupancy %", alpha=0.85)
for i in range(len(x)):
    ax3.text(i-w/2, be_pct[i]+1,  f"{be_pct[i]:.0f}%",  ha="center", fontsize=8.5, color=C_RED,   fontweight="bold")
    ax3.text(i+w/2, act_pct[i]+1, f"{act_pct[i]:.0f}%", ha="center", fontsize=8.5, color=C_GREEN, fontweight="bold")
ax3.set_xticks(x); ax3.set_xticklabels(route_short, fontsize=8.5)
ax3.set_ylabel("Occupancy %"); ax3.set_ylim(0, 120); ax3.legend(fontsize=8)
styled_ax(ax3, "Break-even vs Actual Occupancy")

# Panel 4: Sensitivity curves
ax4 = fig.add_subplot(gs[1, 0])
occ_range = np.linspace(10, 100, 300)
colors_sens = [C_TEAL, C_ORANGE]
for idx, (_, row) in enumerate(df.iterrows()):
    nets = []
    for occ in occ_range:
        seats = (occ/100) * row["total_seats"]
        rev   = seats * row["ticket_price_inr"] * row["trips_per_day"] * 2 * row["operating_days_per_year"]
        nets.append((rev - row["annual_total_cost_inr"]) / 1e5)
    ax4.plot(occ_range, nets, label=route_short[idx], color=colors_sens[idx], linewidth=2.5)
    # fill
    nets_arr = np.array(nets)
    ax4.fill_between(occ_range, nets_arr, 0, where=(nets_arr >= 0), alpha=0.12, color=colors_sens[idx])
    ax4.fill_between(occ_range, nets_arr, 0, where=(nets_arr < 0),  alpha=0.08, color=C_RED)
    be = row["be_occupancy_pct"]
    ax4.axvline(be, color=colors_sens[idx], linewidth=1.2, linestyle="--", alpha=0.7)
    ax4.text(be+0.5, -4, f"BE {be:.0f}%", fontsize=7.5, color=colors_sens[idx])
ax4.axhline(0, color="#555", linewidth=1)
ax4.set_xlabel("Occupancy %"); ax4.set_ylabel("Net P/L (₹ Lakhs)"); ax4.legend(fontsize=8)
styled_ax(ax4, "Sensitivity: Occupancy % vs Net P/L")

# Panel 5: Scenario comparison
ax5 = fig.add_subplot(gs[1, 1])
scen_keys  = ["A: Direct Ops (₹/yr)", "B: Lease Out (₹/yr)", "C: PPP 25% Share (₹/yr)", "D: Scrap (₹ one-time)"]
scen_short = ["A: Direct\nOps", "B: Lease\nOut", "C: PPP\n25%", "D: Scrap"]
scen_clrs  = [C_RED, C_GREEN, C_TEAL, C_ORANGE]
x_s = np.arange(4)
w_s = 0.3
for idx, (rid, sdata) in enumerate(scenarios_out.items()):
    vals = [sdata[k]/1e5 for k in scen_keys]
    offset = (idx - 0.5) * w_s
    bars = ax5.bar(x_s + offset, vals, width=w_s, color=scen_clrs,
                   alpha=0.85 if idx==0 else 0.6,
                   label=f"{rid}", edgecolor="white")
ax5.axhline(0, color="#555", linewidth=1, linestyle="--")
ax5.set_xticks(x_s); ax5.set_xticklabels(scen_short, fontsize=8.5)
ax5.set_ylabel("₹ Lakhs"); ax5.legend(fontsize=8)
styled_ax(ax5, "Policy Scenario Comparison")

# Panel 6: Asset valuation waterfall
ax6 = fig.add_subplot(gs[1, 2])
val_labels = ["Acquisition\nCost", "Book Val\n(Yr 1)", "Lease\nIncome/yr", "PPP Share\n/yr", "Scrap\nValue"]
val_values = [
    df["bus_cost_inr"].mean()/1e5,
    df["book_value_inr"].mean()/1e5,
    df["lease_annual_inr"].mean()/1e5,
    df["ppp_annual_share_inr"].mean()/1e5,
    df["scrap_value_inr"].mean()/1e5,
]
bar_clrs = [C_NAVY, C_TEAL, C_GREEN, C_ORANGE, C_RED]
bars6 = ax6.bar(val_labels, val_values, color=bar_clrs, width=0.52, edgecolor="white")
for bar, val in zip(bars6, val_values):
    ax6.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.3,
             f"₹{val:.1f}L", ha="center", fontsize=8.5, fontweight="bold", color="#222")
ax6.set_ylabel("₹ Lakhs (per bus avg)")
styled_ax(ax6, "Asset Valuation Snapshot (per bus)")

plt.suptitle("MSRLM Tourist Bus Fleet — Financial Dashboard  |  Zilla Parishad, Dharashiv",
             fontsize=14, color=C_NAVY, fontweight="bold", y=1.01)
plt.savefig("outputs/financial_dashboard.png", dpi=150, bbox_inches="tight", facecolor="white")
plt.close()
print("  ✅ outputs/financial_dashboard.png")

# ─────────────────────────────────────────────────────────────────────────────
# STAGE 12: EXPORT OUTPUTS
# ─────────────────────────────────────────────────────────────────────────────
print()
print("=" * 70)
print("STAGE 12 — EXPORTING OUTPUTS")
print("=" * 70)

# Export columns
export_cols = [
    "route_id","route_name","bus_id","route_length_km","total_seats",
    "ticket_price_inr","fuel_cost_per_km_inr","avg_occupancy_seats","occupancy_pct",
    "annual_km","annual_revenue_inr","annual_fuel_cost_inr",
    "annual_driver_salary_inr","annual_conductor_salary_inr",
    "annual_maintenance_inr","annual_misc_inr","annual_depreciation_inr",
    "annual_operating_cost_inr","annual_total_cost_inr",
    "contribution_margin_inr","net_profit_loss_inr",
    "monthly_revenue_inr","monthly_cost_inr","monthly_net_inr",
    "be_occupancy_seats","be_occupancy_pct","be_trips_per_day",
    "gap_seats","revenue_shortfall_inr",
    "book_value_inr","lease_annual_inr","ppp_annual_share_inr",
    "scrap_value_inr","recoverable_value_inr",
    "revenue_per_km_inr","cost_per_km_inr",
]

# Sensitivity table as DataFrame
sens_rows = []
for occ in range(10, 101, 5):
    row_d = {"occupancy_pct": occ}
    for _, row in df.iterrows():
        seats = (occ/100) * row["total_seats"]
        rev   = seats * row["ticket_price_inr"] * row["trips_per_day"] * 2 * row["operating_days_per_year"]
        row_d[f"{row['route_id']}_net_inr"] = round(rev - row["annual_total_cost_inr"], 0)
    sens_rows.append(row_d)
sens_df = pd.DataFrame(sens_rows)

# Scenario DataFrame
scen_rows = []
for rid, sdata in scenarios_out.items():
    scen_rows.append({"route_id": rid, **sdata})
scen_export_df = pd.DataFrame(scen_rows)

# Insights DataFrame
insights_df = pd.DataFrame({"insight_no": range(1, len(insights)+1), "insight": insights})

with pd.ExcelWriter("outputs/MSRLM_Financial_Model_v2.xlsx", engine="openpyxl") as writer:
    df[export_cols].to_excel(writer, sheet_name="Route_Model",         index=False)
    sens_df.to_excel(        writer, sheet_name="Sensitivity_Table",   index=False)
    scen_export_df.to_excel( writer, sheet_name="Scenario_Comparison", index=False)
    insights_df.to_excel(    writer, sheet_name="Key_Insights",        index=False)
print("  ✅ outputs/MSRLM_Financial_Model_v2.xlsx")

df[export_cols].to_csv("outputs/clean_route_dataset_v2.csv", index=False)
print("  ✅ outputs/clean_route_dataset_v2.csv")

sens_df.to_csv("outputs/sensitivity_table_v2.csv", index=False)
print("  ✅ outputs/sensitivity_table_v2.csv")

# ─── FINAL SUMMARY ────────────────────────────────────────────────────────────
print(f"""
{'='*70}
FINAL SUMMARY — FLEET POLICY BRIEF
{'='*70}

  Fleet        : 2 × 17-seat MSRLM Tourist Buses | Dharashiv ZP
  Ticket Price : ₹1,450 / seat | Fuel: ₹90/km
  Routes       : 2 circuits (215–216 km, 180 op. days/yr)

  ─── Annual Financials ────────────────────────────────────────
  Combined Revenue        : ₹{fleet_rev:>12,.0f}
  Combined Operating Cost : ₹{fleet_op_cost:>12,.0f}
  Combined Total Cost     : ₹{fleet_tot_cost:>12,.0f}
  Fleet Net Profit/Loss   : ₹{fleet_net:>12,.0f}

  ─── Breakeven ────────────────────────────────────────────────
  R001 BE Occupancy : {df.iloc[0]['be_occupancy_pct']:.1f}% | Actual: {df.iloc[0]['occupancy_pct']:.1f}%
  R002 BE Occupancy : {df.iloc[1]['be_occupancy_pct']:.1f}% | Actual: {df.iloc[1]['occupancy_pct']:.1f}%

  ─── Asset Value ──────────────────────────────────────────────
  Fleet Book Value (Yr 1) : ₹{fleet_book:>12,.0f}
  Annual Lease Income     : ₹{fleet_lease_pa:>12,.0f}
  Annual PPP Share (25%)  : ₹{fleet_ppp_pa:>12,.0f}
  Fleet Scrap Value       : ₹{fleet_scrap:>12,.0f}

  ─── Outputs Generated ────────────────────────────────────────
  • outputs/financial_dashboard.png
  • outputs/MSRLM_Financial_Model_v2.xlsx  (4 sheets)
  • outputs/clean_route_dataset_v2.csv
  • outputs/sensitivity_table_v2.csv
{'='*70}
  MODEL RUN COMPLETE
{'='*70}
""")
