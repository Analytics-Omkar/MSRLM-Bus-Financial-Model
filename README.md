# 🚌 MSRLM Tourist Bus Fleet — Financial Model (v2)
### Zilla Parishad, Dharashiv District, Maharashtra
**17-Seater | ₹1,450/seat | ₹90/km Fuel | 2 Routes**

---

## 📌 Project Purpose

This Python project builds a **complete financial model and interactive Streamlit dashboard** for two MSRLM-funded 17-seater tourist mini-buses operated by Zilla Parishad, Dharashiv. It covers route-level revenue modelling, operating cost analysis, break-even scenarios, asset valuation, policy scenario simulation, and a white paper PDF.

**CV Statement this project supports:**
> *Built Python-based financial model for Zilla Parishad's rural tourism bus fleet, integrating route-level revenue, operating cost, and break-even scenarios; analysis projected ₹1.5+ crore in recoverable asset value and was adopted as the basis for fleet monetisation strategy by the CEO, Zilla Parishad.*

---

## 🗂 Project Structure

```
project/
│
├── data/
│   └── routes_data.csv              ← Input dataset (edit to update assumptions)
│
├── outputs/
│   ├── financial_dashboard.png      ← 6-panel visual dashboard
│   ├── MSRLM_Financial_Model_v2.xlsx← Excel (4 sheets)
│   ├── clean_route_dataset_v2.csv   ← Processed dataset
│   ├── sensitivity_table_v2.csv     ← Sensitivity analysis
│   └── MSRLM_Fleet_Policy_Brief_v2.pdf ← 9-page white paper
│
├── financial_model.py               ← Full 12-stage model (run this first)
├── streamlit_app.py                 ← Interactive dashboard (upload to Streamlit)
├── generate_whitepaper.py           ← PDF white paper generator
├── requirements.txt                 ← Python dependencies
└── README.md                        ← This file
```

---

## ⚙️ Setup & Run

### Step 1 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 2 — Run the financial model
```bash
python financial_model.py
```
Generates: dashboard chart, Excel workbook, and CSVs in `outputs/`

### Step 3 — Generate the white paper
```bash
python generate_whitepaper.py
```
Generates: `outputs/MSRLM_Fleet_Policy_Brief_v2.pdf`

### Step 4 — Run the Streamlit dashboard
```bash
streamlit run streamlit_app.py
```
Opens interactive dashboard in your browser at `http://localhost:8501`

---

## 📊 Streamlit Dashboard — Pages

| Page | Content |
|---|---|
| 📊 Dashboard | KPI cards, cost vs revenue chart, contribution margin chart, key insights |
| 💰 Financial Model | Route-level P&L, detailed cost table, monthly cashflow |
| 📈 Break-even Analysis | Sensitivity curves, break-even table, colour-coded sensitivity grid |
| 🏦 Asset Valuation | Book value table, depreciation-over-time chart |
| 🔮 Scenario Simulation | 4-scenario comparison table + chart + policy recommendations |
| 📋 Data Table | Full computed dataset + CSV download |

**Sidebar controls (live-updating):**
- Upload your own `routes_data.csv`
- Override ticket price, fuel cost, occupancy per route
- Adjust bus age, lease %, PPP share %
- All charts and tables update in real time

---

## 📈 Key Model Outputs (Updated Parameters)

| Metric | R001: Tuljapur | R002: Naldurg | Fleet Total |
|---|---|---|---|
| Annual Revenue | ₹88.45L | ₹75.17L | ₹163.62L |
| Annual Total Cost | ₹145.86L | ₹146.51L | ₹292.37L |
| Net Profit/Loss | ₹−57.41L | ₹−71.34L | ₹−128.75L |

> **Note:** At 17 seats max and ₹1,450 ticket, actual occupancy in the CSV (28 & 24 seats) exceeds capacity. The model clamps occupancy to 17 seats. Ensure the `avg_occupancy_seats` field in your CSV is ≤ 17 for accurate results.

### Break-even at 17 seats / ₹1,450 ticket:
- Break-even = ~14 of 17 seats (82% load)
- At full occupancy (17 seats): **₹31.6L/yr net per route**

---

## 🔧 Updating the Model

To change any assumption:
1. Edit `data/routes_data.csv`  — or use the Streamlit sidebar sliders
2. Re-run `python financial_model.py`

**Important fields to update:**
- `avg_occupancy_seats` — must be ≤ `total_seats` (17)
- `ticket_price_inr` — currently ₹1,450
- `fuel_cost_per_km_inr` — currently ₹90

---

## 📦 Requirements

```
pandas>=1.5
openpyxl>=3.0
matplotlib>=3.5
numpy>=1.21
reportlab>=3.6
streamlit>=1.25
```

---

## 🚀 Deploying to Streamlit Cloud

1. Push all files to a GitHub repository
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file: `streamlit_app.py`
5. Done — your dashboard is live!

**Required files for Streamlit deployment:**
- `streamlit_app.py`
- `data/routes_data.csv`
- `requirements.txt`

---

## 📋 Model Methodology

### Financial Model
```
Revenue       = Occupancy (clamped to 17) × ₹1,450 × 2 trips/day × 2 (round) × 180 days
Fuel Cost     = Route km × 2 × 2 trips × 180 days × ₹90/km
Depreciation  = (₹27,69,534 − ₹7,00,000) ÷ 15 years = ₹1,37,969/yr
Total Cost    = Fuel + Salaries + Maintenance + Misc + Depreciation
Net P/L       = Revenue − Total Cost
Break-even    = Total Cost ÷ (₹1,450 × 4 trips × 180 days)
```

### Asset Valuation
| Method | Formula |
|---|---|
| Book Value | Cost − (Annual Dep × Age) |
| Lease Income | 35% × Annual Revenue |
| PPP Share | 25% × Annual Revenue |
| Scrap Value | 85% × ₹7,00,000 = ₹5,95,000 |

### Scenario Simulation
| Scenario | ZP Net = |
|---|---|
| A: Direct Ops | Net Profit/Loss |
| B: Lease | Lease Income − Depreciation |
| C: PPP | PPP Revenue Share − Depreciation |
| D: Scrap | Scrap Value (one-time) |

---

## 👤 Author

**[Your Name]**  
Policy & Finance Analysis Unit  
Zilla Parishad / MSRLM, Dharashiv District, Maharashtra

**Skills demonstrated:**
`Python` · `pandas` · `matplotlib` · `reportlab` · `Streamlit` · `Financial Modelling` · `Break-even Analysis` · `Asset Valuation` · `Public Policy Analysis` · `Data Visualisation`
