# 📊 XBRL Financial Statement Crawler (公開資訊觀測站財報爬蟲)

A Python toolkit for automatically scraping XBRL financial statements from the [Taiwan Stock Exchange Market Observation Post System (公開資訊觀測站)](https://mopsov.twse.com.tw), with built-in financial ratio analysis and Excel export.

---

## ✨ Features

- **Three Core Financial Statements** — Balance Sheet (資產負債表), Comprehensive Income Statement (綜合損益表), and Cash Flow Statement (現金流量表)
- **Annual Mode** — Fetch the last N annual periods (e.g. 5 or 7 years)
- **Quarterly Mode** — Fetch the last N quarters as **single-quarter figures** (not YTD cumulative):
  - Balance Sheet: end-of-quarter snapshots
  - Income Statement: single-quarter figures (Q4 = Full Year − Q3 YTD)
  - Cash Flow Statement: single-quarter figures (derived by subtracting prior YTD)
- **Financial Ratio Analysis** — Automatically calculates 25+ ratios across 6 categories:
  - Liquidity & Solvency (流動性與償債能力)
  - Operating Efficiency / Turnover (經營效率週轉率)
  - Profitability (獲利能力)
  - Capital Structure (資本結構)
  - Financial Leverage (財務槓桿)
  - Asset Structure (資產結構)
- **Common-Size Analysis** — Vertical analysis percentages appended alongside raw figures
- **Batch Processing** — Process multiple companies in a single run; individual failures do not abort the batch
- **Excel Export** — One `.xlsx` file per company, with each statement and ratio analysis on a separate sheet
- **Auto-fallback** — Automatically switches between consolidated (合併) and individual (個別) statements if one is unavailable

---

## 📁 Project Structure

```
xbrl_crawler/
├── xbrl/
│   ├── __init__.py         # Package entry point & public API
│   ├── core.py             # HTTP fetching, HTML parsing, data cleaning, ratio calc, Excel export
│   ├── annual.py           # Annual mode: multi-period merging logic
│   ├── quarterly.py        # Quarterly mode: single-quarter derivation logic
│   └── batch.py            # Batch processing for multiple companies
├── xbrl_crawler_5p_FA.py   # Entry script — annual mode
├── xbrl_crawler_8q_FA.py   # Entry script — quarterly mode
├── output_annual/          # Annual Excel outputs (auto-created)
└── output_quarterly/       # Quarterly Excel outputs (auto-created)
```

---

## ⚙️ Requirements

- Python 3.8+
- `pandas`
- `requests`
- `beautifulsoup4`
- `openpyxl`
- `lxml` (recommended HTML parser)

Install all dependencies:

```bash
pip install pandas requests beautifulsoup4 openpyxl lxml
```

---

## 🚀 Quick Start

### Annual Mode — Last N Annual Periods

Edit `xbrl_crawler_5p_FA.py`:

```python
STOCK_IDS = "2330"           # Single stock, or a list: ["2330", "2454"]
YEAR      = 2025             # Latest fiscal year
QUARTER   = "Q4"             # Annual reports use Q4
PERIODS   = 5                # Number of annual periods to fetch
OUTPUT_DIR = "output_annual"
```

Then run:

```bash
python xbrl_crawler_5p_FA.py
```

### Quarterly Mode — Last N Single Quarters

Edit `xbrl_crawler_8q_FA.py`:

```python
STOCK_IDS  = "2330"          # Single stock, or a list
YEAR       = 2026            # Latest quarter's year
QUARTER    = "Q1"            # Latest quarter
N_QUARTERS = 12              # Number of quarters to fetch
OUTPUT_DIR = "output_quarterly"
```

Then run:

```bash
python xbrl_crawler_8q_FA.py
```

---

## 📦 Output

Each company produces one Excel file with the following sheets:

| Sheet | Content |
|---|---|
| `資產負債表` | Balance Sheet with common-size % columns |
| `綜合損益表` | Income Statement with common-size % columns |
| `現金流量表` | Cash Flow Statement |
| `財務比率分析` | 25+ financial ratios across all periods |

**File naming convention:**
- Annual: `{YEAR}_三表分析_{STOCK_ID}_近{N}期.xlsx`
- Quarterly: `{YEAR}{QUARTER}_三表分析_{STOCK_ID}_近{N}季.xlsx`

---

## 🔢 Financial Ratios Calculated

| Category | Ratios |
|---|---|
| Liquidity & Solvency | Current Ratio, Quick Ratio, Working Capital, Interest Coverage, Fixed Charge Coverage |
| Operating Efficiency | Total Asset Turnover, Inventory Turnover & Days, AR Turnover & Days, AP Turnover & Days, Operating Cycle, Cash Turnover, Cash Conversion Cycle (CCC) |
| Profitability | Gross Margin, Operating Margin, Net Margin, EPS, ROA, ROE, Return on Long-term Capital |
| Capital Structure | Debt Ratio, Equity Ratio |
| Financial Leverage | Equity Multiplier, Financial Leverage Index |
| Asset Structure | Long-term Capital to PPE Ratio |

> **Note:** Quarterly ratios use 91 days per period for turnover day calculations; annual ratios use 365 days.

---

## 🛠️ Using as a Library

You can also import individual functions directly:

```python
from xbrl.annual import get_five_periods_statements
from xbrl.core import export_statements_to_excel

# Fetch 5 annual periods for TSMC (2330)
statements = get_five_periods_statements("2330", 2025, "Q4", periods=5)
export_statements_to_excel(statements, "2330_annual.xlsx")
```

```python
from xbrl.quarterly import get_quarterly_statements
from xbrl.core import export_statements_to_excel

# Fetch last 8 single quarters
statements = get_quarterly_statements("2330", 2026, "Q1", n_quarters=8)
export_statements_to_excel(
    statements,
    "2330_quarterly.xlsx",
    days_per_period=91,
    quarterly_labels=True,
)
```

---

## ⚠️ Notes

- Data source: [公開資訊觀測站 XBRL viewer](https://mopsov.twse.com.tw/server-java/t164sb01)
- Only companies that file XBRL reports on MOPS are supported (listed and OTC companies on TWSE/TPEx)
- Cash flow statements for OTC/emerging market companies (興櫃) may be reported as YTD rather than single-quarter
- A small delay (`SLEEP_BETWEEN_STOCKS`) is applied between requests to avoid overloading the server — please use responsibly
- The crawler auto-detects and falls back between consolidated and individual financial statements

---

## 📄 License

MIT License — feel free to use, modify, and distribute.

---

## 🙏 Acknowledgements

Financial data provided by [Taiwan Stock Exchange (TWSE)](https://www.twse.com.tw) via the [公開資訊觀測站](https://mops.twse.com.tw) XBRL disclosure system.
