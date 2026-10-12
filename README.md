# ☕ Cafe Sales Data Analysis

Exploratory data analysis pipeline for a real-world dirty cafe sales dataset (10,000 rows, 2023).  
The pipeline cleans the raw CSV, produces 8 publication-quality charts, prints summary tables, and generates a professional Word report.

---

## Project Structure

```
Cafe_Sales_Data_Analysis/
│
├── data/
│   └── dirty_cafe_sales.csv        # Raw input data (10,000 rows)
│
├── src/
│   ├── __init__.py
│   ├── data_cleaner.py             # Step 1 – clean & impute raw CSV
│   ├── eda.py                      # Step 2 – EDA, charts, summary tables
│   └── report_generator.py         # Step 3 – build Word report
│
├── outputs/                        # Auto-generated (do not edit manually)
│   ├── cleaned_cafe_sales.csv
│   ├── cleaning_log.txt
│   ├── Cafe_Sales_Report.docx
│   └── charts/
│       ├── 01_revenue_by_item.png
│       ├── 02_monthly_revenue.png
│       ├── 03_sales_volume_by_item.png
│       ├── 04_payment_methods.png
│       ├── 05_location_breakdown.png
│       ├── 06_day_of_week_revenue.png
│       ├── 07_heatmap_item_month.png
│       └── 08_quarterly_revenue.png
│
├── main.py                         # Pipeline entry point
├── requirements.txt
└── README.md
```

---

## Prerequisites

- **Python 3.10 or later** — download from [python.org](https://www.python.org/downloads/)  
  ✅ Tick **"Add Python to PATH"** during installation.

---

## Setup

### 1. Clone / open the project folder

```powershell
cd "C:\Users\DELL\Desktop\Cafe_Sales_Data_Analysis"
```

### 2. Create a virtual environment

```powershell
python -m venv venv
```

### 3. Activate the virtual environment

```powershell
# Windows PowerShell
venv\Scripts\Activate.ps1

# Windows CMD
venv\Scripts\activate.bat
```

### 4. Install dependencies

```powershell
pip install --only-binary :all: -r requirements.txt
```

---

## Run the Pipeline

```powershell
python main.py
```

This single command runs all three steps in sequence:

| Step | Module | Output |
|------|--------|--------|
| 1 | `src/data_cleaner.py` | `outputs/cleaned_cafe_sales.csv`, `outputs/cleaning_log.txt` |
| 2 | `src/eda.py` | `outputs/charts/*.png` + console summary tables |
| 3 | `src/report_generator.py` | `outputs/Cafe_Sales_Report.docx` |

---

## Run Individual Modules

```powershell
# Data cleaning only
python -m src.data_cleaner

# EDA + charts only (requires cleaned CSV to exist first)
python -m src.eda

# Word report only (requires cleaned CSV and charts to exist first)
python -m src.report_generator
```

---

## Dataset — Data Quality Issues Found

| Issue Type | Column | Count |
|------------|--------|-------|
| Sentinel "ERROR" / "UNKNOWN" | Item, Quantity, Price Per Unit, Total Spent, Payment Method, Location, Transaction Date | ~3,700 cells |
| Missing (blank) | Location | 3,265 |
| Missing (blank) | Payment Method | 2,579 |
| Missing (blank) | Item, Quantity, Price Per Unit, Total Spent, Transaction Date | ~980 cells |
| **Rows with ≥ 1 issue** | — | **6,911 / 10,000** |

After cleaning: **9,469 usable rows** (531 dropped, 5.3%).

---

## Key Results

| Metric | Value |
|--------|-------|
| Total Transactions | 9,469 |
| Date Range | 2023-01-01 → 2023-12-31 |
| Total Revenue | £84,448.50 |
| Avg Transaction Value | £8.92 |
| Total Units Sold | 28,590 |
| Top Revenue Item | Salad (£18,200) |
| Peak Month | June (£7,334) |

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| pandas | 3.0.6 | Data loading, cleaning, aggregation |
| matplotlib | 3.11.2 | Chart rendering |
| seaborn | 0.13.2 | Statistical heatmaps & palettes |
| python-docx | 1.2.0 | Word report generation |
| openpyxl | 3.1.5 | Excel read/write support |
| numpy | 2.5.4 | Numerical operations |

---

## Deliverables Checklist

- [x] Python source code (`src/*.py`, `main.py`)
- [x] `requirements.txt`
- [x] `README.md`
- [x] Word report (`outputs/Cafe_Sales_Report.docx`)
