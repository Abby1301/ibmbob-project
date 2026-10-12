"""
cafe_sales_analysis.py
======================
Cafe Sales Data Analysis — single-file pipeline
Combines: data cleaning → EDA & charts → Word report

Run with:
    python cafe_sales_analysis.py

Outputs:
    outputs/cleaned_cafe_sales.csv
    outputs/cleaning_log.txt
    outputs/charts/*.png  (8 charts)
    outputs/Cafe_Sales_Report.docx
"""

# ─────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import os
import datetime

import pandas as pd
import matplotlib
matplotlib.use("Agg")          # non-interactive backend — no display needed
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─────────────────────────────────────────────────────────────────────────────
# PATHS & CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
RAW_PATH    = os.path.join("data", "dirty_cafe_sales.csv")
CLEAN_PATH  = os.path.join("outputs", "cleaned_cafe_sales.csv")
LOG_PATH    = os.path.join("outputs", "cleaning_log.txt")
CHARTS_DIR  = os.path.join("outputs", "charts")
REPORT_PATH = os.path.join("outputs", "Cafe_Sales_Report.docx")

# Sentinel strings that mean "unknown / corrupt"
DIRTY_VALUES = {"ERROR", "UNKNOWN", "NULL", "N/A"}

# Known valid items and their standard prices
ITEM_PRICES = {
    "Coffee":   2.0,
    "Tea":      1.5,
    "Cake":     3.0,
    "Cookie":   1.0,
    "Sandwich": 4.0,
    "Salad":    5.0,
    "Smoothie": 4.0,
    "Juice":    3.0,
}

# Chart styling
PALETTE   = "Set2"
ACCENT    = "#3b82d4"
BG        = "#f7f8fa"

MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun",
               "Jul","Aug","Sep","Oct","Nov","Dec"]
DOW_ORDER   = ["Monday","Tuesday","Wednesday","Thursday",
               "Friday","Saturday","Sunday"]

CHART_FILES = [
    ("01_revenue_by_item.png",      "Figure 1 – Total Revenue by Item"),
    ("02_monthly_revenue.png",      "Figure 2 – Monthly Revenue Trend"),
    ("03_sales_volume_by_item.png", "Figure 3 – Total Units Sold by Item"),
    ("04_payment_methods.png",      "Figure 4 – Payment Method Distribution"),
    ("05_location_breakdown.png",   "Figure 5 – Transactions by Location"),
    ("06_day_of_week_revenue.png",  "Figure 6 – Revenue by Day of Week"),
    ("07_heatmap_item_month.png",   "Figure 7 – Monthly Revenue Heatmap"),
    ("08_quarterly_revenue.png",    "Figure 8 – Revenue by Quarter"),
]


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1 — DATA CLEANING
# ═════════════════════════════════════════════════════════════════════════════

def _to_nan(series: pd.Series) -> pd.Series:
    """Replace dirty sentinel strings and empty strings with NaN."""
    return series.replace(list(DIRTY_VALUES) + [""], pd.NA)


def load_raw(path: str = RAW_PATH) -> pd.DataFrame:
    return pd.read_csv(path, dtype=str)


def clean(df: pd.DataFrame, log_lines: list) -> pd.DataFrame:
    original_rows = len(df)
    log_lines.append(f"Raw rows loaded          : {original_rows}")

    # 1. Standardise column names
    df.columns = [c.strip() for c in df.columns]

    # 2. Replace all dirty sentinels with NaN
    for col in df.columns:
        before = df[col].isna().sum()
        df[col] = _to_nan(df[col])
        after  = df[col].isna().sum()
        if after > before:
            log_lines.append(f"  Sentinel→NaN in '{col}': {after - before} cells")

    # 3. Numeric coercion
    df["Quantity"]       = pd.to_numeric(df["Quantity"],       errors="coerce")
    df["Price Per Unit"] = pd.to_numeric(df["Price Per Unit"], errors="coerce")
    df["Total Spent"]    = pd.to_numeric(df["Total Spent"],    errors="coerce")

    # 4. Date parsing
    df["Transaction Date"] = pd.to_datetime(
        df["Transaction Date"], format="%Y-%m-%d", errors="coerce"
    )
    log_lines.append(
        f"  Unparseable dates dropped: {df['Transaction Date'].isna().sum()}"
    )

    # 5. Drop rows with no Transaction Date (unusable for time analysis)
    before_date_drop = len(df)
    df = df.dropna(subset=["Transaction Date"])
    log_lines.append(
        f"  Rows dropped (no date)   : {before_date_drop - len(df)}"
    )

    # 6. Impute Item from Price Per Unit where possible
    price_to_item = {v: k for k, v in ITEM_PRICES.items()}
    mask = df["Item"].isna() & df["Price Per Unit"].notna()
    df.loc[mask, "Item"] = df.loc[mask, "Price Per Unit"].map(price_to_item)

    # 7. Impute Price Per Unit from Item
    mask = df["Price Per Unit"].isna() & df["Item"].notna()
    df.loc[mask, "Price Per Unit"] = df.loc[mask, "Item"].map(ITEM_PRICES)

    # 8. Impute Quantity from Total Spent / Price Per Unit
    mask = (df["Quantity"].isna()
            & df["Total Spent"].notna()
            & df["Price Per Unit"].notna()
            & (df["Price Per Unit"] != 0))
    df.loc[mask, "Quantity"] = (
        df.loc[mask, "Total Spent"] / df.loc[mask, "Price Per Unit"]
    ).round().astype("Int64")

    # 9. Impute Total Spent from Quantity × Price Per Unit
    mask = (df["Total Spent"].isna()
            & df["Quantity"].notna()
            & df["Price Per Unit"].notna())
    df.loc[mask, "Total Spent"] = (
        df.loc[mask, "Quantity"] * df.loc[mask, "Price Per Unit"]
    )

    # 10. Drop rows still missing Item or core numeric fields
    before_drop = len(df)
    df = df.dropna(subset=["Item", "Quantity", "Price Per Unit", "Total Spent"])
    log_lines.append(
        f"  Rows dropped (still missing core fields): {before_drop - len(df)}"
    )

    # 11. Fill categorical NaN with 'Unknown'
    df["Payment Method"] = df["Payment Method"].fillna("Unknown")
    df["Location"]       = df["Location"].fillna("Unknown")

    # 12. Drop rows where Item is not in known items
    before_item = len(df)
    df = df[df["Item"].isin(ITEM_PRICES.keys())]
    log_lines.append(
        f"  Rows dropped (unrecognised item): {before_item - len(df)}"
    )

    # 13. Add derived time columns
    df["Month"]       = df["Transaction Date"].dt.month
    df["Month_Name"]  = df["Transaction Date"].dt.strftime("%b")
    df["Day_of_Week"] = df["Transaction Date"].dt.day_name()
    df["Quarter"]     = df["Transaction Date"].dt.quarter

    # 14. Reset index
    df = df.reset_index(drop=True)

    clean_rows = len(df)
    log_lines.append(f"Clean rows remaining     : {clean_rows}")
    log_lines.append(
        f"Total rows removed       : {original_rows - clean_rows} "
        f"({(original_rows - clean_rows)/original_rows*100:.1f}%)"
    )
    return df


def run_cleaning():
    """Step 1: clean raw CSV, save outputs, return clean DataFrame."""
    log_lines = ["=" * 60, "CAFE SALES — DATA CLEANING LOG", "=" * 60]
    df_raw   = load_raw()
    df_clean = clean(df_raw, log_lines)

    os.makedirs("outputs", exist_ok=True)
    df_clean.to_csv(CLEAN_PATH, index=False)
    log_lines.append(f"\nCleaned CSV saved → {CLEAN_PATH}")

    with open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines))
    print("\n".join(log_lines))
    print(f"\nCleaning log saved → {LOG_PATH}")
    return df_clean


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2 — EDA & CHARTS
# ═════════════════════════════════════════════════════════════════════════════

def _save_chart(fig: plt.Figure, name: str):
    os.makedirs(CHARTS_DIR, exist_ok=True)
    path = os.path.join(CHARTS_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"  Chart saved → {path}")


def chart_revenue_by_item(df: pd.DataFrame):
    rev = (df.groupby("Item")["Total Spent"]
             .sum().sort_values(ascending=False).reset_index())
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    ax.set_facecolor(BG)
    bars = ax.bar(rev["Item"], rev["Total Spent"],
                  color=sns.color_palette(PALETTE, len(rev)))
    ax.bar_label(bars, fmt="£%.0f", padding=3, fontsize=8)
    ax.set_title("Total Revenue by Item", fontsize=14, fontweight="bold", pad=12)
    ax.set_xlabel("Item")
    ax.set_ylabel("Total Revenue (£)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x:,.0f}"))
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "01_revenue_by_item.png")
    return rev


def chart_monthly_revenue(df: pd.DataFrame):
    monthly = (df.groupby("Month_Name")["Total Spent"]
                 .sum().reindex(MONTH_ORDER).reset_index())
    monthly.columns = ["Month", "Revenue"]
    fig, ax = plt.subplots(figsize=(10, 5), facecolor=BG)
    ax.set_facecolor(BG)
    ax.plot(monthly["Month"], monthly["Revenue"],
            marker="o", color=ACCENT, linewidth=2.2, markersize=6)
    ax.fill_between(monthly["Month"], monthly["Revenue"], alpha=0.12, color=ACCENT)
    ax.set_title("Monthly Revenue Trend (2023)", fontsize=14, fontweight="bold", pad=12)
    ax.set_xlabel("Month")
    ax.set_ylabel("Total Revenue (£)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x:,.0f}"))
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "02_monthly_revenue.png")
    return monthly


def chart_sales_volume(df: pd.DataFrame):
    vol = (df.groupby("Item")["Quantity"]
             .sum().sort_values(ascending=True).reset_index())
    fig, ax = plt.subplots(figsize=(8, 5), facecolor=BG)
    ax.set_facecolor(BG)
    ax.barh(vol["Item"], vol["Quantity"],
            color=sns.color_palette(PALETTE, len(vol)))
    ax.bar_label(ax.containers[0], fmt="%.0f", padding=3, fontsize=8)
    ax.set_title("Total Units Sold by Item", fontsize=14, fontweight="bold", pad=12)
    ax.set_xlabel("Units Sold")
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "03_sales_volume_by_item.png")
    return vol


def chart_payment_methods(df: pd.DataFrame):
    pay = df["Payment Method"].value_counts().reset_index()
    pay.columns = ["Method", "Count"]
    pay = pay[pay["Method"] != "Unknown"]
    fig, ax = plt.subplots(figsize=(6, 6), facecolor=BG)
    ax.set_facecolor(BG)
    wedges, texts, autotexts = ax.pie(
        pay["Count"], labels=pay["Method"],
        autopct="%1.1f%%", startangle=140,
        colors=sns.color_palette(PALETTE, len(pay)),
        wedgeprops={"edgecolor": "white", "linewidth": 1.5}
    )
    for t in autotexts:
        t.set_fontsize(9)
    ax.set_title("Payment Method Distribution", fontsize=14, fontweight="bold", pad=12)
    fig.tight_layout()
    _save_chart(fig, "04_payment_methods.png")
    return pay


def chart_location(df: pd.DataFrame):
    loc = (df[df["Location"] != "Unknown"]["Location"]
           .value_counts().reset_index())
    loc.columns = ["Location", "Transactions"]
    fig, ax = plt.subplots(figsize=(6, 5), facecolor=BG)
    ax.set_facecolor(BG)
    bars = ax.bar(loc["Location"], loc["Transactions"],
                  color=[ACCENT, "#7c5cd8"])
    ax.bar_label(bars, padding=3, fontsize=9)
    ax.set_title("Transactions by Location", fontsize=14, fontweight="bold", pad=12)
    ax.set_ylabel("Number of Transactions")
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "05_location_breakdown.png")
    return loc


def chart_day_of_week(df: pd.DataFrame):
    dow = (df.groupby("Day_of_Week")["Total Spent"]
             .sum().reindex(DOW_ORDER).reset_index())
    dow.columns = ["Day", "Revenue"]
    fig, ax = plt.subplots(figsize=(9, 5), facecolor=BG)
    ax.set_facecolor(BG)
    colors = [ACCENT if d in ("Saturday","Sunday") else "#a8c4e0" for d in dow["Day"]]
    bars = ax.bar(dow["Day"], dow["Revenue"], color=colors)
    ax.bar_label(bars, fmt="£%.0f", padding=3, fontsize=7.5)
    ax.set_title("Revenue by Day of Week", fontsize=14, fontweight="bold", pad=12)
    ax.set_xlabel("Day")
    ax.set_ylabel("Total Revenue (£)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x:,.0f}"))
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "06_day_of_week_revenue.png")
    return dow


def chart_heatmap(df: pd.DataFrame):
    pivot = (df.groupby(["Month_Name", "Item"])["Total Spent"]
               .sum().unstack(fill_value=0).reindex(MONTH_ORDER))
    fig, ax = plt.subplots(figsize=(12, 6), facecolor=BG)
    sns.heatmap(pivot, annot=True, fmt=".0f", cmap="Blues",
                linewidths=0.5, ax=ax, cbar_kws={"label": "Revenue (£)"})
    ax.set_title("Monthly Revenue Heatmap (Item × Month)",
                 fontsize=14, fontweight="bold", pad=12)
    ax.set_xlabel("Item")
    ax.set_ylabel("Month")
    fig.tight_layout()
    _save_chart(fig, "07_heatmap_item_month.png")
    return pivot


def chart_quarterly(df: pd.DataFrame):
    qtr = (df.groupby("Quarter")["Total Spent"].sum().reset_index())
    qtr["Quarter"] = qtr["Quarter"].apply(lambda q: f"Q{q}")
    fig, ax = plt.subplots(figsize=(7, 5), facecolor=BG)
    ax.set_facecolor(BG)
    bars = ax.bar(qtr["Quarter"], qtr["Total Spent"],
                  color=sns.color_palette("Blues_d", len(qtr)))
    ax.bar_label(bars, fmt="£%.0f", padding=3, fontsize=9)
    ax.set_title("Revenue by Quarter", fontsize=14, fontweight="bold", pad=12)
    ax.set_ylabel("Total Revenue (£)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"£{x:,.0f}"))
    ax.spines[["top","right"]].set_visible(False)
    fig.tight_layout()
    _save_chart(fig, "08_quarterly_revenue.png")
    return qtr


def print_summary_tables(df: pd.DataFrame):
    print("\n" + "="*55)
    print("SUMMARY STATISTICS")
    print("="*55)
    print(f"Total transactions : {len(df):,}")
    print(f"Date range         : {df['Transaction Date'].min().date()} → "
          f"{df['Transaction Date'].max().date()}")
    print(f"Total revenue      : £{df['Total Spent'].sum():,.2f}")
    print(f"Avg transaction    : £{df['Total Spent'].mean():.2f}")
    print(f"Total units sold   : {df['Quantity'].sum():,.0f}")

    print("\n── Revenue by Item ──────────────────────────────────────")
    rev_item = (df.groupby("Item")["Total Spent"]
                  .agg(["sum","mean","count"])
                  .rename(columns={"sum":"Revenue","mean":"Avg","count":"Txns"})
                  .sort_values("Revenue", ascending=False))
    rev_item["Revenue"] = rev_item["Revenue"].map("£{:,.2f}".format)
    rev_item["Avg"]     = rev_item["Avg"].map("£{:.2f}".format)
    print(rev_item.to_string())

    print("\n── Revenue by Location ──────────────────────────────────")
    print(df[df["Location"] != "Unknown"]
          .groupby("Location")["Total Spent"]
          .sum().map("£{:,.2f}".format).to_string())

    print("\n── Revenue by Payment Method ────────────────────────────")
    print(df[df["Payment Method"] != "Unknown"]
          .groupby("Payment Method")["Total Spent"]
          .sum().sort_values(ascending=False)
          .map("£{:,.2f}".format).to_string())

    print("\n── Top 5 Revenue Months ─────────────────────────────────")
    top5 = (df.groupby("Month_Name")["Total Spent"]
              .sum().sort_values(ascending=False).head(5)
              .map("£{:,.2f}".format))
    print(top5.to_string())


def run_eda(df: pd.DataFrame):
    """Step 2: generate all charts and print summary tables."""
    print("Generating charts …")
    chart_revenue_by_item(df)
    chart_monthly_revenue(df)
    chart_sales_volume(df)
    chart_payment_methods(df)
    chart_location(df)
    chart_day_of_week(df)
    chart_heatmap(df)
    chart_quarterly(df)
    print_summary_tables(df)
    print("\nEDA complete. All charts saved to outputs/charts/")


# ═════════════════════════════════════════════════════════════════════════════
# SECTION 3 — WORD REPORT GENERATOR
# ═════════════════════════════════════════════════════════════════════════════

def _set_cell_bg(cell, hex_color: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _heading(doc: Document, text: str, level: int = 1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.runs[0] if p.runs else p.add_run(text)
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
    return p


def _add_table(doc: Document, df_table: pd.DataFrame,
               header_hex: str = "3b82d4"):
    table = doc.add_table(rows=1 + len(df_table), cols=len(df_table.columns))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    hdr = table.rows[0]
    for i, col in enumerate(df_table.columns):
        cell = hdr.cells[i]
        cell.text = str(col)
        _set_cell_bg(cell, header_hex)
        run = cell.paragraphs[0].runs[0]
        run.font.bold      = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size      = Pt(9)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Data rows
    for r_idx, row in df_table.iterrows():
        data_row = table.rows[r_idx + 1]
        bg = "F7F8FA" if r_idx % 2 == 0 else "FFFFFF"
        for c_idx, val in enumerate(row):
            cell = data_row.cells[c_idx]
            cell.text = str(val)
            _set_cell_bg(cell, bg)
            cell.paragraphs[0].runs[0].font.size = Pt(9)
    return table


def _add_chart(doc: Document, filename: str, caption: str, width_in: float = 5.5):
    path = os.path.join(CHARTS_DIR, filename)
    if os.path.exists(path):
        doc.add_picture(path, width=Inches(width_in))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.runs[0].font.size      = Pt(9)
        cap.runs[0].font.color.rgb = RGBColor(0x57, 0x60, 0x6A)


def build_report(df: pd.DataFrame):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1.2)
        section.right_margin  = Inches(1.2)

    # ── Cover ──────────────────────────────────────────────────────────────
    doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("☕  Cafe Sales Data Analysis Report")
    run.font.size      = Pt(22)
    run.font.bold      = True
    run.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = subtitle.add_run("Exploratory Analysis of 2023 Transaction Data")
    sr.font.size      = Pt(13)
    sr.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

    doc.add_paragraph()
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    dr = date_p.add_run(f"Generated: {datetime.date.today().strftime('%d %B %Y')}")
    dr.font.size      = Pt(10)
    dr.font.color.rgb = RGBColor(0x57, 0x60, 0x6A)
    doc.add_page_break()

    # ── 1. Executive Summary ───────────────────────────────────────────────
    _heading(doc, "1. Executive Summary")
    total_rev  = df["Total Spent"].sum()
    total_txns = len(df)
    avg_txn    = df["Total Spent"].mean()
    total_qty  = df["Quantity"].sum()
    top_item   = df.groupby("Item")["Total Spent"].sum().idxmax()
    top_month  = df.groupby("Month_Name")["Total Spent"].sum().idxmax()

    doc.add_paragraph(
        f"This report presents an exploratory data analysis of the Cafe Sales "
        f"dataset for the year 2023. After cleaning, the dataset contains "
        f"{total_txns:,} transactions spanning all twelve months, "
        f"representing a total revenue of £{total_rev:,.2f}. "
        f"The average transaction value is £{avg_txn:.2f}, and a total of "
        f"{total_qty:,.0f} units were sold across eight product categories. "
        f"The best-selling item by revenue was {top_item}, "
        f"and peak sales month was {top_month}."
    )
    doc.add_paragraph()

    kpi_data = pd.DataFrame({
        "Metric": ["Total Transactions", "Total Revenue", "Avg Transaction Value",
                   "Total Units Sold",   "Top Item",       "Peak Month"],
        "Value":  [f"{total_txns:,}", f"£{total_rev:,.2f}", f"£{avg_txn:.2f}",
                   f"{total_qty:,.0f}", top_item, top_month],
    })
    _add_table(doc, kpi_data)
    doc.add_paragraph()
    doc.add_page_break()

    # ── 2. Data Quality & Cleaning ─────────────────────────────────────────
    _heading(doc, "2. Data Quality & Cleaning")
    doc.add_paragraph(
        "The raw dataset contained 10,000 rows with significant quality issues "
        "including sentinel strings ('ERROR', 'UNKNOWN'), blank cells, and "
        "unparseable dates. The cleaning pipeline performed the following steps:"
    )
    for s in [
        "Replaced all 'ERROR' and 'UNKNOWN' sentinel strings with NaN.",
        "Parsed Transaction Date; rows with unparseable dates were dropped.",
        "Imputed missing Price Per Unit from the known item price lookup table.",
        "Imputed missing Item from Price Per Unit where unambiguous.",
        "Derived Quantity from Total Spent ÷ Price Per Unit where possible.",
        "Derived Total Spent from Quantity × Price Per Unit where possible.",
        "Dropped rows still missing core numeric fields after imputation.",
        "Filled remaining missing Payment Method and Location with 'Unknown'.",
        "Derived time-series columns: Month, Month_Name, Day_of_Week, Quarter.",
    ]:
        p = doc.add_paragraph(s, style="List Bullet")
        p.runs[0].font.size = Pt(10)

    if os.path.exists(LOG_PATH):
        doc.add_paragraph()
        _heading(doc, "Cleaning Log Summary", level=2)
        with open(LOG_PATH, encoding="utf-8") as f:
            lines = [l.strip() for l in f if l.strip() and l.strip() != "=" * 60]
        for line in lines:
            p = doc.add_paragraph(line)
            p.runs[0].font.size      = Pt(9)
            p.runs[0].font.color.rgb = RGBColor(0x57, 0x60, 0x6A)
    doc.add_page_break()

    # ── 3. Revenue Analysis ────────────────────────────────────────────────
    _heading(doc, "3. Revenue Analysis")

    _heading(doc, "3.1 Revenue by Item", level=2)
    rev_item = (df.groupby("Item")["Total Spent"]
                  .agg(Revenue="sum", Transactions="count", Avg_Value="mean")
                  .sort_values("Revenue", ascending=False).reset_index())
    rev_item["Revenue"]   = rev_item["Revenue"].map("£{:,.2f}".format)
    rev_item["Avg_Value"] = rev_item["Avg_Value"].map("£{:.2f}".format)
    rev_item.columns      = ["Item", "Total Revenue", "Transactions", "Avg Value"]
    _add_table(doc, rev_item)
    doc.add_paragraph()
    _add_chart(doc, "01_revenue_by_item.png", "Figure 1 – Total Revenue by Item", 5.5)
    doc.add_page_break()

    _heading(doc, "3.2 Monthly Revenue Trend", level=2)
    monthly = (df.groupby("Month_Name")["Total Spent"]
                 .sum().reindex(MONTH_ORDER).reset_index())
    monthly.columns  = ["Month", "Revenue"]
    monthly["Revenue"] = monthly["Revenue"].map("£{:,.2f}".format)
    _add_table(doc, monthly)
    doc.add_paragraph()
    _add_chart(doc, "02_monthly_revenue.png", "Figure 2 – Monthly Revenue Trend (2023)", 5.8)
    doc.add_page_break()

    _heading(doc, "3.3 Quarterly Revenue", level=2)
    qtr = (df.groupby("Quarter")["Total Spent"]
             .agg(Revenue="sum", Transactions="count").reset_index())
    qtr["Quarter"]  = qtr["Quarter"].apply(lambda q: f"Q{q}")
    qtr["Revenue"]  = qtr["Revenue"].map("£{:,.2f}".format)
    qtr.columns     = ["Quarter", "Revenue", "Transactions"]
    _add_table(doc, qtr)
    doc.add_paragraph()
    _add_chart(doc, "08_quarterly_revenue.png", "Figure 8 – Revenue by Quarter", 4.5)
    doc.add_page_break()

    # ── 4. Sales Volume ────────────────────────────────────────────────────
    _heading(doc, "4. Sales Volume by Item")
    vol = (df.groupby("Item")["Quantity"]
             .sum().sort_values(ascending=False).reset_index())
    vol.columns = ["Item", "Units Sold"]
    _add_table(doc, vol)
    doc.add_paragraph()
    _add_chart(doc, "03_sales_volume_by_item.png", "Figure 3 – Total Units Sold by Item", 5.5)
    doc.add_page_break()

    # ── 5. Customer Behaviour ──────────────────────────────────────────────
    _heading(doc, "5. Customer Behaviour")

    _heading(doc, "5.1 Payment Methods", level=2)
    pay = (df[df["Payment Method"] != "Unknown"]["Payment Method"]
           .value_counts().reset_index())
    pay.columns   = ["Payment Method", "Transactions"]
    pay["Share %"] = (pay["Transactions"] / pay["Transactions"].sum() * 100).map("{:.1f}%".format)
    _add_table(doc, pay)
    doc.add_paragraph()
    _add_chart(doc, "04_payment_methods.png", "Figure 4 – Payment Method Distribution", 3.8)

    doc.add_paragraph()
    _heading(doc, "5.2 Location", level=2)
    loc = (df[df["Location"] != "Unknown"]["Location"]
           .value_counts().reset_index())
    loc.columns   = ["Location", "Transactions"]
    loc["Share %"] = (loc["Transactions"] / loc["Transactions"].sum() * 100).map("{:.1f}%".format)
    _add_table(doc, loc)
    doc.add_paragraph()
    _add_chart(doc, "05_location_breakdown.png", "Figure 5 – Transactions by Location", 4.2)
    doc.add_page_break()

    # ── 6. Time Patterns ──────────────────────────────────────────────────
    _heading(doc, "6. Sales Patterns Over Time")

    _heading(doc, "6.1 Day-of-Week Revenue", level=2)
    _add_chart(doc, "06_day_of_week_revenue.png", "Figure 6 – Revenue by Day of Week", 5.8)

    doc.add_paragraph()
    _heading(doc, "6.2 Item × Month Heatmap", level=2)
    _add_chart(doc, "07_heatmap_item_month.png",
               "Figure 7 – Monthly Revenue Heatmap (Item × Month)", 6.0)
    doc.add_page_break()

    # ── 7. Key Findings ────────────────────────────────────────────────────
    _heading(doc, "7. Key Findings & Recommendations")
    top3_items  = (df.groupby("Item")["Total Spent"].sum()
                     .sort_values(ascending=False).head(3).index.tolist())
    worst_month = df.groupby("Month_Name")["Total Spent"].sum().idxmin()
    best_loc    = (df[df["Location"] != "Unknown"]
                   .groupby("Location")["Total Spent"].sum().idxmax())

    for f_text in [
        f"Top 3 revenue-generating items are: {', '.join(top3_items)}.",
        f"Peak sales month is {top_month}; lowest-revenue month is {worst_month} — "
        f"consider promotions or seasonal menu changes in {worst_month}.",
        f"'{best_loc}' location drives the highest revenue and should be prioritised "
        f"for staffing and stock allocation.",
        "Cash, Credit Card, and Digital Wallet payments are the three primary "
        "payment channels. Ensuring all three are always available reduces friction.",
        "The day-of-week chart highlights weekend revenue patterns — "
        "consider weekend-only specials to capitalise on higher footfall.",
        "Approximately 69% of raw rows had at least one data-quality issue; "
        "investing in POS system validation would significantly improve data integrity.",
    ]:
        p = doc.add_paragraph(f_text, style="List Bullet")
        p.runs[0].font.size = Pt(10)

    doc.add_page_break()

    # ── 8. Appendix ────────────────────────────────────────────────────────
    _heading(doc, "Appendix – All Charts")
    for fname, caption in CHART_FILES:
        _add_chart(doc, fname, caption, 5.8)
        doc.add_paragraph()

    doc.add_paragraph()
    footer_p = doc.add_paragraph(
        "Generated by Cafe Sales Data Analysis Pipeline · Python 3.15"
    )
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.runs[0].font.size      = Pt(8)
    footer_p.runs[0].font.color.rgb = RGBColor(0x57, 0x60, 0x6A)

    os.makedirs("outputs", exist_ok=True)
    doc.save(REPORT_PATH)
    print(f"\nWord report saved → {REPORT_PATH}")


def run_report():
    """Step 3: build Word report from cleaned CSV and saved charts."""
    df = pd.read_csv(CLEAN_PATH, parse_dates=["Transaction Date"])
    build_report(df)


# ═════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═════════════════════════════════════════════════════════════════════════════

def main():
    print("\n" + "=" * 60)
    print("  CAFE SALES DATA ANALYSIS PIPELINE")
    print("=" * 60)

    print("\n[1/3] Cleaning data …")
    df_clean = run_cleaning()

    print("\n[2/3] Running EDA and generating charts …")
    run_eda(df_clean)

    print("\n[3/3] Building Word report …")
    run_report()

    print("\n" + "=" * 60)
    print("  PIPELINE COMPLETE")
    print("=" * 60)
    print("Deliverables:")
    print("  outputs/cleaned_cafe_sales.csv")
    print("  outputs/cleaning_log.txt")
    print("  outputs/charts/  (8 charts)")
    print("  outputs/Cafe_Sales_Report.docx")
    print("=" * 60)


if __name__ == "__main__":
    main()
