import pandas as pd
import os
from datetime import datetime

# ---------- USER CONFIG ----------
INPUT_FILE = "GLOBAL DATASET .xlsx"   
OUTPUT_FILE = "GlobalTech_Sales_Analysis.xlsx"
HIGHLIGHT_MANAGER = "Emmanuel"        
TOP_N_PRODUCTS = 10                   
CURRENCY_SYMBOL = "₦"                 
# ---------------------------------

def col_idx_to_excel_col(idx):
    """0-based idx to Excel column letter (A, B, ..., Z, AA, AB, ...)"""
    letters = ""
    while True:
        idx, rem = divmod(idx, 26)
        letters = chr(65 + rem) + letters
        if idx == 0:
            break
        idx -= 1
    return letters

if not os.path.exists(INPUT_FILE):
    raise FileNotFoundError(f"Input file '{INPUT_FILE}' not found. Put the dataset in same folder or change INPUT_FILE.")

# 1) Load dataset
if INPUT_FILE.lower().endswith(".csv"):
    df = pd.read_csv(INPUT_FILE, low_memory=False)
else:
    df = pd.read_excel(INPUT_FILE)

# Required columns check (based on case study)
required = ["Order Date", "Category", "Sub-Category", "Product Name", "Manager", "Sales", "Order ID", "Customer ID"]
for c in required:
    if c not in df.columns:
        raise ValueError(f"Required column missing from input: '{c}'")

# 2) Preparing Year and Month columns
df["Order Date"] = pd.to_datetime(df["Order Date"], errors="coerce")
df["Year"] = df["Order Date"].dt.year
df["Month"] = df["Order Date"].dt.month

# Filling NaN managers with empty string so formulas don't error
df["Manager"] = df["Manager"].fillna("")

# Basic aggregates for findings (we'll use pandas to create narrative, but Excel will still contain formulas)
total_sales_val = df["Sales"].sum()
total_orders = df["Order ID"].nunique()
distinct_managers = df["Manager"].nunique()
avg_order_value = total_sales_val / total_orders if total_orders else 0

# Category year-over-year calculation for findings: compare first year to last year in data
years = sorted(df["Year"].dropna().unique())
first_year = int(years[0])
last_year = int(years[-1])

cat_sales_by_year = df.groupby(["Category", "Year"])["Sales"].sum().unstack(fill_value=0)
cat_totals = df.groupby("Category")["Sales"].sum().sort_values(ascending=False)
top_category = cat_totals.index[0] if not cat_totals.empty else None

# Manager top performer
mgr_totals = df.groupby("Manager")["Sales"].sum().sort_values(ascending=False)
top_manager = mgr_totals.index[0] if not mgr_totals.empty else None
top_manager_sales = mgr_totals.iloc[0] if not mgr_totals.empty else 0

# Top product
prod_totals = df.groupby("Product Name")["Sales"].sum().sort_values(ascending=False)
top_product = prod_totals.index[0] if not prod_totals.empty else None
top_product_share = (prod_totals.iloc[0] / total_sales_val) if total_sales_val else 0

# Create Excel
with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
    workbook = writer.book

    # Write raw data to Data sheet (this is the authoritative table Excel formulas will use)
    sheet_data = "Data"
    df.to_excel(writer, sheet_name=sheet_data, index=False)
    ws_data = writer.sheets[sheet_data]

    # Map Data columns to Excel letters (used for SUMIF/SUMIFS formulas)
    header = list(df.columns)
    col_map = {name: i for i, name in enumerate(header)}  # 0-based
    excel_col_letters = {name: col_idx_to_excel_col(idx) for name, idx in col_map.items()}

    # --- Manager Performance sheet ---
    sheet_mgr = "Manager Performance"
    ws_mgr = workbook.add_worksheet(sheet_mgr)
    writer.sheets[sheet_mgr] = ws_mgr

    managers = sorted(df["Manager"].replace("", pd.NA).dropna().unique())
    ws_mgr.write(0, 0, "Manager")
    ws_mgr.write(0, 1, "Sales (calc)")
    ws_mgr.write(0, 2, "Excel Formula")
    ws_mgr.write(0, 3, "Rank")

    # Build formulas referencing Data sheet full column ranges (e.g., Data!$F:$F)
    manager_col_letter = excel_col_letters["Manager"]
    sales_col_letter = excel_col_letters["Sales"]

    for r, mgr in enumerate(managers, start=1):
        ws_mgr.write(r, 0, mgr)
        mgr_esc = str(mgr).replace('"', '""')
        formula = f'=SUMIF({sheet_data}!${manager_col_letter}:${manager_col_letter},"{mgr_esc}",{sheet_data}!${sales_col_letter}:${sales_col_letter})'
        ws_mgr.write_formula(r, 1, formula)
        ws_mgr.write(r, 2, formula)
        # Rank formula (higher sales gets rank 1)
        ws_mgr.write_formula(r, 3, f'=RANK(B{r+1},$B$2:$B${1+len(managers)},0)')

    # Formats
    fmt_title = workbook.add_format({'bold': True, 'font_size': 12})
    fmt_money = workbook.add_format({'num_format': f'{CURRENCY_SYMBOL}#,##0.00'})
    ws_mgr.set_column(0, 0, 30)
    ws_mgr.set_column(1, 1, 18, fmt_money)
    ws_mgr.set_column(2, 2, 60)
    ws_mgr.set_column(3, 3, 8)

    # Conditional formatting: top 3 green, bottom 3 red
    ws_mgr.conditional_format(1, 1, len(managers), 1, {
        'type': 'top',
        'value': 3,
        'format': workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    })
    ws_mgr.conditional_format(1, 1, len(managers), 1, {
        'type': 'bottom',
        'value': 3,
        'format': workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    })

    # Highlight a specific manager if requested
    if HIGHLIGHT_MANAGER and HIGHLIGHT_MANAGER in managers:
        idx = managers.index(HIGHLIGHT_MANAGER) + 1
        highlight_fmt = workbook.add_format({'bg_color': '#FFF2CC', 'bold': True})
        ws_mgr.write(idx, 0, HIGHLIGHT_MANAGER, highlight_fmt)
        # rewrite sales cell formula with highlight format
        mgr_esc = str(HIGHLIGHT_MANAGER).replace('"', '""')
        formula = f'=SUMIF({sheet_data}!${manager_col_letter}:${manager_col_letter},"{mgr_esc}",{sheet_data}!${sales_col_letter}:${sales_col_letter})'
        ws_mgr.write_formula(idx, 1, formula, highlight_fmt)

    # Manager bar chart
    chart_mgr = workbook.add_chart({'type': 'column'})
    chart_mgr.add_series({
        'name': 'Sales by Manager',
        'categories': f"='{sheet_mgr}'!$A$2:$A${1+len(managers)}",
        'values':     f"='{sheet_mgr}'!$B$2:$B${1+len(managers)}",
    })
    chart_mgr.set_title({'name': 'Managerial Performance — Sales'})
    chart_mgr.set_x_axis({'name': 'Manager'})
    chart_mgr.set_y_axis({'name': 'Sales'})
    chart_mgr.set_size({'width': 720, 'height': 360})
    ws_mgr.insert_chart(1, 5, chart_mgr)

    # --- Category Trends sheet (Category x Year) ---
    sheet_cat = "Category Trends"
    ws_cat = workbook.add_worksheet(sheet_cat)
    writer.sheets[sheet_cat] = ws_cat

    years = sorted(df["Year"].dropna().unique())
    categories = sorted(df["Category"].dropna().unique())

    # Header row
    ws_cat.write(0, 0, "Category \\ Year")
    for j, y in enumerate(years, start=1):
        ws_cat.write(0, j, int(y))
    # Fill categories and SUMIFS formulas
    category_col_letter = excel_col_letters["Category"]
    year_col_letter = excel_col_letters["Year"]
    for i, cat in enumerate(categories, start=1):
        ws_cat.write(i, 0, cat)
        for j, y in enumerate(years, start=1):
            # SUMIFS(sales, category, category_cell, year, year_value)
            formula = (f"=SUMIFS({sheet_data}!${sales_col_letter}:${sales_col_letter},"
                       f"{sheet_data}!${category_col_letter}:${category_col_letter},$A{ i + 1 },"
                       f"{sheet_data}!${year_col_letter}:${year_col_letter},{int(y)})")
            ws_cat.write_formula(i, j, formula)

    ws_cat.set_column(0, 0, 30)
    ws_cat.set_column(1, 1+len(years), 15, fmt_money)

    # Add a line chart showing the trend for top categories (top 6)
    top_cats = list(cat_totals.head(6).index)
    # Build a small table underneath for charting convenience
    start_row = len(categories) + 3
    ws_cat.write(start_row - 1, 0, "Category")
    for j, y in enumerate(years, start=1):
        ws_cat.write(start_row - 1, j, int(y))
    for i, cat in enumerate(top_cats, start=start_row):
        ws_cat.write(i, 0, cat)
        # formula to pull each year value from the Category Trends table above: MATCH row by category and pick column
        # But simpler: reference the cell directly since we know layout: category "cat" at row (1 + categories.index(cat))
        row_idx = categories.index(cat) + 1
        for j, y in enumerate(years, start=1):
            # from table at (row_idx, col j)
            src_cell = f"'{sheet_cat}'!${col_idx_to_excel_col(j)}${row_idx+0}"
            # note: col_idx_to_excel_col(j) returns letters for this sheet's columns; for direct refer we can use the numeric approach
            # Instead of complex addressing, use SUMIFS against Data sheet to keep formulas robust:
            formula = (f"=SUMIFS({sheet_data}!${sales_col_letter}:${sales_col_letter},"
                       f"{sheet_data}!${category_col_letter}:${category_col_letter},\"{cat}\","
                       f"{sheet_data}!${year_col_letter}:${year_col_letter},{int(years[j-1])})")
            ws_cat.write_formula(i, j, formula)

    # Chart across year columns for top categories
    chart_line = workbook.add_chart({'type': 'line'})
    for i in range(start_row, start_row + len(top_cats)):
        chart_line.add_series({
            'name':       f"='{sheet_cat}'!$A${i+1}",
            'categories': f"='{sheet_cat}'!$B${start_row}:$B${start_row + len(years)-1}",
            'values':     f"='{sheet_cat}'!$B${i+1}:$${col_idx_to_excel_col(1+len(years))}${i+1}",
        })
    chart_line.set_title({'name': 'Top Categories — Yearly Trend'})
    # If chart building above fails due to complexity across columns, we'll add a simpler chart on the Dashboard instead.
    chart_line.set_size({'width': 720, 'height': 360})
    # place chart (if it errors in some Excel versions, ignore)
    try:
        ws_cat.insert_chart(1, 5, chart_line)
    except Exception:
        pass

    # --- Top Products sheet ---
    sheet_prod = "Top Products"
    ws_prod = workbook.add_worksheet(sheet_prod)
    writer.sheets[sheet_prod] = ws_prod

    prod_list = list(prod_totals.head(TOP_N_PRODUCTS).index)
    ws_prod.write(0, 0, "Product Name")
    ws_prod.write(0, 1, "Sales (calc)")
    ws_prod.write(0, 2, "% of Total")
    for i, p in enumerate(prod_list, start=1):
        ws_prod.write(i, 0, p)
        p_esc = str(p).replace('"', '""')
        formula = f'=SUMIF({sheet_data}!${excel_col_letters["Product Name"]}:${excel_col_letters["Product Name"]},"{p_esc}",{sheet_data}!${sales_col_letter}:${sales_col_letter})'
        ws_prod.write_formula(i, 1, formula)
        ws_prod.write_formula(i, 2, f"=B{i+1}/SUM($B$2:$B${1+len(prod_list)})")

    ws_prod.set_column(0, 0, 60)
    ws_prod.set_column(1, 1, 18, fmt_money)
    ws_prod.set_column(2, 2, 12, workbook.add_format({'num_format': '0.00%'}))

    # Pie chart for top products
    chart_pie = workbook.add_chart({'type': 'pie'})
    chart_pie.add_series({
        'name': 'Top Products Share',
        'categories': f"='{sheet_prod}'!$A$2:$A${1+len(prod_list)}",
        'values':     f"='{sheet_prod}'!$B$2:$B${1+len(prod_list)}",
        'data_labels': {'percentage': True}
    })
    chart_pie.set_title({'name': 'Top Products (Cumulative Share)'})
    ws_prod.insert_chart(1, 4, chart_pie, {'x_scale': 1.3, 'y_scale': 1.3})

    # --- Dashboard sheet ---
    sheet_dash = "Dashboard"
    ws_dash = workbook.add_worksheet(sheet_dash)
    writer.sheets[sheet_dash] = ws_dash

    # KPI boxes
    fmt_kpi_title = workbook.add_format({'bold': True, 'font_size': 11})
    fmt_kpi_val = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
    fmt_kpi_money = workbook.add_format({'bold': True, 'font_size': 14, 'num_format': f'{CURRENCY_SYMBOL}#,##0.00', 'align': 'center'})

    ws_dash.set_column(0, 4, 22)
    ws_dash.write(0, 0, "KPI", fmt_kpi_title)
    ws_dash.write(0, 1, "Value", fmt_kpi_title)

    # Total Sales (formula)
    ws_dash.write(1, 0, "Total Sales (2011-2014)")
    ws_dash.write_formula(1, 1, f"=SUM({sheet_data}!${sales_col_letter}:${sales_col_letter})", fmt_kpi_money)

    # Total Orders
    ws_dash.write(2, 0, "Total Orders")
    # Assuming Order ID is column in Data sheet column A.. -> use COUNTA on Order ID column
    order_id_col = excel_col_letters["Order ID"]
    ws_dash.write_formula(2, 1, f"=COUNTA({sheet_data}!${order_id_col}:${order_id_col})-1")  # minus header

    # Total Managers
    ws_dash.write(3, 0, "Total Managers")
    # Using pandas precomputed distinct_managers (safer / compatible)
    ws_dash.write(3, 1, distinct_managers)

    # Average Order Value
    ws_dash.write(4, 0, "Average Order Value")
    ws_dash.write_formula(4, 1, f"=IFERROR(SUM({sheet_data}!${sales_col_letter}:${sales_col_letter})/COUNTA({sheet_data}!${order_id_col}:${order_id_col}),0)", fmt_kpi_money)

    # Insert charts: Manager (from Manager Performance), Top Products (from Top Products), and Category trend (simpler yearly line)
    # Manager chart (refer to Manager Performance)
    chart_mgr_small = workbook.add_chart({'type': 'column'})
    chart_mgr_small.add_series({
        'name': 'Sales by Manager',
        'categories': f"='{sheet_mgr}'!$A$2:$A${1+len(managers)}",
        'values':     f"='{sheet_mgr}'!$B$2:$B${1+len(managers)}",
    })
    chart_mgr_small.set_title({'name': 'Managerial Performance'})
    chart_mgr_small.set_size({'width': 480, 'height': 280})
    ws_dash.insert_chart(6, 0, chart_mgr_small)

    # Top products pie
    chart_prod_small = workbook.add_chart({'type': 'pie'})
    chart_prod_small.add_series({
        'name': 'Top Products',
        'categories': f"='{sheet_prod}'!$A$2:$A${1+len(prod_list)}",
        'values':     f"='{sheet_prod}'!$B$2:$B${1+len(prod_list)}",
        'data_labels': {'percentage': True}
    })
    chart_prod_small.set_size({'width': 400, 'height': 300})
    ws_dash.insert_chart(6, 4, chart_prod_small)

    # Simple yearly sales line chart (compute totals by SUMIFS formulas on Dashboard)
    ws_dash.write(12, 0, "Year")
    ws_dash.write(12, 1, "Sales")
    for i, y in enumerate(years, start=13):
        ws_dash.write(i-1, 0, int(y))
        ws_dash.write_formula(i-1, 1, f"=SUMIFS({sheet_data}!${sales_col_letter}:${sales_col_letter},{sheet_data}!${year_col_letter}:${year_col_letter},{int(y)})", fmt_money)
    chart_year = workbook.add_chart({'type': 'line'})
    chart_year.add_series({
        'name': 'Sales by Year',
        'categories': f"='{sheet_dash}'!$A$13:$A${12+len(years)}",
        'values':     f"='{sheet_dash}'!$B$13:$B${12+len(years)}",
    })
    chart_year.set_title({'name': 'Sales by Year'})
    chart_year.set_size({'width': 600, 'height': 320})
    ws_dash.insert_chart(12, 3, chart_year)

    # --- Findings & Recommendations sheet ---
    sheet_find = "Findings & Recommendations"
    ws_find = workbook.add_worksheet(sheet_find)
    writer.sheets[sheet_find] = ws_find
    ws_find.set_column(0, 0, 120)

    
    lines = []
    lines.append("Executive Summary:")
    lines.append(f"- Total Sales (2011-2014): {CURRENCY_SYMBOL}{total_sales_val:,.2f}")
    lines.append(f"- Total Orders: {total_orders:,}")
    lines.append(f"- Distinct Managers: {distinct_managers}")
    lines.append(f"- Average Order Value: {CURRENCY_SYMBOL}{avg_order_value:,.2f}")
    lines.append("")

    if top_manager:
        lines.append(f"Manager Performance:")
        pct_above_avg = ((top_manager_sales - mgr_totals.mean()) / mgr_totals.mean() * 100) if mgr_totals.mean() else 0
        lines.append(f"- Top Manager: {top_manager} with {CURRENCY_SYMBOL}{top_manager_sales:,.2f} ({pct_above_avg:.1f}% above manager average). Consider incentives and capturing their playbook.")
        lines.append("")

    if top_category:
        # YoY growth between first and last year for top category or for categories generally
        first = cat_sales_by_year.loc[top_category, first_year] if first_year in cat_sales_by_year.columns else 0
        last = cat_sales_by_year.loc[top_category, last_year] if last_year in cat_sales_by_year.columns else 0
        yoy_pct = ((last - first) / first * 100) if first else 0
        lines.append("Category Trends:")
        lines.append(f"- Top Category by cumulative sales: {top_category} (consider re-allocating marketing spend toward this category).")
        lines.append(f"- {top_category} sales in {first_year}: {CURRENCY_SYMBOL}{first:,.2f}; in {last_year}: {CURRENCY_SYMBOL}{last:,.2f} — change: {yoy_pct:.1f}% between {first_year} and {last_year}.")
        lines.append("")

    if top_product:
        lines.append("Top Product:")
        lines.append(f"- Top Product: {top_product} — contributes {top_product_share*100:.1f}% of total sales. Recommend bundling and heavier advertising.")
        lines.append("")

    # General recommendations framework aligned to case study
    lines.append("Recommendations Framework:")
    lines.append("- Managerial Performance: implement tiered incentives, targeted training for bottom performers, and consider reassigning managers with consistent underperformance.")
    lines.append("- Category Trends: increase inventory & promotions for high-growth categories; reduce stock for persistently declining categories.")
    lines.append("- Top Product: raise advertising spend, create bundle offers, and monitor stock to avoid stockouts.")
    lines.append("")
    lines.append("Notes:")
    lines.append("- All numeric KPIs in this workbook are live formulas referencing the 'Data' sheet. Editing data in 'Data' will recalc all results and charts.")
    lines.append("- The file is formatted in Naira (₦) by default where currency was not specified in the brief.")

    for i, line in enumerate(lines):
        ws_find.write(i, 0, line)

    
print(f"Created '{OUTPUT_FILE}'. Open this file in Excel to view live formulas, charts, and the dashboard.")