import os
import sys
from datetime import datetime
from pathlib import Path
import subprocess

import pandas as pd

# Attempt to import xlsxwriter; install if missing
try:
    import xlsxwriter
except ImportError:
    print("Installing xlsxwriter...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "xlsxwriter", "-q"])
    import xlsxwriter


if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))


def find_source_file():
    candidates = [
        os.path.join(base_dir, 'supermarket_sales.xlsx'),
        os.path.join(os.getcwd(), 'supermarket_sales.xlsx'),
    ]

    parent_dir = str(Path(base_dir).parent)
    if parent_dir not in {os.getcwd(), base_dir}:
        candidates.append(os.path.join(parent_dir, 'supermarket_sales.xlsx'))

    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate

    return None


def get_monthly_data(month):
    source_file = find_source_file()
    if source_file is None:
        raise FileNotFoundError("supermarket_sales.xlsx is missing from the project or current folder.")

    df = pd.read_excel(source_file)
    if 'Date' not in df.columns:
        raise KeyError("The source file does not contain a 'Date' column.")

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df[df['Date'].dt.strftime('%B').str.lower() == month.lower()].copy()

    if df.empty:
        raise ValueError(f"No sales data found for month: {month}")

    if {'Gender', 'Product line', 'Total'} <= set(df.columns):
        df = df[['Gender', 'Product line', 'Total']]
        pivot = df.pivot_table(
            index='Gender',
            columns='Product line',
            values='Total',
            aggfunc='sum',
        )
        return pivot

    raise KeyError("The source file must contain 'Gender', 'Product line', and 'Total' columns.")


try:
    month = input('Introduce month: ').strip()
    if not month:
        month = 'february'

    pivot = get_monthly_data(month)
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(base_dir, f'report_{month}_{stamp}.xlsx')

    # Create workbook with xlsxwriter (generates clean, valid XML)
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('Report')

    # Formats
    title_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 20,
        'bold': True,
    })
    month_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'bold': True,
    })
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9D9D9',
        'border': 1,
    })
    data_format = workbook.add_format({
        'num_format': '#,##0.00',
        'border': 1,
    })
    total_format = workbook.add_format({
        'num_format': '#,##0.00',
        'bold': True,
        'bg_color': '#E7E6E6',
        'border': 1,
    })

    # Headers
    worksheet.write('A1', 'Sales Report', title_format)
    worksheet.write('A2', month, month_format)

    product_lines = list(pivot.columns)
    genders = list(pivot.index)

    # Column headers
    worksheet.write(4, 0, 'Gender', header_format)
    for col_index, product in enumerate(product_lines, start=1):
        worksheet.write(4, col_index, product, header_format)

    # Data rows
    for gender_index, gender in enumerate(genders, start=5):
        worksheet.write(gender_index, 0, gender, data_format)
        for col_index, product in enumerate(product_lines, start=1):
            value = float(pivot.loc[gender, product])
            worksheet.write(gender_index, col_index, value, data_format)

    # Total row
    total_row = len(genders) + 5
    worksheet.write(total_row, 0, 'Total', total_format)
    for col_index in range(1, len(product_lines) + 1):
        worksheet.write_formula(
            total_row, col_index,
            f'=SUM({chr(64+col_index)}5:{chr(64+col_index)}{len(genders)+4})',
            total_format
        )

    # Create column chart (vertical bars)
    chart = workbook.add_chart({'type': 'column'})
    chart.set_title({'name': 'Sales by Product line'})

    # Add data series for each product line
    for col_index, product in enumerate(product_lines, start=1):
        chart.add_series({
            'name': product,
            'categories': f'=Report!$A$6:$A${len(genders)+5}',
            'values': f'=Report!${chr(64+col_index)}$6:${chr(64+col_index)}${len(genders)+5}',
        })

    chart.set_plotarea({'layout': {'x': 0.13, 'y': 0.13, 'width': 0.75, 'height': 0.75}})
    chart.set_size({'width': 720, 'height': 480})
    worksheet.insert_chart('B12', chart)

    # Auto-fit columns to prevent ### display
    for col_index in range(len(product_lines) + 1):
        worksheet.set_column(col_index, col_index, 18)

    workbook.close()
    print(f'Created: {output_path}')
    
    try:
        if sys.stdin.isatty():
            input('Press Enter to close...')
    except EOFError:
        pass

except Exception as exc:
    print(f'ERROR: {exc}')
    try:
        if sys.stdin.isatty():
            input('Press Enter to close...')
    except EOFError:
        pass