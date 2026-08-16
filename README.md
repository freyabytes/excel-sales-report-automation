# Automated Excel Sales Report Generator

## Project Overview

This project automates the process of transforming raw supermarket sales data into a structured monthly Excel report.

The Python program reads transaction data from an Excel workbook, filters the records according to a month selected by the user, creates a pivot-table-style summary, calculates totals using Excel formulas, applies formatting and adds a chart. The completed report is then exported as a new timestamped Excel file.

I built this project to develop my Python automation and data-processing skills while learning how repetitive spreadsheet-reporting tasks can be made faster, more consistent and less dependent on manual work.

## Key Features

- Reads sales data from an Excel workbook
- Accepts a month as user input
- Converts the `Date` column into a usable date format
- Filters transactions for the selected month
- Groups sales by gender and product line
- Aggregates total sales using a pivot table
- Creates Excel formulas for column totals
- Applies headers, borders, number formatting and column sizing
- Generates a column chart showing sales by product line
- Exports each report with the month and generation timestamp
- Checks for missing files, columns and monthly records
- Can be packaged as a standalone executable using PyInstaller

## Technologies Used

- Python
- pandas
- XlsxWriter
- openpyxl
- PyInstaller
- Microsoft Excel
- Git and GitHub

## Project Structure

```text
excel-sales-report-automation/
├── py-to-exe.py
├── py-to-exe.spec
├── supermarket_sales.xlsx
├── requirements.txt
├── README.md
├── .gitignore
└── examples/
    └── example_march_sales_report.xlsx
```

## How the Automation Works

### 1. Locate the source workbook

The program searches for `supermarket_sales.xlsx` in the script directory, current working directory and relevant parent directory. If the source workbook cannot be found, the program displays an error.

### 2. Read and validate the data

The Excel workbook is loaded into a pandas DataFrame. The program checks that the dataset contains the required columns:

- `Date`
- `Gender`
- `Product line`
- `Total`

This validation prevents the program from continuing with an incompatible dataset.

### 3. Filter by month

The user enters a month such as `January`, `February` or `March`. The program converts the `Date` column into pandas datetime values and filters the records using the selected month.

If no month is entered, the program uses February as the default value. If there are no transactions for the selected month, the program displays an error.

### 4. Create the sales summary

The filtered data is transformed into a pivot table using:

- Gender as the row index
- Product line as the columns
- Total as the value
- Sum as the aggregation method

This produces a concise comparison of sales across customer groups and product categories.

### 5. Generate the Excel report

XlsxWriter creates a new workbook containing a worksheet named `Report`. The program writes the title, selected month, table headings and aggregated values into the worksheet.

### 6. Apply formulas and formatting

The report includes Excel `SUM` formulas for product-line totals. It also applies:

- Arial fonts
- Bold titles and headings
- Borders
- Number formatting
- Background colours
- Adjusted column widths

These changes improve readability and create a more professional report.

### 7. Add the chart

A column chart is generated to compare sales by product line. The chart uses the pivot-table results and is automatically inserted into the report.

### 8. Export the completed report

The final workbook is saved using the month and current timestamp, for example:

```text
report_March_20260816_170110.xlsx
```

Timestamped filenames prevent an existing report from being overwritten and make different reporting runs easier to track.

## Installation

Clone the repository:

```bash
git clone https://github.com/freyabytes/excel-sales-report-automation.git
```

Move into the project folder:

```bash
cd excel-sales-report-automation
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it in Git Bash on Windows:

```bash
source .venv/Scripts/activate
```

Install the required packages:

```bash
pip install -r requirements.txt
```

## How to Run the Program

Make sure `supermarket_sales.xlsx` is in the project folder.

Run:

```bash
python py-to-exe.py
```

Enter a month when prompted:

```text
Introduce month: March
```

The completed Excel report will be created in the project folder.

## Creating a Standalone Executable

The program can be packaged as a Windows executable using PyInstaller:

```bash
pyinstaller py-to-exe.spec
```

The executable will be created inside:

```text
dist/
```

This allows the reporting tool to be used on a Windows computer without manually opening the Python source code.

## Example Output

The generated report contains:

- A monthly sales-report heading
- Sales grouped by gender and product line
- Automatically calculated totals
- Consistent number and cell formatting
- A chart comparing product-line sales

An example output workbook is available in the `examples` folder.

## What I Did

I first divided the reporting workflow into smaller Python scripts. These scripts separately created the pivot table, inserted the chart, applied Excel formulas and formatted the final report.

After understanding each part of the process, I combined the workflow into one main program. I added month-based filtering, timestamped filenames, data validation, file-location handling and error messages. I then prepared the program so it could be converted into a standalone executable.

This development process helped me move from completing individual spreadsheet tasks to designing an end-to-end reporting workflow.

## What I Learned

Through this project, I learned how to:

- Read and manipulate Excel data using pandas
- Convert text-based dates into datetime values
- Filter records according to user input
- Create pivot tables using multiple categories
- Aggregate transaction values programmatically
- Generate Excel formulas from Python
- Format worksheets using XlsxWriter and openpyxl
- Create charts dynamically from processed data
- Validate input data and handle common errors
- Generate unique output filenames using timestamps
- Work with file paths in normal and packaged Python programs
- Package Python code as a Windows executable
- Structure and document a project for GitHub
- Use Git for version control

## Challenges and Solutions

### Handling dates

The source data needed to be filtered by calendar month. I converted the `Date` column using `pandas.to_datetime()` before extracting the month name.

### Avoiding overwritten reports

Saving every result under the same filename would replace the previous report. I added a timestamp to each output filename so multiple reports could be retained.

### Supporting executable builds

File locations may behave differently after a Python program is packaged. I added logic that checks the executable directory, script directory and current working directory for the source workbook.

### Improving spreadsheet readability

Raw output may contain narrow columns or poorly displayed numbers. I used number formats, borders, headings and fixed column widths to create a cleaner report.

### Handling invalid input

The program checks for missing files, required columns, invalid dates and months with no available records. Clear error messages help users identify the problem.

## Skills Demonstrated

- Python programming
- Data cleaning and transformation
- Spreadsheet automation
- Pivot-table analysis
- Report generation
- Data visualisation
- Error handling
- File-path management
- Process automation
- Version control
- Technical documentation

## Potential Improvements

Future versions of the project could include:

- A graphical interface for selecting the month and input file
- Automatic detection of all available months
- Additional filters for branch, city, customer type and payment method
- Comparison between monthly sales periods
- More charts and performance indicators
- Automatic PDF export
- Email distribution of completed reports
- Configurable output folders
- Automated testing
- A dashboard covering revenue, gross income and customer ratings

## Data

The project uses a supermarket sales dataset containing transaction details such as date, gender, product line and transaction total.

The dataset is used for educational and portfolio purposes. Users should confirm the original dataset licence and attribution requirements before redistributing it.

## Author

**Eaint Chue Chue Swe**

GitHub: [freyabytes](https://github.com/freyabytes)

## Acknowledgements

This project was developed as a practical learning exercise in Python and Excel automation. Documentation for pandas, openpyxl, XlsxWriter and PyInstaller was used to understand the relevant data-processing, spreadsheet-generation and application-packaging techniques.
