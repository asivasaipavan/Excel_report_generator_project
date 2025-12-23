# Excel Report Generator Project

## Purpose
This project generates an Excel report from a CSV of sales data. The generated `report.xlsx` contains:
- **Raw Data:** A copy of the original sales records.
- **Summary Statistics:** Basic stats (count, total, average, min, max) of the Sales column.
- **Pivot Table:** Total sales aggregated by product category.
- **Bar Chart:** Visual chart of total sales by category embedded in the Excel file.

## Features
- Reads sales data from `sales_data.csv`.
- Computes and writes summary statistics to an Excel sheet.
- Creates a pivot table (by category) and writes it to another sheet.
- Embeds a bar chart of sales by category in the Excel report.
- Uses **pandas** for data processing and **openpyxl** for Excel output and charting.

## Installation
1. Clone or download the repository.
2. (Optional) Create a virtual environment:
   ```bash
   python3 -m venv env
   source env/bin/activate      # On Windows use `env\\Scripts\\activate`
