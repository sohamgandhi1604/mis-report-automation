# MIS Report Automation Tool

Automates MIS reporting — converts any raw sales CSV into a formatted, 
multi-sheet Excel report with charts using Python.

## Problem It Solves
MIS analysts spend hours manually cleaning data, building pivot tables, 
and formatting Excel reports. This tool does it in one click.

## What It Does
- Cleans and standardises raw CSV data automatically
- Generates a formatted Excel report with 6 sheets:
  - Executive Summary (KPIs)
  - Cleaned Data
  - Monthly Revenue with bar chart
  - Top 10 Customers
  - Top 10 Products
  - Region-Wise Performance

## Tech Stack
- Python 3.x
- pandas — data cleaning and analysis
- openpyxl — Excel report generation
- tkinter — GUI

## How To Run
```bash
pip install pandas openpyxl
python main.py
```
1. Select any sales CSV file
2. Select output folder
3. Click Generate — Excel report saved instantly

## Project Structure
```
├── main.py              # GUI entry point
├── data_cleaner.py      # Data cleaning logic
├── analyzer.py          # Aggregations and KPIs
├── report_generator.py  # Excel report builder
└── requirements.txt
```
