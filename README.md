# Intern Performance KPI Tracker

An automated pipeline that tracks intern performance using key metrics (KPIs) and generates ready-to-share monthly Excel reports for supervisors — built with Python and SQL.

## Project Goal
Replace manual, ad-hoc intern performance tracking with a repeatable, automated system that:
- Defines clear KPIs: **Project Quality**, **Mentor Feedback**, and **Task Completion** (engagement proxy)
- Extracts and loads data automatically (no manual spreadsheet entry)
- Surfaces underperforming interns automatically via a flagging rule
- Outputs a polished, formula-driven Excel report ready for supervisors

## Tech Stack
- **Python** (pandas, openpyxl) — data cleaning, automation, report generation
- **SQL** (SQLite) — KPI aggregation queries
- **Excel** — final supervisor-facing deliverable (with live formulas, not hardcoded numbers)

## Pipeline / How it works
1. `interns_cleaned.csv` — cleaned source data (1,470 records, 6 key fields)
2. `load_data.py` — loads the CSV into a SQLite database (`interns.db`)
3. `kpi_queries.sql` — SQL queries: average KPIs by role, by department, and a list of flagged underperformers
4. `generate_report.py` — builds the final Excel report (`Monthly_Intern_Report.xlsx`) with:
   - Summary by Intern Role
   - Summary by Department
   - Raw data with an automatic "Needs Attention" flag column
   - All values calculated via live Excel formulas, so the report self-updates if data changes

## Sample Output
See `Monthly_Intern_Report_June_2026.xlsx` for a full example report generated from this pipeline.

## Key Finding
During data exploration, `project_quality` showed very limited variation across the dataset (almost all interns scored 3 or 4), making it a weak standalone differentiator — documented here as part of standard data-quality review, a step any solid analytics workflow includes before reporting.

## Run it yourself
```bash
pip install pandas openpyxl
python load_data.py interns_cleaned.csv
python generate_report.py "Your Month Here"
```

## Author
Areeba — Data Analyst | BI & Python/SQL automation | [Fiverr: areebaimtiaz787]
