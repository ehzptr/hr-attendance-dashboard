"""hrdash: HR Attendance, Payroll & Sales Performance Dashboard.

Modular, configuration-driven backend for a car-dealership HR system.

Modules
-------
config              Application configuration (rules, thresholds, payroll rates).
employee_master     Employee master data loading / validation / template.
attendance.parser   Raw attendance-machine log ingestion & cleaning.
attendance.engine   Daily attendance calculation engine (first/last scan rules).
attendance.analytics Employee / department / trend KPI aggregation.
payroll             Payroll deduction engine (configurable, no hard-coded rates).
sales               Sales activity ingestion & sales-performance engine.
reporting.excel_report  Multi-sheet professional Excel workbook generator.
app                 Streamlit UI entrypoint tying every module together.
"""

__version__ = "2.0.0"
