# HR Attendance, Payroll & Sales Performance Dashboard

Production refactor of the original single-file Streamlit script into a
modular, configuration-driven application for a car-dealership HR team.

## Run it

```bash
pip install -r <(python -c "import tomllib;print('\n'.join(tomllib.load(open('pyproject.toml','rb'))['project']['dependencies']))")
uv run streamlit run streamlit_app.py
streamlit run streamlit_app.py
```

(or just `python run_hr_dashboard.py` / the packaged Windows `.exe` — see
`README_PYCRUCIBLE_WINDOWS.md`.)

## Architecture

```
RAW ATTENDANCE (mesin absensi)
        |
hrdash/attendance/parser.py      -> validasi & cleaning (RAW_LOG)
        |
hrdash/employee_master.py        -> merge dengan Employee Master (Name/Dept/Type)
        |
hrdash/attendance/engine.py      -> DAILY ATTENDANCE ENGINE
        |                            (first scan = masuk, last scan = pulang)
        v
hrdash/attendance/analytics.py   -> Employee/Department summary + trend
hrdash/payroll.py                -> Potongan gaji (hanya jika tarif diatur di CONFIG)

SALES ACTIVITY (terpisah, non-fingerprint)
        |
hrdash/sales.py                  -> SALES PERFORMANCE ENGINE
        v
hrdash/reporting/excel_report.py -> Workbook 11-sheet (Dashboard, Config, ...)
hrdash/app.py                    -> Streamlit UI (presentation layer only)
```

| Module | Responsibility |
|---|---|
| `hrdash/config.py` | Every configurable business rule (schedule, tolerance, score thresholds, payroll rates, sales-score weights). Round-trips to/from the workbook's `CONFIG` sheet. |
| `hrdash/employee_master.py` | Employee Master loading/validation, `OFFICE`/`SALES` type detection, template generator. |
| `hrdash/attendance/parser.py` | Raw machine-log validation & cleaning. Never mutates/deletes raw rows — rejects go to a separate audit table. |
| `hrdash/attendance/engine.py` | Daily attendance calculation. First scan = clock-in, last scan = clock-out, always — a rule that is never inferred or silently changed. |
| `hrdash/attendance/analytics.py` | Employee/department KPIs, daily trend, HR insights. |
| `hrdash/payroll.py` | Deduction calculation. A rate that HR hasn't configured stays blank ("Belum Diatur") — never assumed. |
| `hrdash/sales.py` | Sales Activity ingestion + Sales Performance scoring. Reads only the Sales Activity file — never the fingerprint log — matching the rule that canvassing is not observable on the attendance machine. |
| `hrdash/reporting/excel_report.py` | Builds the professional 11-sheet workbook (`DASHBOARD`, `CONFIG`, `EMPLOYEE_MASTER`, `RAW_LOG`, `DAILY_ATTENDANCE`, `EMPLOYEE_SUMMARY`, `ANOMALY`, `PAYROLL`, `SALES_PERFORMANCE`, `Sales_Activity`, `README`), including native Excel KPI cards and charts. |
| `hrdash/app.py` | Streamlit UI only — uploads, sidebar configuration, tabs, exports. No business logic lives here. |
| `tests/test_pipeline.py` | End-to-end regression test reproducing the worked examples from the spec (single scan, multi-scan, absence, holiday scan, payroll with/without a configured rate, sales performance independent of attendance). |

## Business rules preserved / made explicit

- **First scan = masuk, last scan = pulang.** Extra scans in between are
  preserved for audit and flagged as `Anomaly_Type`, never interpreted as a
  break, canvassing trip, or re-entry.
- **Sales canvassing is not in the fingerprint log.** `hrdash/sales.py`
  never reads scan counts or scan times; sales KPIs come exclusively from
  the Sales Activity upload, joined to attendance only at the reporting
  layer via Employee ID.
- **No invented payroll rates.** `PayrollConfig` fields default to `None`;
  `hrdash/payroll.py` leaves a deduction column blank rather than assuming
  a rate HR has not supplied.
- **Configurable, not hard-coded.** Schedule, tolerance, holiday calendar,
  score classification thresholds, payroll rates, and sales-score weights
  all live in `hrdash/config.py` and round-trip through the workbook's
  `CONFIG` sheet — HR can change them without touching code.
- **Raw data is never deleted.** Unparseable/invalid rows are captured in a
  `_reject_reason`-tagged table instead of being dropped.

## What's new vs. the original script

- Split into a proper package (`hrdash/`) instead of one ~1000-line file.
- Employee Master module (Section 9) with `OFFICE`/`SALES` type detection.
- Payroll engine (Section 7) with HR-configurable, never-assumed rates.
- Sales Activity ingestion + Sales Performance engine (Sections 3-4),
  intentionally decoupled from the attendance pipeline.
- `CONFIG` sheet round-trip so all rules are visible/editable in Excel.
- Severity taxonomy (`HIGH` / `LOW` / blank) and `Anomaly_Type` categories
  matching Section 5 exactly.
- Native-Excel `DASHBOARD` sheet (KPI cards, priority tables, line/bar
  charts) in addition to the interactive Streamlit dashboard.
- A regression test (`tests/test_pipeline.py`) that exercises the exact
  scenarios described in the spec.

## Things intentionally left configurable rather than decided for you

- Deduction rate per late minute / early-leave minute / absence day.
- Attendance-score classification thresholds (default 90/80/70).
- Sales performance scoring weights (default: Visit 15%, Test Drive 15%,
  SPK 30%, Delivery 25%, Conversion Rate 15%).
- Severity assigned to a holiday scan (default `LOW`) — flagged for review,
  not automatically treated as misconduct.

If any of these differ from your company's actual policy, change them in
the sidebar or in the exported `CONFIG` sheet — no code changes needed.
