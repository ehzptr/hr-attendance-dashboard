# Performance & Streamlit Architecture Refactoring Plan

This plan addresses all items from the performance and architecture review: vectorizing the attendance engine in Polars, optimizing Excel workbook generation, enforcing strict type boundaries, modularizing Streamlit components, and adding `@st.cache_data` caching for snappy UI interactions.

---

## User Review Required

> [!IMPORTANT]
> - **Vectorized Attendance Engine**: We will eliminate all row-by-row Python loops (`iter_rows`) in [engine.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/engine.py), replacing them with high-performance native Polars expressions. All edge cases (single-scan noon cutoff, grace period, shift overrides, multi-scan flags, leave overrides) will produce identical outputs.
> - **Streamlit Modularization & Caching**: [app.py](file:///f:/hr-dashboard-refactor-1/hrdash/app.py) will be broken into clean modular functions (`render_sidebar_config`, `process_pipeline`, `render_overview_tab`, etc.) with `@st.cache_data` on file ingestion and pipeline calculations.
> - **Typing & Boundaries**: Deprecate defensive `_ensure_polars()` runtime coercions in favor of explicit `pl.DataFrame` types across all modules.

---

## Proposed Changes

### 1. Vectorized Attendance Engine ([engine.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/engine.py))

#### [MODIFY] [engine.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/engine.py)
- **First/Last Scan Extraction**: Replace the loop over `scans_grouped.iter_rows()` with pure Polars expressions:
  - Extract `first_scan_raw = pl.col("Date/Time").first()` and `last_scan_raw = pl.col("Date/Time").last()`.
  - For single scans (`Scan_Count == 1`), compare scan time against cutoff (12:00 / `cfg.single_scan_cutoff`):
    - If `scan < cutoff`: `First_Scan = scan`, `Last_Scan = None`.
    - If `scan >= cutoff`: `First_Scan = None`, `Last_Scan = scan`.
  - For multi scans (`Scan_Count > 1`): `First_Scan = first_scan_raw`, `Last_Scan = last_scan_raw`.
- **Vectorized Shift Schedule Construction**:
  - Build `schedule_in` timestamp expression per row using `Jam_Masuk_Master` (fallback to `cfg.clock_in`).
  - Build `grace_until = schedule_in + timedelta(minutes=cfg.grace_minutes)`.
  - Build `schedule_out` timestamp expression using `Jam_Pulang_Master` (fallback to Saturday vs weekday clock out).
- **Vectorized Evaluation Expressions**:
  - `Status_Masuk`: `when(Is_Approved_Leave).then("Cuti Disetujui").when(~Is_Working_Day & (Scan_Count == 0)).then("Libur").when(~Is_Working_Day).then("Scan Hari Libur").when(Scan_Count == 0).then("Mangkir").when(First_Scan.is_null()).then("Lupa Absen Masuk").when(First_Scan > grace_until).then("Terlambat").otherwise("Hadir")`.
  - `Status_Pulang`: `when(Is_Approved_Leave).then(Leave_Type).when(~Is_Working_Day & (Scan_Count == 0)).then("Libur Normal").when(~Is_Working_Day).then("Scan Hari Libur").when(Scan_Count == 0).then("Mangkir").when(Last_Scan.is_null()).then("Lupa Absen Pulang").when(Last_Scan < schedule_out).then("Pulang Cepat").otherwise("OK")`.
  - `Menit_Telat`: `when(Status_Masuk == "Terlambat").then(((First_Scan - schedule_in).dt.total_seconds() / 60).cast(pl.Int64)).when(Status_Masuk == "Lupa Absen Masuk").then(cfg.default_late_minutes_when_only_clock_out).otherwise(0)`.
  - `Menit_Pulang_Cepat`: `when(Status_Pulang == "Pulang Cepat").then(((schedule_out - Last_Scan).dt.total_seconds() / 60).cast(pl.Int64)).otherwise(0)`.
  - `Anomaly_Type` & `Severity`: Vectorize with `pl.concat_list` or chained `pl.when().then()`.
- **Remove `_ensure_polars`** and use strict `pl.DataFrame` signature.

---

### 2. Type Boundaries & Code Cleanup

#### [MODIFY] [analytics.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/analytics.py)
#### [MODIFY] [payroll.py](file:///f:/hr-dashboard-refactor-1/hrdash/payroll.py)
#### [MODIFY] [sales.py](file:///f:/hr-dashboard-refactor-1/hrdash/sales.py)
#### [MODIFY] [employee_master.py](file:///f:/hr-dashboard-refactor-1/hrdash/employee_master.py)
#### [MODIFY] [parser.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/parser.py)
- Replace defensive `_ensure_polars(df)` helper with strict `pl.DataFrame` type hints and direct operations.

---

### 3. Excel Export Performance

#### [MODIFY] [excel_report.py](file:///f:/hr-dashboard-refactor-1/hrdash/reporting/excel_report.py)
- Use Polars bulk row extraction (`df.iter_rows(named=False)`) with batch cell assignments or direct worksheet table dumping.
- Remove redundant type checks and streamline styling passes.

---

### 4. Streamlit Modular Architecture & Caching

#### [MODIFY] [app.py](file:///f:/hr-dashboard-refactor-1/hrdash/app.py)
- Break monolithic `main()` into clean, single-responsibility functions:
  - `render_sidebar_config() -> tuple[AppConfig, UploadedFiles, str]`
  - `@st.cache_data` cached ingestion: `load_attendance_pipeline(attendance_bytes, master_bytes, leave_bytes, sales_bytes, config_dict)`
  - `render_kpi_overview_tab(...)`
  - `render_employee_recap_tab(...)`
  - `render_department_recap_tab(...)`
  - `render_anomalies_tab(...)`
  - `render_payroll_tab(...)`
  - `render_sales_tab(...)`
  - `render_export_buttons(...)`
- Store pipeline results in session state / cached functions so switching tabs and tweaking UI filters will render immediately with zero latency.

---

## Verification Plan

### Automated Verification
- Run [tests/test_pipeline.py](file:///f:/hr-dashboard-refactor-1/tests/test_pipeline.py) logic to ensure 100% equivalence on:
  - Single-scan afternoon vs morning logic
  - Triple-scan days
  - Mangkir / Sunday off scans
  - Late minutes & early minutes calculations
  - Multi-sheet leave overrides & payroll deductions
- Benchmark runtime performance on the 50-employee dataset generated by `dummy_data.py`.

### Manual Verification
- Test Streamlit web app UI responsiveness across all tabs.
- Verify download of generated Excel and PDF reports.
