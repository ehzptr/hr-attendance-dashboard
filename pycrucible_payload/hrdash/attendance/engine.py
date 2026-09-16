"""
Daily attendance calculation engine (Sections 2, 3, 5, 17).

Core rule, unchanged from the original implementation and preserved as the
single source of truth for the whole system:

    * FIRST scan of a calendar day  -> Status Masuk (clock-in)
    * LAST  scan of a calendar day  -> Status Pulang (clock-out)
    * Any scan in between is preserved for audit but is NEVER interpreted
      as a break / canvassing / re-entry — for SALES employees especially,
      canvassing time is simply not observable in the fingerprint log
      (Section 3) and this engine makes no attempt to infer it.

Severity taxonomy follows Section 5 exactly: HIGH, LOW, or blank ("") when
the day is normal. Anomaly_Type follows the taxonomy listed in Section 5.
"""

from __future__ import annotations

from datetime import date, datetime, time
from typing import Iterable, Optional

import numpy as np
import pandas as pd

from ..config import AttendanceConfig
from .parser import DAY_ID


def _combine(day: date, t: time) -> pd.Timestamp:
    return pd.Timestamp(datetime.combine(day, t))


def _parse_time_value(val: object) -> Optional[time]:
    if val is None or pd.isna(val):
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.time()
    s = str(val).strip().replace(".", ":")
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            return time(int(parts[0]), int(parts[1]))
        except ValueError:
            return None
    return None


def build_employee_directory(enriched_log: pd.DataFrame) -> pd.DataFrame:
    """One row per employee, resolving the most frequent Name/Department/
    Employee_Type seen in the log (data-quality safety net when no master
    file was supplied for a given ID).
    """
    rows: list[dict] = []
    for emp_id, group in enriched_log.groupby("No.", dropna=False):
        def _mode(col: str, fallback: object) -> object:
            m = group[col].mode()
            return m.iloc[0] if not m.empty else fallback

        rows.append(
            {
                "No.": emp_id,
                "Name": _mode("Name", group["Name"].iloc[0]),
                "Department": _mode("Department", group["Department"].iloc[0]),
                "Employee_Type": _mode("Employee_Type", group.get("Employee_Type", pd.Series(["OFFICE"])).iloc[0]),
                "Position": _mode("Position", "") if "Position" in group.columns else "",
                "Jam_Masuk_Master": _mode("Jam_Masuk_Master", pd.NA) if "Jam_Masuk_Master" in group.columns else pd.NA,
                "Jam_Pulang_Master": _mode("Jam_Pulang_Master", pd.NA) if "Jam_Pulang_Master" in group.columns else pd.NA,
            }
        )
    return pd.DataFrame(rows).sort_values(["Department", "Name"], kind="stable").reset_index(drop=True)


def _scan_pair(group: pd.DataFrame, cutoff: time) -> pd.Series:
    scans = group["Date/Time"].sort_values().tolist()
    if not scans:
        return pd.Series({"First_Scan": pd.NaT, "Last_Scan": pd.NaT, "Scan_Count": 0})

    if len(scans) == 1:
        scan = scans[0]
        noon = _combine(scan.date(), cutoff)
        if scan < noon:
            return pd.Series({"First_Scan": scan, "Last_Scan": pd.NaT, "Scan_Count": 1})
        return pd.Series({"First_Scan": pd.NaT, "Last_Scan": scan, "Scan_Count": 1})

    return pd.Series({"First_Scan": scans[0], "Last_Scan": scans[-1], "Scan_Count": len(scans)})


def build_daily_attendance(
    enriched_log: pd.DataFrame,
    cfg: AttendanceConfig,
    holidays: Optional[Iterable[date]] = None,
    approved_leave: Optional[pd.DataFrame | object] = None,
    leave_result: Optional[object] = None,
) -> pd.DataFrame:
    """Build the Daily Attendance table (Section 5).

    ``enriched_log`` is the cleaned raw log after being joined with the
    Employee Master (``No., Name, Department, Employee_Type, Position``
    columns present — see ``employee_master.merge_master_into_log``).
    """
    if enriched_log.empty:
        return pd.DataFrame()

    holidays_set = set(holidays or [])
    
    # Support both legacy DataFrame and LeaveLoadResult
    lr = leave_result
    if lr is None and hasattr(approved_leave, "leaves") and hasattr(approved_leave, "group_events"):
        lr = approved_leave
        leave = pd.DataFrame(columns=["No.", "Tanggal", "Leave_Type"])
    elif isinstance(approved_leave, pd.DataFrame):
        leave = approved_leave
    else:
        leave = pd.DataFrame(columns=["No.", "Tanggal", "Leave_Type"])

    start = cfg.start_date or enriched_log["Tanggal"].min()
    end = cfg.end_date or enriched_log["Tanggal"].max()

    employees = build_employee_directory(enriched_log)
    dates = pd.DataFrame({"Tanggal": pd.date_range(start, end).date})
    employees["_key"] = 1
    dates["_key"] = 1
    daily = employees.merge(dates, on="_key").drop(columns="_key")

    daily["Hari"] = pd.to_datetime(daily["Tanggal"]).dt.day_name().map(DAY_ID)
    daily["Is_Holiday"] = daily["Tanggal"].isin(holidays_set)

    working_days = cfg.working_days

    daily["Is_Approved_Leave"] = False
    daily["Leave_Type"] = ""
    if not leave.empty:
        daily = daily.merge(leave, on=["No.", "Tanggal"], how="left", suffixes=("", "_leave"))
        leave_col = "Leave_Type_leave" if "Leave_Type_leave" in daily.columns else "Leave_Type"
        daily["Is_Approved_Leave"] = daily[leave_col].notna()
        daily["Leave_Type"] = daily[leave_col].fillna("")
        if leave_col != "Leave_Type":
            daily = daily.drop(columns=[leave_col])

    daily["Is_Working_Day"] = (
        daily["Hari"].isin(working_days) & ~daily["Is_Holiday"] & ~daily["Is_Approved_Leave"]
    )

    scans = (
        enriched_log.groupby(["No.", "Tanggal"], group_keys=False)
        .apply(lambda g: _scan_pair(g, cfg.single_scan_cutoff))
        .reset_index()
    )
    daily = daily.merge(scans, on=["No.", "Tanggal"], how="left")
    daily["Scan_Count"] = daily["Scan_Count"].fillna(0).astype(int)

    daily["Jam_Masuk"] = daily["First_Scan"].dt.strftime("%H:%M:%S").fillna("")
    daily["Jam_Pulang"] = daily["Last_Scan"].dt.strftime("%H:%M:%S").fillna("")

    def evaluate(row: pd.Series) -> pd.Series:
        tap_in = row["First_Scan"]
        tap_out = row["Last_Scan"]
        day = row["Tanggal"]
        is_working = bool(row["Is_Working_Day"])
        scan_count = int(row["Scan_Count"])

        if bool(row.get("Is_Approved_Leave", False)):
            return pd.Series(
                {
                    "Status_Masuk": "Cuti Disetujui",
                    "Status_Pulang": str(row.get("Leave_Type") or "Approved Leave"),
                    "Menit_Telat": 0,
                    "Menit_Pulang_Cepat": 0,
                    "Anomaly_Type": "",
                    "Severity": "",
                }
            )

        if not is_working:
            if scan_count == 0:
                return pd.Series(
                    {
                        "Status_Masuk": "Libur",
                        "Status_Pulang": "Libur Normal",
                        "Menit_Telat": 0,
                        "Menit_Pulang_Cepat": 0,
                        "Anomaly_Type": "",
                        "Severity": "",
                    }
                )
            return pd.Series(
                {
                    "Status_Masuk": "Scan Hari Libur",
                    "Status_Pulang": "Scan Hari Libur",
                    "Menit_Telat": 0,
                    "Menit_Pulang_Cepat": 0,
                    "Anomaly_Type": "Scan Pada Hari Libur",
                    "Severity": "LOW",
                }
            )

        # Custom division/employee schedule from master or global fallback
        emp_in_time = _parse_time_value(row.get("Jam_Masuk_Master")) or cfg.clock_in
        schedule_in = _combine(day, emp_in_time)
        grace_until = schedule_in + pd.Timedelta(minutes=cfg.grace_minutes)

        emp_out_time = _parse_time_value(row.get("Jam_Pulang_Master"))
        if emp_out_time is None:
            emp_out_time = cfg.saturday_clock_out if row["Hari"] == "Sabtu" else cfg.weekday_clock_out
        schedule_out = _combine(day, emp_out_time)

        if scan_count == 0:
            return pd.Series(
                {
                    "Status_Masuk": "Mangkir",
                    "Status_Pulang": "Mangkir",
                    "Menit_Telat": 0,
                    "Menit_Pulang_Cepat": 0,
                    "Anomaly_Type": "Tidak Ada Transaksi",
                    "Severity": "HIGH",
                }
            )

        anomaly_types: list[str] = []

        if pd.isna(tap_in):
            status_in = "Lupa Absen Masuk"
            minutes_late = cfg.default_late_minutes_when_only_clock_out
            anomaly_types.append("Hanya Scan Siang/Sore")
        elif tap_in > grace_until:
            status_in = "Terlambat"
            minutes_late = max(0, int((tap_in - schedule_in).total_seconds() // 60))
        else:
            status_in = "Hadir"
            minutes_late = 0

        if pd.isna(tap_out):
            status_out = "Lupa Absen Pulang"
            minutes_early = 0
            if "Hanya Scan Siang/Sore" not in anomaly_types:
                anomaly_types.append("Hanya Scan Pagi")
        elif tap_out < schedule_out:
            minutes_early = max(0, int((schedule_out - tap_out).total_seconds() // 60))
            status_out = "Pulang Cepat"
        else:
            minutes_early = 0
            status_out = "OK"

        if scan_count > 2:
            anomaly_types.append("Transaksi Lebih Dari Dua Kali")

        severities = []
        if status_in == "Lupa Absen Masuk" or status_out == "Lupa Absen Pulang":
            severities.append("HIGH")
        if status_in == "Terlambat" or status_out == "Pulang Cepat":
            severities.append("LOW")
        if scan_count > 2:
            severities.append("LOW")

        severity = "HIGH" if "HIGH" in severities else ("LOW" if severities else "")

        return pd.Series(
            {
                "Status_Masuk": status_in,
                "Status_Pulang": status_out,
                "Menit_Telat": minutes_late,
                "Menit_Pulang_Cepat": minutes_early,
                "Anomaly_Type": ", ".join(anomaly_types),
                "Severity": severity,
            }
        )

    evaluation = daily.apply(evaluate, axis=1)
    daily = pd.concat([daily, evaluation], axis=1)

    duration = (daily["Last_Scan"] - daily["First_Scan"])
    daily["Work_Duration_Minutes"] = (
        duration.dt.total_seconds().div(60).where(duration.notna(), np.nan).round(0)
    )

    daily["Is_Anomali"] = (daily["Severity"] != "").astype(int)
    daily["Work_Day_Flag"] = daily["Is_Working_Day"].astype(int)
    daily["Present_Flag"] = daily["Status_Masuk"].isin(["Hadir", "Terlambat"]).astype(int)
    daily["Late_Flag"] = daily["Status_Masuk"].eq("Terlambat").astype(int)
    daily["Absent_Flag"] = daily["Status_Masuk"].eq("Mangkir").astype(int)
    daily["Early_Leave_Flag"] = daily["Status_Pulang"].eq("Pulang Cepat").astype(int)
    daily["Forgot_Punch_Flag"] = (
        daily["Status_Masuk"].eq("Lupa Absen Masuk") | daily["Status_Pulang"].eq("Lupa Absen Pulang")
    ).astype(int)
    daily["Tidak_Absen_Pulang_Flag"] = daily["Status_Pulang"].eq("Lupa Absen Pulang").astype(int)

    # If leave result (multi-sheet Cuti, Lembur, Kegiatan Bersama) provided, apply integration
    if lr is not None:
        from ..leave import apply_leave_kegiatan
        daily = apply_leave_kegiatan(daily, lr)

    ordered_cols = [
        "No.", "Name", "Department", "Employee_Type", "Position",
        "Tanggal", "Hari", "Is_Working_Day", "Is_Holiday",
        "Jam_Masuk", "Jam_Pulang", "Scan_Count",
        "Status_Masuk", "Status_Pulang", "Menit_Telat", "Menit_Pulang_Cepat",
        "Work_Duration_Minutes", "Anomaly_Type", "Severity", "Is_Anomali",
        "Work_Day_Flag", "Present_Flag", "Late_Flag", "Absent_Flag",
        "Early_Leave_Flag", "Forgot_Punch_Flag", "Tidak_Absen_Pulang_Flag",
        "Overtime_Hours", "Overtime_Status", "Group_Event",
        "First_Scan", "Last_Scan",
    ]
    ordered_cols = [c for c in ordered_cols if c in daily.columns]
    daily = daily[ordered_cols]

    return daily.sort_values(["Tanggal", "Department", "Name"], kind="stable").reset_index(drop=True)
