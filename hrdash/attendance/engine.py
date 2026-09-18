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

import datetime as dt
from datetime import date, datetime, time, timedelta
from typing import Iterable, Optional

import polars as pl

from ..config import AttendanceConfig
from .parser import DAY_ID, WEEKDAY_NUM_MAP


def _parse_time_str(val: object) -> Optional[str]:
    if val is None:
        return None
    if isinstance(val, time):
        return val.strftime("%H:%M:%S")
    if isinstance(val, datetime):
        return val.time().strftime("%H:%M:%S")
    s = str(val).strip().replace(".", ":")
    if not s or s.lower() in ("none", "nan", "nat", "<null>", ""):
        return None
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            h, m = int(parts[0]), int(parts[1])
            sec = int(parts[2]) if len(parts) > 2 else 0
            return f"{h:02d}:{m:02d}:{sec:02d}"
        except ValueError:
            return None
    return None


def build_employee_directory(df: pl.DataFrame) -> pl.DataFrame:
    """One row per employee, resolving the Name/Department/Employee_Type/Schedule
    seen in the log or merged master.
    """
    if df.is_empty():
        return pl.DataFrame(
            schema={
                "No.": pl.String,
                "Name": pl.String,
                "Department": pl.String,
                "Employee_Type": pl.String,
                "Position": pl.String,
                "Jam_Masuk_Master": pl.String,
                "Jam_Pulang_Master": pl.String,
            }
        )

    agg_exprs = [
        pl.col("Name").drop_nulls().first().alias("Name"),
        pl.col("Department").drop_nulls().first().alias("Department"),
    ]
    if "Employee_Type" in df.columns:
        agg_exprs.append(pl.col("Employee_Type").drop_nulls().first().fill_null("OFFICE").alias("Employee_Type"))
    else:
        agg_exprs.append(pl.lit("OFFICE", dtype=pl.String).alias("Employee_Type"))

    if "Position" in df.columns:
        agg_exprs.append(pl.col("Position").drop_nulls().first().fill_null("").alias("Position"))
    else:
        agg_exprs.append(pl.lit("", dtype=pl.String).alias("Position"))

    if "Jam_Masuk_Master" in df.columns:
        agg_exprs.append(pl.col("Jam_Masuk_Master").drop_nulls().first().alias("Jam_Masuk_Master"))
    else:
        agg_exprs.append(pl.lit(None, dtype=pl.String).alias("Jam_Masuk_Master"))

    if "Jam_Pulang_Master" in df.columns:
        agg_exprs.append(pl.col("Jam_Pulang_Master").drop_nulls().first().alias("Jam_Pulang_Master"))
    else:
        agg_exprs.append(pl.lit(None, dtype=pl.String).alias("Jam_Pulang_Master"))

    directory = df.group_by("No.").agg(agg_exprs)
    return directory.sort(["Department", "Name"])


def build_daily_attendance(
    enriched_log: pl.DataFrame,
    cfg: AttendanceConfig,
    holidays: Optional[Iterable[date]] = None,
    approved_leave: Optional[pl.DataFrame | object] = None,
    leave_result: Optional[object] = None,
) -> pl.DataFrame:
    """Build the Daily Attendance table (Section 5) using native vectorized Polars expressions."""
    if enriched_log.is_empty():
        return pl.DataFrame()

    df_log = enriched_log
    holidays_set = set(holidays or [])

    # Support both legacy DataFrame and LeaveLoadResult
    lr = leave_result
    if lr is None and hasattr(approved_leave, "leaves") and hasattr(approved_leave, "group_events"):
        lr = approved_leave
        leave = pl.DataFrame(schema={"No.": pl.String, "Tanggal": pl.Date, "Leave_Type": pl.String})
    elif isinstance(approved_leave, pl.DataFrame):
        leave = approved_leave
    elif approved_leave is not None and hasattr(approved_leave, "to_dict"):
        leave = pl.from_pandas(approved_leave)
    else:
        leave = pl.DataFrame(schema={"No.": pl.String, "Tanggal": pl.Date, "Leave_Type": pl.String})

    start = cfg.start_date or df_log["Tanggal"].min()
    end = cfg.end_date or df_log["Tanggal"].max()

    employees = build_employee_directory(df_log)
    date_list = [start + timedelta(days=i) for i in range((end - start).days + 1)]
    dates_df = pl.DataFrame({"Tanggal": date_list}, schema={"Tanggal": pl.Date})

    daily = employees.join(dates_df, how="cross")

    daily = daily.with_columns(
        Hari=pl.col("Tanggal").dt.weekday().replace_strict(WEEKDAY_NUM_MAP, default="Senin"),
        Is_Holiday=pl.col("Tanggal").is_in(list(holidays_set)),
    )

    working_days = list(cfg.working_days)

    if not leave.is_empty() and "No." in leave.columns and "Tanggal" in leave.columns:
        leave_sub = leave.select(["No.", "Tanggal", "Leave_Type"]).unique(subset=["No.", "Tanggal"])
        daily = daily.join(leave_sub, on=["No.", "Tanggal"], how="left")
        daily = daily.with_columns(
            Is_Approved_Leave=pl.col("Leave_Type").is_not_null() & (pl.col("Leave_Type") != ""),
            Leave_Type=pl.col("Leave_Type").fill_null(""),
        )
    else:
        daily = daily.with_columns(
            Is_Approved_Leave=pl.lit(False, dtype=pl.Boolean),
            Leave_Type=pl.lit("", dtype=pl.String),
        )

    daily = daily.with_columns(
        Is_Working_Day=pl.col("Hari").is_in(working_days) & ~pl.col("Is_Holiday") & ~pl.col("Is_Approved_Leave")
    )

    # Vectorized scan aggregation per employee per day
    cutoff_time = cfg.single_scan_cutoff
    cutoff_str = cutoff_time.strftime("%H:%M:%S")

    scans_grouped = (
        df_log.sort(["No.", "Tanggal", "Date/Time"])
        .group_by(["No.", "Tanggal"], maintain_order=True)
        .agg(
            pl.col("Date/Time").first().alias("_first_raw"),
            pl.col("Date/Time").last().alias("_last_raw"),
            pl.len().alias("Scan_Count"),
        )
    )

    # Vectorize single-scan vs multi-scan extraction
    is_single_morning = (pl.col("Scan_Count") == 1) & (
        pl.col("_first_raw").dt.strftime("%H:%M:%S") < cutoff_str
    )
    is_single_afternoon = (pl.col("Scan_Count") == 1) & (
        pl.col("_first_raw").dt.strftime("%H:%M:%S") >= cutoff_str
    )
    is_multi = pl.col("Scan_Count") > 1

    scans_df = scans_grouped.with_columns(
        First_Scan=pl.when(is_single_morning | is_multi)
        .then(pl.col("_first_raw"))
        .otherwise(None)
        .cast(pl.Datetime),
        Last_Scan=pl.when(is_single_afternoon)
        .then(pl.col("_first_raw"))
        .when(is_multi)
        .then(pl.col("_last_raw"))
        .otherwise(None)
        .cast(pl.Datetime),
    ).select(["No.", "Tanggal", "First_Scan", "Last_Scan", "Scan_Count"])

    daily = daily.join(scans_df, on=["No.", "Tanggal"], how="left")
    daily = daily.with_columns(Scan_Count=pl.col("Scan_Count").fill_null(0).cast(pl.Int64))

    # Vectorized schedule calculation
    in_time_str = cfg.clock_in.strftime("%H:%M:%S")
    out_weekday_str = cfg.weekday_clock_out.strftime("%H:%M:%S")
    out_sat_str = cfg.saturday_clock_out.strftime("%H:%M:%S")

    # Format Master In/Out to HH:MM:SS format
    daily = daily.with_columns(
        _in_time_norm=pl.coalesce(
            [
                pl.col("Jam_Masuk_Master").str.to_time("%H:%M", strict=False).dt.strftime("%H:%M:%S"),
                pl.col("Jam_Masuk_Master").str.to_time("%H:%M:%S", strict=False).dt.strftime("%H:%M:%S"),
                pl.lit(in_time_str),
            ]
        ),
        _out_default=pl.when(pl.col("Hari") == "Sabtu").then(pl.lit(out_sat_str)).otherwise(pl.lit(out_weekday_str)),
    )
    daily = daily.with_columns(
        _out_time_norm=pl.coalesce(
            [
                pl.col("Jam_Pulang_Master").str.to_time("%H:%M", strict=False).dt.strftime("%H:%M:%S"),
                pl.col("Jam_Pulang_Master").str.to_time("%H:%M:%S", strict=False).dt.strftime("%H:%M:%S"),
                pl.col("_out_default"),
            ]
        )
    )

    # Combine Tanggal + Time into Datetime
    schedule_in_expr = (
        pl.concat_str([pl.col("Tanggal").cast(pl.String), pl.lit(" "), pl.col("_in_time_norm")])
        .str.to_datetime("%Y-%m-%d %H:%M:%S", strict=False)
    )
    grace_until_expr = schedule_in_expr + pl.duration(minutes=cfg.grace_minutes)

    schedule_out_expr = (
        pl.concat_str([pl.col("Tanggal").cast(pl.String), pl.lit(" "), pl.col("_out_time_norm")])
        .str.to_datetime("%Y-%m-%d %H:%M:%S", strict=False)
    )

    daily = daily.with_columns(
        _schedule_in=schedule_in_expr,
        _grace_until=grace_until_expr,
        _schedule_out=schedule_out_expr,
    )

    # Vectorized Status Masuk
    status_masuk_expr = (
        pl.when(pl.col("Is_Approved_Leave"))
        .then(pl.lit("Cuti Disetujui"))
        .when(~pl.col("Is_Working_Day") & (pl.col("Scan_Count") == 0))
        .then(pl.lit("Libur"))
        .when(~pl.col("Is_Working_Day"))
        .then(pl.lit("Scan Hari Libur"))
        .when(pl.col("Scan_Count") == 0)
        .then(pl.lit("Mangkir"))
        .when(pl.col("First_Scan").is_null())
        .then(pl.lit("Lupa Absen Masuk"))
        .when(pl.col("First_Scan") > pl.col("_grace_until"))
        .then(pl.lit("Terlambat"))
        .otherwise(pl.lit("Hadir"))
    )

    # Vectorized Status Pulang
    status_pulang_expr = (
        pl.when(pl.col("Is_Approved_Leave"))
        .then(pl.when(pl.col("Leave_Type") != "").then(pl.col("Leave_Type")).otherwise(pl.lit("Approved Leave")))
        .when(~pl.col("Is_Working_Day") & (pl.col("Scan_Count") == 0))
        .then(pl.lit("Libur Normal"))
        .when(~pl.col("Is_Working_Day"))
        .then(pl.lit("Scan Hari Libur"))
        .when(pl.col("Scan_Count") == 0)
        .then(pl.lit("Mangkir"))
        .when(pl.col("Last_Scan").is_null())
        .then(pl.lit("Lupa Absen Pulang"))
        .when(pl.col("Last_Scan") < pl.col("_schedule_out"))
        .then(pl.lit("Pulang Cepat"))
        .otherwise(pl.lit("OK"))
    )

    daily = daily.with_columns(
        Status_Masuk=status_masuk_expr,
        Status_Pulang=status_pulang_expr,
    )

    # Vectorized Menit Telat & Pulang Cepat
    late_diff_minutes = (
        (pl.col("First_Scan") - pl.col("_schedule_in")).dt.total_seconds() / 60.0
    ).floor().cast(pl.Int64).clip(lower_bound=0)

    early_diff_minutes = (
        (pl.col("_schedule_out") - pl.col("Last_Scan")).dt.total_seconds() / 60.0
    ).floor().cast(pl.Int64).clip(lower_bound=0)

    daily = daily.with_columns(
        Menit_Telat=pl.when(pl.col("Is_Approved_Leave") | ~pl.col("Is_Working_Day") | (pl.col("Scan_Count") == 0))
        .then(pl.lit(0, dtype=pl.Int64))
        .when(pl.col("Status_Masuk") == "Lupa Absen Masuk")
        .then(pl.lit(cfg.default_late_minutes_when_only_clock_out, dtype=pl.Int64))
        .when(pl.col("Status_Masuk") == "Terlambat")
        .then(late_diff_minutes)
        .otherwise(pl.lit(0, dtype=pl.Int64)),
        Menit_Pulang_Cepat=pl.when(
            pl.col("Is_Approved_Leave") | ~pl.col("Is_Working_Day") | (pl.col("Scan_Count") == 0)
        )
        .then(pl.lit(0, dtype=pl.Int64))
        .when(pl.col("Status_Pulang") == "Pulang Cepat")
        .then(early_diff_minutes)
        .otherwise(pl.lit(0, dtype=pl.Int64)),
    )

    # Vectorized Anomaly_Type & Severity
    anom_list_expr = pl.concat_list(
        [
            pl.when(
                pl.col("First_Scan").is_null()
                & (pl.col("Scan_Count") > 0)
                & pl.col("Is_Working_Day")
                & ~pl.col("Is_Approved_Leave")
            )
            .then(pl.lit("Hanya Scan Siang/Sore"))
            .otherwise(None),
            pl.when(
                pl.col("Last_Scan").is_null()
                & (pl.col("Scan_Count") > 0)
                & pl.col("Is_Working_Day")
                & pl.col("First_Scan").is_not_null()
                & ~pl.col("Is_Approved_Leave")
            )
            .then(pl.lit("Hanya Scan Pagi"))
            .otherwise(None),
            pl.when(
                (pl.col("Scan_Count") > 2)
                & pl.col("Is_Working_Day")
                & ~pl.col("Is_Approved_Leave")
            )
            .then(pl.lit("Transaksi Lebih Dari Dua Kali"))
            .otherwise(None),
        ]
    ).list.drop_nulls().list.join(", ")

    anomaly_type_expr = (
        pl.when(pl.col("Is_Approved_Leave"))
        .then(pl.lit(""))
        .when(~pl.col("Is_Working_Day") & (pl.col("Scan_Count") > 0))
        .then(pl.lit("Scan Pada Hari Libur"))
        .when(~pl.col("Is_Working_Day"))
        .then(pl.lit(""))
        .when(pl.col("Scan_Count") == 0)
        .then(pl.lit("Tidak Ada Transaksi"))
        .otherwise(anom_list_expr)
    )

    severity_expr = (
        pl.when(pl.col("Is_Approved_Leave"))
        .then(pl.lit(""))
        .when(~pl.col("Is_Working_Day") & (pl.col("Scan_Count") > 0))
        .then(pl.lit("LOW"))
        .when(~pl.col("Is_Working_Day"))
        .then(pl.lit(""))
        .when(pl.col("Scan_Count") == 0)
        .then(pl.lit("HIGH"))
        .when(
            (pl.col("Status_Masuk") == "Lupa Absen Masuk")
            | (pl.col("Status_Pulang") == "Lupa Absen Pulang")
        )
        .then(pl.lit("HIGH"))
        .when(
            (pl.col("Status_Masuk") == "Terlambat")
            | (pl.col("Status_Pulang") == "Pulang Cepat")
            | (pl.col("Scan_Count") > 2)
        )
        .then(pl.lit("LOW"))
        .otherwise(pl.lit(""))
    )

    daily = daily.with_columns(
        Anomaly_Type=anomaly_type_expr,
        Severity=severity_expr,
    )

    # Duration and flags
    duration_secs = (pl.col("Last_Scan") - pl.col("First_Scan")).dt.total_seconds()
    daily = daily.with_columns(
        Jam_Masuk=pl.col("First_Scan").dt.strftime("%H:%M:%S").fill_null(""),
        Jam_Pulang=pl.col("Last_Scan").dt.strftime("%H:%M:%S").fill_null(""),
        Work_Duration_Minutes=pl.when(pl.col("First_Scan").is_not_null() & pl.col("Last_Scan").is_not_null())
        .then((duration_secs / 60.0).round(0))
        .otherwise(None),
        Is_Anomali=pl.when(pl.col("Severity") != "").then(1).otherwise(0).cast(pl.Int64),
        Work_Day_Flag=pl.when(pl.col("Is_Working_Day")).then(1).otherwise(0).cast(pl.Int64),
        Present_Flag=pl.when(pl.col("Status_Masuk").is_in(["Hadir", "Terlambat"])).then(1).otherwise(0).cast(pl.Int64),
        Late_Flag=pl.when(pl.col("Status_Masuk") == "Terlambat").then(1).otherwise(0).cast(pl.Int64),
        Absent_Flag=pl.when(pl.col("Status_Masuk") == "Mangkir").then(1).otherwise(0).cast(pl.Int64),
        Early_Leave_Flag=pl.when(pl.col("Status_Pulang") == "Pulang Cepat").then(1).otherwise(0).cast(pl.Int64),
        Forgot_Punch_Flag=pl.when(
            (pl.col("Status_Masuk") == "Lupa Absen Masuk") | (pl.col("Status_Pulang") == "Lupa Absen Pulang")
        )
        .then(1)
        .otherwise(0)
        .cast(pl.Int64),
        Tidak_Absen_Pulang_Flag=pl.when(pl.col("Status_Pulang") == "Lupa Absen Pulang").then(1).otherwise(0).cast(pl.Int64),
    )

    # Drop temporary calculation columns
    daily = daily.drop(
        ["_in_time_norm", "_out_default", "_out_time_norm", "_schedule_in", "_grace_until", "_schedule_out"]
    )

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
    avail_cols = [c for c in ordered_cols if c in daily.columns]
    daily = daily.select(avail_cols)

    return daily.sort(["Tanggal", "Department", "Name"])
