"""
Raw attendance log ingestion.

This module only validates and cleans the raw machine export — it never
decides what a scan *means* (that is the attendance engine's job). Raw rows
are never mutated or dropped silently: anything excluded from calculation is
returned in a separate ``rejected`` DataFrame with a reason, preserving full
audit trail (Section 15).
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Optional

import polars as pl

REQUIRED_COLUMNS = {"No.", "Date/Time"}

DAY_ID = {
    "Monday": "Senin",
    "Tuesday": "Selasa",
    "Wednesday": "Rabu",
    "Thursday": "Kamis",
    "Friday": "Jumat",
    "Saturday": "Sabtu",
    "Sunday": "Minggu",
}

WEEKDAY_NUM_MAP = {
    1: "Senin",
    2: "Selasa",
    3: "Rabu",
    4: "Kamis",
    5: "Jumat",
    6: "Sabtu",
    7: "Minggu",
}


class AttendanceValidationError(ValueError):
    """Raised when the uploaded attendance file cannot be processed safely."""


def _clean_column_name(value: object) -> str:
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def parse_datetime_expr(col_name: str = "Date/Time") -> pl.Expr:
    """Robustly parse attendance-machine timestamps into pl.Datetime."""
    s = pl.col(col_name).cast(pl.String).str.strip_chars().str.replace_all(r"\.", ":")
    return pl.coalesce(
        [
            s.str.to_datetime("%d/%m/%Y %H:%M:%S", strict=False),
            s.str.to_datetime("%Y-%m-%d %H:%M:%S", strict=False),
            s.str.to_datetime("%d-%m-%Y %H:%M:%S", strict=False),
            s.str.to_datetime("%Y/%m/%d %H:%M:%S", strict=False),
            s.str.to_datetime("%d/%m/%Y %H:%M", strict=False),
            s.str.to_datetime("%Y-%m-%d %H:%M", strict=False),
            s.str.to_datetime(strict=False),
        ]
    )


def validate_schema(df: pl.DataFrame) -> list[str]:
    issues: list[str] = []
    missing = REQUIRED_COLUMNS.difference(set(df.columns))
    if missing:
        issues.append(f"Kolom wajib tidak ditemukan: {', '.join(sorted(missing))}")
    if df.is_empty():
        issues.append("File tidak memiliki baris data.")
    return issues


def clean_attendance_data(raw_input: pl.DataFrame) -> tuple[pl.DataFrame, pl.DataFrame, dict]:
    """Validate + clean a raw attendance-machine export.

    Returns
    -------
    clean_df   rows safe to feed into the attendance engine
    rejected_df rows excluded, each tagged with ``_reject_reason``
    quality    data-quality metrics for the UI / README sheet
    """
    raw_df = raw_input
    if raw_df.is_empty() and not raw_df.columns:
        raise AttendanceValidationError("File tidak memiliki baris data.")

    rename_map = {c: _clean_column_name(c) for c in raw_df.columns}
    df = raw_df.rename(rename_map)

    schema_issues = validate_schema(df)
    if schema_issues:
        raise AttendanceValidationError("; ".join(schema_issues))

    input_row_count = len(df)
    df = df.with_columns(_source_row=pl.int_range(2, pl.len() + 2, dtype=pl.Int64))

    for col in ["No.", "Name", "Department"]:
        if col not in df.columns:
            df = df.with_columns(pl.lit(None, dtype=pl.String).alias(col))
        else:
            df = df.with_columns(pl.col(col).cast(pl.String).str.strip_chars().alias(col))

    rejected_parts: list[pl.DataFrame] = []

    # 1. Check valid identity
    invalid_identity_mask = pl.col("No.").is_null() | (pl.col("No.").str.strip_chars() == "")
    bad_id = df.filter(invalid_identity_mask)
    if not bad_id.is_empty():
        bad_id = bad_id.with_columns(_reject_reason=pl.lit("Employee ID (No.) kosong/tidak valid"))
        rejected_parts.append(bad_id)
    df = df.filter(~invalid_identity_mask)

    # Missing metadata fallback
    df = df.with_columns(
        Name=pl.when(pl.col("Name").is_null() | (pl.col("Name").str.strip_chars() == ""))
        .then(pl.concat_str([pl.lit("Karyawan "), pl.col("No.")]))
        .otherwise(pl.col("Name")),
        Department=pl.when(pl.col("Department").is_null() | (pl.col("Department").str.strip_chars() == ""))
        .then(pl.lit("Belum Dipetakan"))
        .otherwise(pl.col("Department")),
    )

    # 2. Parse Date/Time
    parsed_dt = parse_datetime_expr("Date/Time")
    df = df.with_columns(_parsed_dt=parsed_dt)

    invalid_dt_mask = pl.col("_parsed_dt").is_null()
    bad_dt = df.filter(invalid_dt_mask)
    if not bad_dt.is_empty():
        bad_dt = bad_dt.drop("_parsed_dt").with_columns(
            _reject_reason=pl.lit("Tanggal/jam tidak dapat dibaca")
        )
        rejected_parts.append(bad_dt)
    df = df.filter(~invalid_dt_mask)
    df = df.with_columns(pl.col("_parsed_dt").alias("Date/Time")).drop("_parsed_dt")

    all_cols = list(df.columns)
    if rejected_parts:
        rejected = pl.concat(
            [part.select([c for c in all_cols if c in part.columns] + ["_reject_reason"]) for part in rejected_parts],
            how="diagonal",
        )
    else:
        schema = {c: df.schema.get(c, pl.String) for c in all_cols}
        schema["_reject_reason"] = pl.String
        rejected = pl.DataFrame(schema=schema)

    df = df.sort(["No.", "Date/Time", "_source_row"])

    # 3. Deduplicate
    before_dedup = len(df)
    df = df.unique(subset=["No.", "Date/Time"], keep="first")
    duplicate_count = before_dedup - len(df)

    df = df.with_columns(
        Tanggal=pl.col("Date/Time").dt.date(),
        Waktu=pl.col("Date/Time").dt.time(),
        Hari=pl.col("Date/Time").dt.weekday().replace_strict(WEEKDAY_NUM_MAP, default="Senin"),
    )

    min_date = df["Tanggal"].min() if not df.is_empty() else None
    max_date = df["Tanggal"].max() if not df.is_empty() else None
    employee_count = df["No."].n_unique() if not df.is_empty() else 0

    quality = {
        "input_rows": input_row_count,
        "valid_rows": len(df),
        "rejected_rows": len(rejected),
        "duplicate_rows_removed": duplicate_count,
        "date_min": min_date,
        "date_max": max_date,
        "employees": employee_count,
    }
    return df, rejected, quality


def clean_approved_leave(leave_input: Optional[pl.DataFrame]) -> pl.DataFrame:
    """Optional HR-approved leave/permission input (excludes days from
    absence calculation; never invented, only used when HR supplies it).
    """
    if leave_input is None or leave_input.is_empty():
        return pl.DataFrame(schema={"No.": pl.String, "Tanggal": pl.Date, "Leave_Type": pl.String})

    df = leave_input
    rename_map = {c: _clean_column_name(c) for c in df.columns}
    df = df.rename(rename_map)

    id_col = next((c for c in ["No.", "Employee ID", "ID", "NIK"] if c in df.columns), None)
    date_col = next((c for c in ["Tanggal", "Date", "Leave Date"] if c in df.columns), None)
    type_col = next((c for c in ["Type", "Leave Type", "Jenis"] if c in df.columns), None)

    if not id_col or not date_col:
        raise AttendanceValidationError(
            "File leave harus memiliki kolom employee ID (No./Employee ID/ID/NIK) "
            "dan tanggal (Tanggal/Date/Leave Date)."
        )

    out = df.select(
        pl.col(id_col).cast(pl.String).str.strip_chars().alias("No."),
        parse_datetime_expr(date_col).dt.date().alias("Tanggal"),
        (
            pl.col(type_col).cast(pl.String).str.strip_chars()
            if type_col
            else pl.lit("Approved Leave", dtype=pl.String)
        ).alias("Leave_Type"),
    )
    out = out.filter(pl.col("No.").is_not_null() & (pl.col("No.") != "") & pl.col("Tanggal").is_not_null())
    out = out.unique(subset=["No.", "Tanggal"], keep="last")
    return out
