"""
Raw attendance log ingestion.

This module only validates and cleans the raw machine export — it never
decides what a scan *means* (that is the attendance engine's job). Raw rows
are never mutated or dropped silently: anything excluded from calculation is
returned in a separate ``rejected`` DataFrame with a reason, preserving full
audit trail (Section 15).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

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


class AttendanceValidationError(ValueError):
    """Raised when the uploaded attendance file cannot be processed safely."""


def _clean_column_name(value: object) -> str:
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def parse_datetime_series(series: pd.Series) -> pd.Series:
    """Robustly parse attendance-machine timestamps.

    The machine export uses ``DD/MM/YYYY HH.MM.SS`` (dots instead of colons
    for the time portion); both dot and colon separators are accepted.
    """
    s = series.astype("string").str.strip()
    s = s.str.replace(".", ":", regex=False)

    parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")

    missing = parsed.isna()
    if missing.any():
        parsed.loc[missing] = pd.to_datetime(
            s.loc[missing], format="%d/%m/%Y %H:%M:%S", errors="coerce"
        )
    return parsed


def validate_schema(df: pd.DataFrame) -> list[str]:
    issues: list[str] = []
    missing = REQUIRED_COLUMNS.difference(df.columns)
    if missing:
        issues.append(f"Kolom wajib tidak ditemukan: {', '.join(sorted(missing))}")
    if df.empty:
        issues.append("File tidak memiliki baris data.")
    return issues


def clean_attendance_data(raw_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Validate + clean a raw attendance-machine export.

    Returns
    -------
    clean_df   rows safe to feed into the attendance engine
    rejected_df rows excluded, each tagged with ``_reject_reason``
    quality    data-quality metrics for the UI / README sheet
    """
    df = raw_df.copy()
    df.columns = [_clean_column_name(c) for c in df.columns]

    schema_issues = validate_schema(df)
    if schema_issues:
        raise AttendanceValidationError("; ".join(schema_issues))

    df["_source_row"] = np.arange(2, len(df) + 2)

    for col in ["No.", "Name", "Department"]:
        if col not in df.columns:
            df[col] = pd.NA
        df[col] = df[col].astype("string").str.strip()

    rejected_parts: list[pd.DataFrame] = []

    invalid_identity = df["No."].isna() | (df["No."] == "")
    if invalid_identity.any():
        bad = df.loc[invalid_identity].copy()
        bad["_reject_reason"] = "Employee ID (No.) kosong/tidak valid"
        rejected_parts.append(bad)
    df = df.loc[~invalid_identity].copy()

    # Missing identity metadata gets a clear fallback rather than rejection —
    # the machine log itself is still valid attendance evidence.
    df["Name"] = df["Name"].fillna("").replace("", pd.NA)
    df["Department"] = df["Department"].fillna("").replace("", pd.NA)
    df["Name"] = df["Name"].fillna(df["No."].map(lambda x: f"Karyawan {x}"))
    df["Department"] = df["Department"].fillna("Belum Dipetakan")

    df["Date/Time"] = parse_datetime_series(df["Date/Time"])
    invalid_datetime = df["Date/Time"].isna()
    if invalid_datetime.any():
        bad = df.loc[invalid_datetime].copy()
        bad["_reject_reason"] = "Tanggal/jam tidak dapat dibaca"
        rejected_parts.append(bad)
    df = df.loc[~invalid_datetime].copy()

    rejected = (
        pd.concat(rejected_parts, ignore_index=True)
        if rejected_parts
        else pd.DataFrame(columns=list(df.columns) + ["_reject_reason"])
    )

    df = df.sort_values(["No.", "Date/Time", "_source_row"]).reset_index(drop=True)
    before_dupes = len(df)
    df = df.drop_duplicates(subset=["No.", "Date/Time"], keep="first").reset_index(drop=True)
    duplicate_count = before_dupes - len(df)

    df["Tanggal"] = df["Date/Time"].dt.date
    df["Hari"] = df["Date/Time"].dt.day_name().map(DAY_ID)
    df["Jam"] = df["Date/Time"].dt.time

    quality = {
        "input_rows": len(raw_df),
        "valid_rows": len(df),
        "rejected_rows": len(rejected),
        "duplicate_rows_removed": duplicate_count,
        "date_min": df["Tanggal"].min() if not df.empty else None,
        "date_max": df["Tanggal"].max() if not df.empty else None,
        "employees": int(df["No."].nunique()) if not df.empty else 0,
    }
    return df, rejected, quality


def clean_approved_leave(leave_df: pd.DataFrame | None) -> pd.DataFrame:
    """Optional HR-approved leave/permission input (excludes days from
    absence calculation; never invented, only used when HR supplies it).
    """
    if leave_df is None or leave_df.empty:
        return pd.DataFrame(columns=["No.", "Tanggal", "Leave_Type"])

    df = leave_df.copy()
    df.columns = [_clean_column_name(c) for c in df.columns]

    id_col = next((c for c in ["No.", "Employee ID", "ID", "NIK"] if c in df.columns), None)
    date_col = next((c for c in ["Tanggal", "Date", "Leave Date"] if c in df.columns), None)
    type_col = next((c for c in ["Type", "Leave Type", "Jenis"] if c in df.columns), None)

    if not id_col or not date_col:
        raise AttendanceValidationError(
            "File leave harus memiliki kolom employee ID (No./Employee ID/ID/NIK) "
            "dan tanggal (Tanggal/Date/Leave Date)."
        )

    out = pd.DataFrame()
    out["No."] = df[id_col].astype("string").str.strip()
    out["Tanggal"] = pd.to_datetime(
        df[date_col], format="mixed", dayfirst=True, errors="coerce"
    ).dt.date
    out["Leave_Type"] = df[type_col].astype("string").str.strip() if type_col else "Approved Leave"
    out = out.dropna(subset=["No.", "Tanggal"]).drop_duplicates(["No.", "Tanggal"], keep="last")
    return out.reset_index(drop=True)
