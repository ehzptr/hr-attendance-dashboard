"""
Payroll deduction engine (Section 7).

The engine only ever computes a monetary deduction when HR has explicitly
supplied the corresponding rate in ``PayrollConfig``. When a rate is
missing, the column is left as ``pd.NA`` (rendered as "Belum Diatur" in the
Excel export) instead of silently defaulting to zero or to an invented
number — the spec is explicit that HR-defined nominal values must never be
assumed.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

from .config import PayrollConfig

PAYROLL_COLUMNS = [
    "No.",
    "Name",
    "Department",
    "Employee_Type",
    "Monthly_Salary",
    "Jumlah_Terlambat",
    "Total_Menit_Telat",
    "Jumlah_Pulang_Cepat",
    "Total_Menit_Pulang_Cepat",
    "Jumlah_Mangkir",
    "Jumlah_Tidak_Absen_Pulang",
    "Potongan_Telat",
    "Potongan_Pulang_Cepat",
    "Potongan_Mangkir",
    "Potongan_Lupa_Pulang_Sales",
    "Potongan_Lain",
    "Total_Potongan",
    "Estimasi_Take_Home",
]


def calculate_payroll(
    employee_summary: pd.DataFrame,
    master: pd.DataFrame,
    cfg: PayrollConfig,
) -> pd.DataFrame:
    if employee_summary.empty:
        return pd.DataFrame(columns=PAYROLL_COLUMNS)

    cols = [
        "No.", "Name", "Department", "Employee_Type",
        "Terlambat", "Total_Menit_Telat",
        "Pulang_Cepat", "Total_Menit_Pulang_Cepat", "Mangkir",
    ]
    df = employee_summary[[c for c in cols if c in employee_summary.columns]].copy()
    df = df.rename(
        columns={
            "Terlambat": "Jumlah_Terlambat",
            "Pulang_Cepat": "Jumlah_Pulang_Cepat",
            "Mangkir": "Jumlah_Mangkir",
        }
    )

    if "Tidak_Absen_Pulang" in employee_summary.columns:
        df["Jumlah_Tidak_Absen_Pulang"] = employee_summary["Tidak_Absen_Pulang"].astype(int)
    else:
        df["Jumlah_Tidak_Absen_Pulang"] = 0

    if master is not None and not master.empty:
        salary_map = master.set_index("No.")["Monthly_Salary"]
        df["Monthly_Salary"] = df["No."].map(salary_map)
    else:
        df["Monthly_Salary"] = np.nan

    def rate_or_na(rate: float | None, quantity: pd.Series) -> pd.Series:
        if rate is None:
            return pd.Series([pd.NA] * len(quantity), index=quantity.index, dtype="Float64")
        return (quantity.astype(float) * rate).round(0).astype("Float64")

    # 1. Potongan Keterlambatan: Opsi per_occurrence (per 1x kejadian) vs per_minute
    if cfg.late_deduction_mode == "per_occurrence":
        df["Potongan_Telat"] = rate_or_na(cfg.late_deduction_per_occurrence, df["Jumlah_Terlambat"])
        active_late_rate = cfg.late_deduction_per_occurrence
    else:
        df["Potongan_Telat"] = rate_or_na(cfg.late_deduction_per_minute, df["Total_Menit_Telat"])
        active_late_rate = cfg.late_deduction_per_minute

    df["Potongan_Pulang_Cepat"] = rate_or_na(
        cfg.early_leave_deduction_per_minute, df["Total_Menit_Pulang_Cepat"]
    )
    df["Potongan_Mangkir"] = rate_or_na(cfg.absence_deduction_per_day, df["Jumlah_Mangkir"])

    # 2. Potongan Khusus Tim Sales: per 1x tidak absen pulang
    if cfg.sales_no_clock_out_deduction is None:
        df["Potongan_Lupa_Pulang_Sales"] = pd.Series([pd.NA] * len(df), dtype="Float64")
    else:
        is_sales = df["Employee_Type"].astype(str).str.upper() == "SALES"
        series = pd.Series([0.0] * len(df), index=df.index, dtype="Float64")
        if is_sales.any():
            sales_deduct = (
                df.loc[is_sales, "Jumlah_Tidak_Absen_Pulang"].astype(float)
                * cfg.sales_no_clock_out_deduction
            ).round(0)
            series.loc[is_sales] = sales_deduct.astype("Float64")
        df["Potongan_Lupa_Pulang_Sales"] = series

    # "Potongan Lain" is inherently ad-hoc (per Section 7 no formula is
    # given) and is intentionally left blank for HR to fill in manually.
    df["Potongan_Lain"] = pd.Series([pd.NA] * len(df), dtype="Float64")

    configured = [
        c
        for c, rate in (
            ("Potongan_Telat", active_late_rate),
            ("Potongan_Pulang_Cepat", cfg.early_leave_deduction_per_minute),
            ("Potongan_Mangkir", cfg.absence_deduction_per_day),
            ("Potongan_Lupa_Pulang_Sales", cfg.sales_no_clock_out_deduction),
        )
        if rate is not None
    ]
    if configured:
        df["Total_Potongan"] = df[configured].sum(axis=1, skipna=True)
    else:
        df["Total_Potongan"] = pd.Series([pd.NA] * len(df), dtype="Float64")

    df["Estimasi_Take_Home"] = df["Monthly_Salary"] - df["Total_Potongan"]

    return df[PAYROLL_COLUMNS].sort_values(["Department", "Name"], kind="stable").reset_index(drop=True)
