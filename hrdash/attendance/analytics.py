"""Attendance analytics: employee KPIs, department roll-up, daily trend."""

from __future__ import annotations

import numpy as np
import pandas as pd

from ..config import ScoreConfig


def calculate_employee_summary(daily: pd.DataFrame, score_cfg: ScoreConfig) -> pd.DataFrame:
    """Employee-level KPI table (Section 6).

    ``Attendance_Score`` counts one "issue day" per calendar day that has
    any anomaly (late, early leave, forgotten punch, unscheduled scan, ...),
    so a single day is never penalised twice even if several things went
    wrong on it. This mirrors how compliance is judged in HR practice: one
    problematic day is one problematic day.
    """
    if daily.empty:
        return pd.DataFrame()

    work = daily[daily["Is_Working_Day"]].copy()

    df_calc = daily.copy()
    if "Tidak_Absen_Pulang_Flag" not in df_calc.columns:
        df_calc["Tidak_Absen_Pulang_Flag"] = (
            df_calc["Status_Pulang"].eq("Lupa Absen Pulang").astype(int)
            if "Status_Pulang" in df_calc.columns
            else 0
        )
    if "Overtime_Hours" not in df_calc.columns:
        df_calc["Overtime_Hours"] = 0.0

    summary = df_calc.groupby(["No.", "Name", "Department", "Employee_Type"], as_index=False).agg(
        Total_Hari=("Tanggal", "size"),
        Hari_Kerja=("Is_Working_Day", "sum"),
        Hadir=("Present_Flag", "sum"),
        Terlambat=("Late_Flag", "sum"),
        Total_Menit_Telat=("Menit_Telat", "sum"),
        Pulang_Cepat=("Early_Leave_Flag", "sum"),
        Total_Menit_Pulang_Cepat=("Menit_Pulang_Cepat", "sum"),
        Mangkir=("Absent_Flag", "sum"),
        Lupa_Absen=("Forgot_Punch_Flag", "sum"),
        Tidak_Absen_Pulang=("Tidak_Absen_Pulang_Flag", "sum"),
        Total_Jam_Lembur=("Overtime_Hours", "sum"),
        Anomali=("Is_Anomali", "sum"),
    )

    summary["Attendance_Rate_%"] = np.where(
        summary["Hari_Kerja"] > 0, summary["Hadir"] / summary["Hari_Kerja"] * 100, 100.0
    )
    summary["Punctuality_Rate_%"] = np.where(
        summary["Hari_Kerja"] > 0,
        (summary["Hari_Kerja"] - summary["Terlambat"] - summary["Mangkir"]) / summary["Hari_Kerja"] * 100,
        100.0,
    ).clip(min=0, max=100)

    issue_days = (
        work.groupby(["No.", "Tanggal"], as_index=False)["Is_Anomali"].max()
        .groupby("No.", as_index=False)["Is_Anomali"].sum()
        .rename(columns={"Is_Anomali": "_Issue_Days"})
    )
    summary = summary.merge(issue_days, on="No.", how="left")
    summary["_Issue_Days"] = summary["_Issue_Days"].fillna(0)
    summary["Attendance_Score"] = np.clip(
        np.where(
            summary["Hari_Kerja"] > 0,
            (1 - summary["_Issue_Days"] / summary["Hari_Kerja"]) * 100,
            100.0,
        ),
        0,
        100,
    )
    summary = summary.drop(columns=["_Issue_Days"])

    summary["HR_Classification"] = summary["Attendance_Score"].map(score_cfg.classify)

    return summary.sort_values(
        ["Attendance_Score", "Department", "Name"], ascending=[True, True, True], kind="stable"
    ).reset_index(drop=True)


def calculate_department_summary(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    work = daily[daily["Is_Working_Day"]].copy()
    if work.empty:
        return pd.DataFrame()

    out = work.groupby("Department", as_index=False).agg(
        Employee=("No.", "nunique"),
        Work_Days=("Work_Day_Flag", "sum"),
        Present=("Present_Flag", "sum"),
        Late=("Late_Flag", "sum"),
        Absent=("Absent_Flag", "sum"),
        Early_Leave=("Early_Leave_Flag", "sum"),
        Late_Minutes=("Menit_Telat", "sum"),
        Anomalies=("Is_Anomali", "sum"),
    )
    out["Attendance_Rate_%"] = np.where(out["Work_Days"] > 0, out["Present"] / out["Work_Days"] * 100, 100)

    issue_days = (
        work.assign(_Issue_Day=work["Is_Anomali"].astype(int))
        .groupby(["Department", "Tanggal"], as_index=False)["_Issue_Day"].max()
        .groupby("Department", as_index=False)["_Issue_Day"].sum()
        .rename(columns={"_Issue_Day": "_Issue_Days"})
    )
    out = out.merge(issue_days, on="Department", how="left")
    out["_Issue_Days"] = out["_Issue_Days"].fillna(0)
    out["Attendance_Score_%"] = (
        100 - out["_Issue_Days"] / out["Work_Days"].replace(0, np.nan) * 100
    ).fillna(100).clip(lower=0, upper=100)
    out = out.drop(columns=["_Issue_Days"])
    return out.sort_values("Attendance_Score_%").reset_index(drop=True)


def calculate_daily_trend(daily: pd.DataFrame) -> pd.DataFrame:
    if daily.empty:
        return pd.DataFrame()

    work = daily[daily["Is_Working_Day"]].copy()
    trend = work.groupby("Tanggal", as_index=False).agg(
        Work_Days=("Work_Day_Flag", "sum"),
        Present=("Present_Flag", "sum"),
        Late=("Late_Flag", "sum"),
        Absent=("Absent_Flag", "sum"),
        Early_Leave=("Early_Leave_Flag", "sum"),
        Late_Minutes=("Menit_Telat", "sum"),
    )
    trend["Attendance_Rate_%"] = np.where(trend["Work_Days"] > 0, trend["Present"] / trend["Work_Days"] * 100, 100)
    return trend


def derive_hr_insights(summary: pd.DataFrame, dept: pd.DataFrame) -> list[str]:
    insights: list[str] = []
    if summary.empty:
        return ["Belum ada data yang dapat dianalisis."]

    worst = summary.nsmallest(5, "Attendance_Score")
    if not worst.empty:
        employee = worst.iloc[0]
        insights.append(
            f"Prioritas HR: {employee['Name']} ({employee['Department']}) "
            f"memiliki attendance score {employee['Attendance_Score']:.1f}%."
        )

    late_minutes = int(summary["Total_Menit_Telat"].sum())
    if late_minutes:
        insights.append(
            f"Total keterlambatan mencapai {late_minutes:,} menit; "
            "pertimbangkan coaching berdasarkan pola departemen/hari."
        )

    if not dept.empty:
        d = dept.iloc[0]
        insights.append(
            f"Departemen dengan attendance score terendah saat ini adalah "
            f"{d['Department']} ({d['Attendance_Score_%']:.1f}%)."
        )

    high_absence = summary[summary["Mangkir"] > 0]
    if len(high_absence):
        insights.append(f"{len(high_absence)} karyawan memiliki setidaknya satu hari mangkir.")

    forgot = summary[summary["Lupa_Absen"] > 0]
    if len(forgot):
        insights.append(f"{len(forgot)} karyawan memiliki kejadian lupa absen masuk/pulang.")

    if "Tidak_Absen_Pulang" in summary.columns and "Employee_Type" in summary.columns:
        sales_no_out = int(
            summary.loc[summary["Employee_Type"] == "SALES", "Tidak_Absen_Pulang"].sum()
        )
        if sales_no_out > 0:
            insights.append(
                f"Anomali Sales: Terdeteksi {sales_no_out} kejadian tidak absen pulang pada tim Sales "
                "(perlu verifikasi tugas lapangan vs mangkir/pulang awal)."
            )

    return insights
