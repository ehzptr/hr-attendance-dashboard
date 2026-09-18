"""Attendance analytics: employee KPIs, department roll-up, daily trend."""

from __future__ import annotations

from typing import Optional

import polars as pl

from ..config import ScoreConfig


def calculate_employee_summary(daily: pl.DataFrame, score_cfg: ScoreConfig) -> pl.DataFrame:
    """Employee-level KPI table (Section 6).

    ``Attendance_Score`` counts one "issue day" per calendar day that has
    any anomaly (late, early leave, forgotten punch, unscheduled scan, ...),
    so a single day is never penalised twice even if several things went
    wrong on it. This mirrors how compliance is judged in HR practice: one
    problematic day is one problematic day.
    """
    if daily.is_empty():
        return pl.DataFrame()

    df_calc = daily
    if "Tidak_Absen_Pulang_Flag" not in df_calc.columns:
        if "Status_Pulang" in df_calc.columns:
            df_calc = df_calc.with_columns(
                Tidak_Absen_Pulang_Flag=pl.when(pl.col("Status_Pulang") == "Lupa Absen Pulang")
                .then(1)
                .otherwise(0)
                .cast(pl.Int64)
            )
        else:
            df_calc = df_calc.with_columns(Tidak_Absen_Pulang_Flag=pl.lit(0, dtype=pl.Int64))

    if "Overtime_Hours" not in df_calc.columns:
        df_calc = df_calc.with_columns(Overtime_Hours=pl.lit(0.0, dtype=pl.Float64))

    summary = df_calc.group_by(["No.", "Name", "Department", "Employee_Type"], maintain_order=True).agg(
        Total_Hari=pl.len(),
        Hari_Kerja=pl.col("Is_Working_Day").cast(pl.Int64).sum(),
        Hadir=pl.col("Present_Flag").cast(pl.Int64).sum(),
        Terlambat=pl.col("Late_Flag").cast(pl.Int64).sum(),
        Total_Menit_Telat=pl.col("Menit_Telat").cast(pl.Int64).sum(),
        Pulang_Cepat=pl.col("Early_Leave_Flag").cast(pl.Int64).sum(),
        Total_Menit_Pulang_Cepat=pl.col("Menit_Pulang_Cepat").cast(pl.Int64).sum(),
        Mangkir=pl.col("Absent_Flag").cast(pl.Int64).sum(),
        Lupa_Absen=pl.col("Forgot_Punch_Flag").cast(pl.Int64).sum(),
        Tidak_Absen_Pulang=pl.col("Tidak_Absen_Pulang_Flag").cast(pl.Int64).sum(),
        Total_Jam_Lembur=pl.col("Overtime_Hours").cast(pl.Float64).sum(),
        Anomali=pl.col("Is_Anomali").cast(pl.Int64).sum(),
    )

    summary = summary.with_columns(
        **{
            "Attendance_Rate_%": pl.when(pl.col("Hari_Kerja") > 0)
            .then(pl.col("Hadir") / pl.col("Hari_Kerja") * 100.0)
            .otherwise(100.0),
            "Punctuality_Rate_%": pl.when(pl.col("Hari_Kerja") > 0)
            .then(
                (
                    (pl.col("Hari_Kerja") - pl.col("Terlambat") - pl.col("Mangkir"))
                    / pl.col("Hari_Kerja")
                    * 100.0
                ).clip(0.0, 100.0)
            )
            .otherwise(100.0),
        }
    )

    work = df_calc.filter(pl.col("Is_Working_Day"))
    if not work.is_empty():
        issue_days = (
            work.group_by(["No.", "Tanggal"])
            .agg(pl.col("Is_Anomali").max())
            .group_by("No.")
            .agg(pl.col("Is_Anomali").sum().alias("_Issue_Days"))
        )
        summary = summary.join(issue_days, on="No.", how="left")
        summary = summary.with_columns(_Issue_Days=pl.col("_Issue_Days").fill_null(0))
    else:
        summary = summary.with_columns(_Issue_Days=pl.lit(0, dtype=pl.Int64))

    summary = summary.with_columns(
        Attendance_Score=pl.when(pl.col("Hari_Kerja") > 0)
        .then(((1.0 - pl.col("_Issue_Days") / pl.col("Hari_Kerja")) * 100.0).clip(0.0, 100.0))
        .otherwise(100.0)
    ).drop("_Issue_Days")

    classifications = [score_cfg.classify(s) for s in summary["Attendance_Score"].to_list()]
    summary = summary.with_columns(HR_Classification=pl.Series("HR_Classification", classifications, dtype=pl.String))

    return summary.sort(["Attendance_Score", "Department", "Name"])


def calculate_department_summary(daily: pl.DataFrame) -> pl.DataFrame:
    """Department-level roll-up KPI table."""
    if daily.is_empty():
        return pl.DataFrame()

    work = daily.filter(pl.col("Is_Working_Day"))
    if work.is_empty():
        return pl.DataFrame()

    out = work.group_by("Department", maintain_order=True).agg(
        Employee=pl.col("No.").n_unique(),
        Working_Days=pl.col("Tanggal").n_unique(),
        Total_Man_Days=pl.col("Work_Day_Flag").cast(pl.Int64).sum(),
        Present=pl.col("Present_Flag").cast(pl.Int64).sum(),
        Late=pl.col("Late_Flag").cast(pl.Int64).sum(),
        Absent=pl.col("Absent_Flag").cast(pl.Int64).sum(),
        Early_Leave=pl.col("Early_Leave_Flag").cast(pl.Int64).sum(),
        Late_Minutes=pl.col("Menit_Telat").cast(pl.Int64).sum(),
        Anomalies=pl.col("Is_Anomali").cast(pl.Int64).sum(),
    )

    out = out.with_columns(
        **{
            "Attendance_Rate_%": pl.when(pl.col("Total_Man_Days") > 0)
            .then(pl.col("Present") / pl.col("Total_Man_Days") * 100.0)
            .otherwise(100.0)
        }
    )

    issue_days = (
        work.group_by(["Department", "No.", "Tanggal"])
        .agg(pl.col("Is_Anomali").max())
        .group_by("Department")
        .agg(pl.col("Is_Anomali").sum().alias("_Total_Issue_Days"))
    )
    out = out.join(issue_days, on="Department", how="left")
    out = out.with_columns(_Total_Issue_Days=pl.col("_Total_Issue_Days").fill_null(0))
    out = out.with_columns(
        **{
            "Attendance_Score_%": pl.when(pl.col("Total_Man_Days") > 0)
            .then(((1.0 - pl.col("_Total_Issue_Days") / pl.col("Total_Man_Days")) * 100.0).clip(0.0, 100.0))
            .otherwise(100.0)
        }
    ).drop("_Total_Issue_Days")

    return out.sort("Attendance_Score_%")


def calculate_daily_trend(daily: pl.DataFrame) -> pl.DataFrame:
    """Daily attendance trend table."""
    if daily.is_empty():
        return pl.DataFrame()

    work = daily.filter(pl.col("Is_Working_Day"))
    if work.is_empty():
        return pl.DataFrame()

    trend = work.group_by("Tanggal", maintain_order=True).agg(
        Work_Days=pl.col("Work_Day_Flag").cast(pl.Int64).sum(),
        Present=pl.col("Present_Flag").cast(pl.Int64).sum(),
        Late=pl.col("Late_Flag").cast(pl.Int64).sum(),
        Absent=pl.col("Absent_Flag").cast(pl.Int64).sum(),
        Early_Leave=pl.col("Early_Leave_Flag").cast(pl.Int64).sum(),
        Late_Minutes=pl.col("Menit_Telat").cast(pl.Int64).sum(),
    )
    trend = trend.with_columns(
        **{
            "Attendance_Rate_%": pl.when(pl.col("Work_Days") > 0)
            .then(pl.col("Present") / pl.col("Work_Days") * 100.0)
            .otherwise(100.0)
        }
    )
    return trend.sort("Tanggal")


def derive_hr_insights(summary: pl.DataFrame, dept: pl.DataFrame) -> list[str]:
    """Derive automated textual HR insights and priorities from analytics tables."""
    insights: list[str] = []
    if summary.is_empty():
        return ["Belum ada data yang dapat dianalisis."]

    worst = summary.sort("Attendance_Score").head(5)
    if not worst.is_empty():
        emp = worst.row(0, named=True)
        insights.append(
            f"Prioritas HR: {emp['Name']} ({emp['Department']}) "
            f"memiliki attendance score {emp['Attendance_Score']:.1f}%."
        )

    late_minutes = int(summary["Total_Menit_Telat"].sum()) if "Total_Menit_Telat" in summary.columns else 0
    if late_minutes:
        insights.append(
            f"Total keterlambatan mencapai {late_minutes:,} menit; "
            "pertimbangkan coaching berdasarkan pola departemen/hari."
        )

    if not dept.is_empty():
        d = dept.row(0, named=True)
        insights.append(
            f"Departemen dengan attendance score terendah saat ini adalah "
            f"{d['Department']} ({d['Attendance_Score_%']:.1f}%)."
        )

    high_absence = summary.filter(pl.col("Mangkir") > 0) if "Mangkir" in summary.columns else pl.DataFrame()
    if len(high_absence):
        insights.append(f"{len(high_absence)} karyawan memiliki setidaknya satu hari mangkir.")

    forgot = summary.filter(pl.col("Lupa_Absen") > 0) if "Lupa_Absen" in summary.columns else pl.DataFrame()
    if len(forgot):
        insights.append(f"{len(forgot)} karyawan memiliki kejadian lupa absen masuk/pulang.")

    if "Tidak_Absen_Pulang" in summary.columns and "Employee_Type" in summary.columns:
        sales_subset = summary.filter(pl.col("Employee_Type") == "SALES")
        sales_no_out = int(sales_subset["Tidak_Absen_Pulang"].sum()) if not sales_subset.is_empty() else 0
        if sales_no_out > 0:
            insights.append(
                f"Anomali Sales: Terdeteksi {sales_no_out} kejadian tidak absen pulang pada tim Sales "
                "(perlu verifikasi tugas lapangan vs mangkir/pulang awal)."
            )

    return insights
