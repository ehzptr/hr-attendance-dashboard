"""
Payroll deduction and payslip data engine (Section 7).

The engine only ever computes a monetary deduction when HR has explicitly
supplied the corresponding rate in ``PayrollConfig``. When a rate is
missing, the column is left as ``None`` / null (rendered as "Belum Diatur" in the
Excel export) instead of silently defaulting to zero or to an invented
number — the spec is explicit that HR-defined nominal values must never be
assumed.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import polars as pl

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


@dataclass
class PayslipData:
    """Complete employee payslip record for PDF and Excel rendering."""
    # Employee Identity
    employee_id: str
    name: str
    department: str
    position: str = "Staff"
    employee_type: str = "OFFICE"
    period_label: str = ""
    join_date: str = "-"

    # Earnings (Pendapatan)
    gaji_pokok: float = 0.0
    lembur_hours: float = 0.0
    lembur_pay: float = 0.0
    tunjangan_jabatan: float = 0.0
    uang_makan: float = 0.0
    tunjangan_transport: float = 0.0
    tunjangan_lain: float = 0.0
    total_pendapatan: float = 0.0

    # Deductions (Potongan)
    potongan_mangkir: float = 0.0
    potongan_telat: float = 0.0
    potongan_pulang_cepat: float = 0.0
    potongan_lupa_pulang_sales: float = 0.0
    potongan_bpjs_kesehatan: float = 0.0
    potongan_bpjs_tk_jht: float = 0.0
    potongan_bpjs_tk_jp: float = 0.0
    potongan_pph21: float = 0.0
    potongan_lain: float = 0.0
    total_potongan: float = 0.0

    # Net Pay
    take_home_pay: float = 0.0

    # Attendance Summary (Rangkuman Informasi Kehadiran)
    hari_kehadiran: int = 0
    hari_mangkir: int = 0
    hari_cuti: int = 0
    hari_izin: int = 0
    hari_sakit: int = 0
    kali_terlambat: int = 0
    menit_terlambat: int = 0
    kali_pulang_cepat: int = 0
    kali_lupa_absen: int = 0
    total_jam_lembur: float = 0.0


def calculate_payroll(
    employee_summary: pl.DataFrame,
    master: Optional[pl.DataFrame],
    cfg: PayrollConfig,
) -> pl.DataFrame:
    """Calculate payroll deductions and estimated take-home pay."""
    if employee_summary.is_empty():
        schema = {c: pl.String if c in ("No.", "Name", "Department", "Employee_Type") else pl.Float64 for c in PAYROLL_COLUMNS}
        return pl.DataFrame(schema=schema)

    df_summary = employee_summary
    cols_needed = [
        "No.", "Name", "Department", "Employee_Type",
        "Terlambat", "Total_Menit_Telat",
        "Pulang_Cepat", "Total_Menit_Pulang_Cepat", "Mangkir",
    ]
    present_cols = [c for c in cols_needed if c in df_summary.columns]
    df = df_summary.select(present_cols)

    rename_map = {
        "Terlambat": "Jumlah_Terlambat",
        "Pulang_Cepat": "Jumlah_Pulang_Cepat",
        "Mangkir": "Jumlah_Mangkir",
    }
    df = df.rename({k: v for k, v in rename_map.items() if k in df.columns})

    for col in ["Jumlah_Terlambat", "Total_Menit_Telat", "Jumlah_Pulang_Cepat", "Total_Menit_Pulang_Cepat", "Jumlah_Mangkir"]:
        if col not in df.columns:
            df = df.with_columns(pl.lit(0, dtype=pl.Int64).alias(col))
        else:
            df = df.with_columns(pl.col(col).cast(pl.Int64))

    if "Tidak_Absen_Pulang" in df_summary.columns:
        df = df.with_columns(Jumlah_Tidak_Absen_Pulang=df_summary["Tidak_Absen_Pulang"].cast(pl.Int64))
    else:
        df = df.with_columns(Jumlah_Tidak_Absen_Pulang=pl.lit(0, dtype=pl.Int64))

    if master is not None and not master.is_empty() and "No." in master.columns and "Monthly_Salary" in master.columns:
        salary_sub = master.select(["No.", "Monthly_Salary"]).unique(subset=["No."])
        df = df.join(salary_sub, on="No.", how="left")
        df = df.with_columns(Monthly_Salary=pl.col("Monthly_Salary").cast(pl.Float64))
    else:
        df = df.with_columns(Monthly_Salary=pl.lit(None, dtype=pl.Float64))

    # 1. Potongan Keterlambatan: per_occurrence vs per_minute
    if cfg.late_deduction_mode == "per_occurrence":
        active_late_rate = cfg.late_deduction_per_occurrence
        if active_late_rate is not None:
            df = df.with_columns(
                Potongan_Telat=(pl.col("Jumlah_Terlambat").cast(pl.Float64) * active_late_rate).round(0)
            )
        else:
            df = df.with_columns(Potongan_Telat=pl.lit(None, dtype=pl.Float64))
    else:
        active_late_rate = cfg.late_deduction_per_minute
        if active_late_rate is not None:
            df = df.with_columns(
                Potongan_Telat=(pl.col("Total_Menit_Telat").cast(pl.Float64) * active_late_rate).round(0)
            )
        else:
            df = df.with_columns(Potongan_Telat=pl.lit(None, dtype=pl.Float64))

    # 2. Potongan Pulang Cepat
    if cfg.early_leave_deduction_per_minute is not None:
        df = df.with_columns(
            Potongan_Pulang_Cepat=(
                pl.col("Total_Menit_Pulang_Cepat").cast(pl.Float64) * cfg.early_leave_deduction_per_minute
            ).round(0)
        )
    else:
        df = df.with_columns(Potongan_Pulang_Cepat=pl.lit(None, dtype=pl.Float64))

    # 3. Potongan Mangkir
    if cfg.absence_deduction_per_day is not None:
        df = df.with_columns(
            Potongan_Mangkir=(
                pl.col("Jumlah_Mangkir").cast(pl.Float64) * cfg.absence_deduction_per_day
            ).round(0)
        )
    else:
        df = df.with_columns(Potongan_Mangkir=pl.lit(None, dtype=pl.Float64))

    # 4. Potongan Khusus Tim Sales: per 1x tidak absen pulang
    if cfg.sales_no_clock_out_deduction is not None:
        df = df.with_columns(
            Potongan_Lupa_Pulang_Sales=pl.when(
                pl.col("Employee_Type").cast(pl.String).str.to_uppercase() == "SALES"
            )
            .then(
                (pl.col("Jumlah_Tidak_Absen_Pulang").cast(pl.Float64) * cfg.sales_no_clock_out_deduction).round(0)
            )
            .otherwise(0.0)
        )
    else:
        df = df.with_columns(Potongan_Lupa_Pulang_Sales=pl.lit(None, dtype=pl.Float64))

    # 5. Potongan Lain (Manual)
    df = df.with_columns(Potongan_Lain=pl.lit(None, dtype=pl.Float64))

    # 6. Total Potongan
    configured_rates = [
        ("Potongan_Telat", active_late_rate),
        ("Potongan_Pulang_Cepat", cfg.early_leave_deduction_per_minute),
        ("Potongan_Mangkir", cfg.absence_deduction_per_day),
        ("Potongan_Lupa_Pulang_Sales", cfg.sales_no_clock_out_deduction),
    ]
    active_cols = [col for col, rate in configured_rates if rate is not None]

    if active_cols:
        sum_expr = pl.sum_horizontal([pl.col(c).fill_null(0.0) for c in active_cols])
        df = df.with_columns(Total_Potongan=sum_expr)
    else:
        df = df.with_columns(Total_Potongan=pl.lit(None, dtype=pl.Float64))

    # 7. Estimasi Take Home
    df = df.with_columns(
        Estimasi_Take_Home=pl.when(pl.col("Monthly_Salary").is_not_null())
        .then(pl.col("Monthly_Salary") - pl.col("Total_Potongan").fill_null(0.0))
        .otherwise(None)
    )

    return df.select(PAYROLL_COLUMNS).sort(["Department", "Name"])


def build_payslip_records(
    payroll_df: pl.DataFrame,
    employee_summary_df: pl.DataFrame,
    daily_df: pl.DataFrame,
    master_df: Optional[pl.DataFrame] = None,
    overtime_hourly_rate: float = 0.0,
    period_label: str = "",
) -> list[PayslipData]:
    """Compile structured PayslipData objects for each employee."""
    if payroll_df.is_empty():
        return []

    # Map positions and join dates from master if present
    pos_map = {}
    join_map = {}
    if master_df is not None and not master_df.is_empty():
        for r in master_df.iter_rows(named=True):
            eid = str(r.get("No.", "")).strip()
            if eid:
                if "Position" in r and r["Position"]:
                    pos_map[eid] = str(r["Position"])
                if "Join_Date" in r and r["Join_Date"]:
                    join_map[eid] = str(r["Join_Date"])

    # Leave breakdown from daily attendance
    leave_breakdown = {}
    if not daily_df.is_empty() and "No." in daily_df.columns:
        for r in daily_df.iter_rows(named=True):
            eid = str(r.get("No.", "")).strip()
            if eid not in leave_breakdown:
                leave_breakdown[eid] = {"cuti": 0, "izin": 0, "sakit": 0}
            
            # Check leave status
            sm = str(r.get("Status_Masuk", "")).lower()
            sp = str(r.get("Status_Pulang", "")).lower()
            lt = str(r.get("Leave_Type", "")).lower()
            text_check = f"{sm} {sp} {lt}"
            if "sakit" in text_check:
                leave_breakdown[eid]["sakit"] += 1
            elif "cuti" in text_check:
                leave_breakdown[eid]["cuti"] += 1
            elif "izin" in text_check or "ijin" in text_check or "dispensasi" in text_check:
                leave_breakdown[eid]["izin"] += 1

    summary_map = {}
    if not employee_summary_df.is_empty():
        for r in employee_summary_df.iter_rows(named=True):
            eid = str(r.get("No.", "")).strip()
            if eid:
                summary_map[eid] = r

    payslips = []
    for r in payroll_df.iter_rows(named=True):
        eid = str(r.get("No.", "")).strip()
        name = str(r.get("Name", ""))
        dept = str(r.get("Department", ""))
        etype = str(r.get("Employee_Type", "OFFICE"))

        summ = summary_map.get(eid, {})
        leaves = leave_breakdown.get(eid, {"cuti": 0, "izin": 0, "sakit": 0})

        # Earnings
        salary = float(r.get("Monthly_Salary") or 0.0)
        overtime_hrs = float(summ.get("Total_Jam_Lembur") or 0.0)
        overtime_pay = overtime_hrs * overtime_hourly_rate if overtime_hourly_rate > 0 else 0.0
        total_income = salary + overtime_pay

        # Deductions
        pot_mangkir = float(r.get("Potongan_Mangkir") or 0.0)
        pot_telat = float(r.get("Potongan_Telat") or 0.0)
        pot_pulang_cepat = float(r.get("Potongan_Pulang_Cepat") or 0.0)
        pot_sales = float(r.get("Potongan_Lupa_Pulang_Sales") or 0.0)
        pot_lain = float(r.get("Potongan_Lain") or 0.0)
        total_pot = pot_mangkir + pot_telat + pot_pulang_cepat + pot_sales + pot_lain
        net_pay = max(0.0, total_income - total_pot)

        slip = PayslipData(
            employee_id=eid,
            name=name,
            department=dept,
            position=pos_map.get(eid, "Staff"),
            employee_type=etype,
            period_label=period_label,
            join_date=join_map.get(eid, "-"),
            # Earnings
            gaji_pokok=salary,
            lembur_hours=overtime_hrs,
            lembur_pay=overtime_pay,
            total_pendapatan=total_income,
            # Deductions
            potongan_mangkir=pot_mangkir,
            potongan_telat=pot_telat,
            potongan_pulang_cepat=pot_pulang_cepat,
            potongan_lupa_pulang_sales=pot_sales,
            potongan_lain=pot_lain,
            total_potongan=total_pot,
            take_home_pay=net_pay,
            # Attendance
            hari_kehadiran=int(summ.get("Hadir") or 0),
            hari_mangkir=int(summ.get("Mangkir") or 0),
            hari_cuti=leaves["cuti"],
            hari_izin=leaves["izin"],
            hari_sakit=leaves["sakit"],
            kali_terlambat=int(summ.get("Terlambat") or 0),
            menit_terlambat=int(summ.get("Total_Menit_Telat") or 0),
            kali_pulang_cepat=int(summ.get("Pulang_Cepat") or 0),
            kali_lupa_absen=int(summ.get("Lupa_Absen") or 0),
            total_jam_lembur=overtime_hrs,
        )
        payslips.append(slip)

    return payslips
