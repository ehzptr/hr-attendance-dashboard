"""
Streamlit UI entrypoint.

Thin presentation layer only — all business logic lives in the other
``hrdash`` modules. Run with:

    streamlit run -m hrdash.app
"""

from __future__ import annotations

import datetime as dt
import io
import logging
from datetime import date, datetime, time
from typing import Optional

import plotly.express as px
import polars as pl
import streamlit as st

from .attendance import analytics as att_analytics
from .attendance import engine as att_engine
from .attendance.parser import AttendanceValidationError, clean_approved_leave, clean_attendance_data
from .config import AppConfig, AttendanceConfig, PayrollConfig, ScoreConfig, SalesScoreConfig
from .employee_master import (
    EmployeeMasterError,
    build_employee_master_template,
    load_employee_master,
    merge_master_into_log,
)
from .leave import build_leave_template, load_leave_workbook
from .payroll import calculate_payroll, build_payslip_records, PayslipData
from .reporting.excel_report import ReportData, generate_workbook
from .reporting.payslip import (
    generate_payslip_pdf,
    generate_bulk_payslip_pdf,
    generate_payslip_excel,
    generate_bulk_payslip_excel,
)
from .sales import (
    SalesActivityError,
    build_sales_activity_template,
    calculate_sales_performance,
    clean_sales_activity,
    combine_attendance_and_sales,
)

LOGGER = logging.getLogger("hrdash")
if not LOGGER.handlers:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


def _read_excel_fast(source: object, **kwargs) -> pl.DataFrame:
    try:
        return pl.read_excel(source, engine="calamine", **kwargs)
    except Exception:
        return pl.read_excel(source, **kwargs)


@st.cache_data(show_spinner=False)
def _read_excel_bytes(file_bytes: bytes, **kwargs) -> pl.DataFrame:
    return _read_excel_fast(io.BytesIO(file_bytes), **kwargs)


@st.cache_data(show_spinner=False)
def _read_csv_bytes(file_bytes: bytes, **kwargs) -> pl.DataFrame:
    return pl.read_csv(io.BytesIO(file_bytes), **kwargs)


def _build_sample_attendance() -> bytes:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["Department", "Name", "No.", "Date/Time"])
    ws.append(["OUR COMPANY", "1111", "1111", "21/08/2026 07.34.23"])
    ws.append(["Sales Mobil Baru", "Budi Santoso", "EMP001", "21/08/2026 07.50.12"])
    ws.append(["Sales Mobil Baru", "Budi Santoso", "EMP001", "21/08/2026 17.10.45"])
    ws.append(["HRD & ADMIN", "Ani Lestari", "12021007", "21/08/2026 07.40.16"])
    ws.append(["HRD & ADMIN", "Ani Lestari", "12021007", "21/08/2026 17.24.11"])
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


def _export_pdf(employee_summary: pl.DataFrame, department_summary: pl.DataFrame, insights: list[str]) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Laporan Kehadiran HR", styles["Title"]),
        Spacer(1, 6 * mm),
        Paragraph("Ringkasan Insight HR", styles["Heading2"]),
    ]
    for text in insights:
        story.append(Paragraph(f"\u2022 {text}", styles["BodyText"]))

    story.append(Spacer(1, 5 * mm))
    story.append(Paragraph("Top Prioritas HR", styles["Heading2"]))
    top = employee_summary.sort("Attendance_Score").head(15) if not employee_summary.is_empty() else pl.DataFrame()
    data = [["Karyawan", "Departemen", "Hadir", "Terlambat", "Mangkir", "Attendance Score"]]
    for r in top.iter_rows(named=True):
        score_val = r.get("Attendance_Score", 0.0) or 0.0
        data.append(
            [str(r.get("Name", "")), str(r.get("Department", "")), str(int(r.get("Hadir", 0) or 0)),
             str(int(r.get("Terlambat", 0) or 0)), str(int(r.get("Mangkir", 0) or 0)), f"{score_val:.1f}%"]
        )
    table = Table(data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
            ]
        )
    )
    story.append(table)

    if not department_summary.is_empty():
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("Ringkasan Departemen", styles["Heading2"]))
        ddata = [["Departemen", "Employee", "Working Days", "Total Man Days", "Attendance", "Score", "Anomali"]]
        for r in department_summary.iter_rows(named=True):
            att_rate = r.get("Attendance_Rate_%", 0.0) or 0.0
            att_score = r.get("Attendance_Score_%", 0.0) or 0.0
            ddata.append(
                [
                    str(r.get("Department", "")),
                    str(int(r.get("Employee", 0) or 0)),
                    str(int(r.get("Working_Days", 0) or 0)),
                    str(int(r.get("Total_Man_Days", 0) or 0)),
                    f"{att_rate:.1f}%",
                    f"{att_score:.1f}%",
                    str(int(r.get("Anomalies", 0) or 0)),
                ]
            )
        dtable = Table(ddata, repeatRows=1)
        dtable.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(dtable)

    doc.build(story)
    return buffer.getvalue()


def _metric_grid(metrics: list[tuple[str, str, str]]) -> None:
    for start in range(0, len(metrics), 4):
        batch = metrics[start : start + 4]
        cols = st.columns(4, gap="medium")
        for col, m in zip(cols, batch):
            with col:
                st.metric(label=m[0], value=m[1], help=m[2] or None)


# ---------------------------------------------------------------------------
# Cached Data Pipeline
# ---------------------------------------------------------------------------


@st.cache_data(show_spinner="Memproses kalkulasi absensi, payroll & performa sales...")
def run_attendance_pipeline(
    attendance_bytes: bytes,
    master_bytes: Optional[bytes],
    leave_bytes: Optional[bytes],
    leave_is_csv: bool,
    sales_bytes: Optional[bytes],
    sales_is_csv: bool,
    clock_in_str: str,
    weekday_out_str: str,
    saturday_out_str: str,
    grace_minutes: int,
    saturday_working: bool,
    sunday_off: bool,
    start_date_val: Optional[date],
    end_date_val: Optional[date],
    holidays_tuple: tuple[str, ...],
    score_thresholds: tuple[float, float, float],
    payroll_args: dict,
    sales_weights: dict,
    overtime_hourly_rate: float,
) -> dict:
    """Pure cached pipeline execution."""
    raw = _read_excel_bytes(attendance_bytes)
    clean, rejected, quality = clean_attendance_data(raw)
    if clean.is_empty():
        raise AttendanceValidationError("Tidak ada baris valid setelah proses cleaning.")

    master = pl.DataFrame()
    if master_bytes:
        master = load_employee_master(_read_excel_bytes(master_bytes))

    enriched, unmapped = merge_master_into_log(clean, master)

    leave_result = None
    approved_leave = None
    if leave_bytes:
        if leave_is_csv:
            leave_raw = _read_csv_bytes(leave_bytes)
            approved_leave = clean_approved_leave(leave_raw)
        else:
            leave_result = load_leave_workbook(io.BytesIO(leave_bytes))

    holidays = set()
    for line in holidays_tuple:
        line = line.strip()
        if line:
            try:
                holidays.add(datetime.strptime(line[:10], "%Y-%m-%d").date())
            except ValueError:
                pass

    clock_in_t = time.fromisoformat(clock_in_str)
    weekday_out_t = time.fromisoformat(weekday_out_str)
    saturday_out_t = time.fromisoformat(saturday_out_str)

    cfg = AppConfig(
        attendance=AttendanceConfig(
            clock_in=clock_in_t,
            weekday_clock_out=weekday_out_t,
            saturday_clock_out=saturday_out_t,
            grace_minutes=grace_minutes,
            saturday_is_working=saturday_working,
            sunday_is_off=sunday_off,
            start_date=start_date_val,
            end_date=end_date_val,
        ),
        score=ScoreConfig(
            excellent_threshold=score_thresholds[0],
            good_threshold=score_thresholds[1],
            watch_threshold=score_thresholds[2],
        ),
        payroll=PayrollConfig(**payroll_args),
        sales_score=SalesScoreConfig(**sales_weights),
    )

    daily = att_engine.build_daily_attendance(
        enriched, cfg.attendance, holidays, approved_leave=approved_leave, leave_result=leave_result
    )
    employee_summary = att_analytics.calculate_employee_summary(daily, cfg.score)
    department_summary = att_analytics.calculate_department_summary(daily)
    trend = att_analytics.calculate_daily_trend(daily)
    anomalies = daily.filter(pl.col("Is_Anomali") == 1) if not daily.is_empty() else pl.DataFrame()
    insights = att_analytics.derive_hr_insights(employee_summary, department_summary)
    payroll = calculate_payroll(employee_summary, master, cfg.payroll)

    sales_activity = pl.DataFrame()
    sales_performance = pl.DataFrame()
    combined_view = pl.DataFrame()
    if sales_bytes:
        sales_raw = _read_csv_bytes(sales_bytes) if sales_is_csv else _read_excel_bytes(sales_bytes)
        sales_activity = clean_sales_activity(sales_raw)
        sales_performance = calculate_sales_performance(sales_activity, master, cfg.sales_score)
        combined_view = combine_attendance_and_sales(employee_summary, sales_performance)

    period_label = ""
    if not daily.is_empty() and "Tanggal" in daily.columns:
        period_label = f"{daily['Tanggal'].min()} s/d {daily['Tanggal'].max()}"

    payslip_records = build_payslip_records(
        payroll,
        employee_summary,
        daily,
        master,
        overtime_hourly_rate=overtime_hourly_rate,
        period_label=period_label,
    )

    return {
        "cfg": cfg,
        "clean": clean,
        "rejected": rejected,
        "quality": quality,
        "master": master,
        "unmapped": unmapped,
        "daily": daily,
        "employee_summary": employee_summary,
        "department_summary": department_summary,
        "trend": trend,
        "anomalies": anomalies,
        "insights": insights,
        "payroll": payroll,
        "sales_activity": sales_activity,
        "sales_performance": sales_performance,
        "combined_view": combined_view,
        "payslips": payslip_records,
        "period_label": period_label,
    }


# ---------------------------------------------------------------------------
# UI Components
# ---------------------------------------------------------------------------


def render_sidebar():
    with st.sidebar:
        st.header("1. Data")
        attendance_file = st.file_uploader("Export mesin absensi (wajib)", type=["xlsx", "xls"])
        master_file = st.file_uploader("Employee Master (opsional)", type=["xlsx", "xls"])
        leave_file = st.file_uploader(
            "Upload daftar cuti, lembur & kegiatan bersama (opsional)",
            type=["xlsx", "xls", "csv"],
            help="Mendukung multi-sheet Excel: (1) Cuti_Izin, (2) Lembur_Karyawan, (3) Kegiatan_Bersama.",
        )
        sales_file = st.file_uploader("Sales Activity (opsional)", type=["xlsx", "xls", "csv"])

        st.download_button(
            "\u2B07\uFE0F Template absensi",
            data=_build_sample_attendance(),
            file_name="attendance_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )
        st.download_button(
            "\U0001F4CB Template Employee Master",
            data=build_employee_master_template(),
            file_name="employee_master_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )
        st.download_button(
            "\U0001F4C5 Template Cuti, Lembur & Kegiatan Bersama",
            data=build_leave_template(),
            file_name="template_cuti_lembur_kegiatan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )
        st.download_button(
            "\U0001F4C8 Template Sales Activity",
            data=build_sales_activity_template(),
            file_name="sales_activity_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
        )

        st.header("2. Jadwal & Toleransi")
        clock_in = st.time_input("Jam masuk", time(8, 0))
        weekday_out = st.time_input("Pulang Senin\u2013Jumat", time(17, 0))
        saturday_out = st.time_input("Pulang Sabtu", time(14, 0))
        grace = st.number_input("Toleransi keterlambatan (menit)", min_value=0, max_value=120, value=10)
        saturday_working = st.checkbox("Sabtu hari kerja", value=True)
        sunday_off = st.checkbox("Minggu libur", value=True)

        st.subheader("Periode Perhitungan")
        use_custom_period = st.checkbox("Gunakan periode kustom", value=False)
        start_date_val = end_date_val = None
        if use_custom_period:
            start_date_val = st.date_input("Tanggal mulai", value=None)
            end_date_val = st.date_input("Tanggal akhir", value=None)

        st.subheader("Kalender Libur")
        holidays_text = st.text_area(
            "Hari libur (YYYY-MM-DD, satu per baris)", placeholder="2026-08-17\n2026-12-25"
        )

        with st.expander("3. Klasifikasi Skor Kehadiran"):
            excellent_t = st.number_input("Ambang EXCELLENT", value=90.0, min_value=0.0, max_value=100.0)
            good_t = st.number_input("Ambang GOOD", value=80.0, min_value=0.0, max_value=100.0)
            watch_t = st.number_input("Ambang WATCH", value=70.0, min_value=0.0, max_value=100.0)

        with st.expander("4. Tarif Potongan Payroll"):
            st.caption("Kosongkan jika belum ditentukan perusahaan \u2014 sistem tidak akan menebak nominal.")
            use_late_rate = st.checkbox("Atur tarif potongan telat", value=False)
            late_mode_choice = st.radio(
                "Metode potongan telat",
                ["Per Menit (Rp / menit)", "Per 1x Telat (Rp / kejadian)"],
                index=0,
                disabled=not use_late_rate,
            )
            is_per_occ = "1x" in late_mode_choice
            late_rate_per_min = 0.0
            late_rate_per_occ = 0.0
            if is_per_occ:
                late_rate_per_occ = st.number_input(
                    "Rp / 1x terlambat", min_value=0.0, value=0.0, disabled=not use_late_rate
                )
            else:
                late_rate_per_min = st.number_input(
                    "Rp / menit telat", min_value=0.0, value=0.0, disabled=not use_late_rate
                )

            use_sales_no_out_rate = st.checkbox("Atur potongan tidak absen pulang (khusus Sales)", value=False)
            sales_no_out_rate = st.number_input(
                "Rp / 1x tidak absen pulang (Sales)",
                min_value=0.0,
                value=0.0,
                disabled=not use_sales_no_out_rate,
                help="Potongan per kejadian tidak scan pulang khusus tim Sales (anomali sales pulang awal / mangkir).",
            )

            use_early_rate = st.checkbox("Atur tarif potongan pulang cepat")
            early_rate = st.number_input(
                "Rp / menit pulang cepat", min_value=0.0, value=0.0, disabled=not use_early_rate
            )
            use_absent_rate = st.checkbox("Atur potongan per hari mangkir")
            absent_rate = st.number_input("Rp / hari mangkir", min_value=0.0, value=0.0, disabled=not use_absent_rate)

        with st.expander("5. Bobot Skor Sales"):
            w_visit = st.slider("Bobot Visit", 0.0, 1.0, 0.15)
            w_td = st.slider("Bobot Test Drive", 0.0, 1.0, 0.15)
            w_spk = st.slider("Bobot SPK", 0.0, 1.0, 0.30)
            w_delivery = st.slider("Bobot Delivery", 0.0, 1.0, 0.25)
            w_conv = st.slider("Bobot Conversion Rate", 0.0, 1.0, 0.15)

        with st.expander("6. Identitas Slip Gaji & Upah Lembur"):
            company_name = st.text_input("Nama Perusahaan / Dealer", value="PT. Dealership Maju Bersama")
            company_address = st.text_input("Alamat Perusahaan", value="Surabaya, Indonesia")
            overtime_hourly_rate = st.number_input("Tarif Lembur per Jam (Rp / jam)", min_value=0.0, value=0.0, step=5000.0)

    payroll_args = {
        "late_deduction_mode": "per_occurrence" if is_per_occ else "per_minute",
        "late_deduction_per_minute": late_rate_per_min if (use_late_rate and not is_per_occ) else None,
        "late_deduction_per_occurrence": late_rate_per_occ if (use_late_rate and is_per_occ) else None,
        "sales_no_clock_out_deduction": sales_no_out_rate if use_sales_no_out_rate else None,
        "early_leave_deduction_per_minute": early_rate if use_early_rate else None,
        "absence_deduction_per_day": absent_rate if use_absent_rate else None,
    }
    sales_weights = {
        "weight_visit": w_visit,
        "weight_test_drive": w_td,
        "weight_spk": w_spk,
        "weight_delivery": w_delivery,
        "weight_conversion_rate": w_conv,
    }
    company_info = {
        "name": company_name,
        "address": company_address,
    }
    holidays_tuple = tuple(line.strip() for line in holidays_text.splitlines() if line.strip())

    return (
        attendance_file,
        master_file,
        leave_file,
        sales_file,
        clock_in.isoformat(),
        weekday_out.isoformat(),
        saturday_out.isoformat(),
        int(grace),
        saturday_working,
        sunday_off,
        start_date_val,
        end_date_val,
        holidays_tuple,
        (excellent_t, good_t, watch_t),
        payroll_args,
        sales_weights,
        company_info,
        overtime_hourly_rate,
    )


def render_overview_section(quality: dict, unmapped: pl.DataFrame, daily: pl.DataFrame, employee_summary: pl.DataFrame, insights: list[str]):
    _metric_grid(
        [
            ("Baris sumber", f"{quality['input_rows']:,}", "Jumlah baris dari file mesin."),
            ("Baris valid", f"{quality['valid_rows']:,}", "Baris yang lolos validasi dan cleaning."),
            ("Baris ditolak", f"{quality['rejected_rows']:,}", "Baris yang tidak dapat dipakai untuk kalkulasi."),
            ("Duplikat dihapus", f"{quality['duplicate_rows_removed']:,}", "Scan duplikat No. + Date/Time."),
            ("Karyawan", f"{quality['employees']:,}", "Jumlah Employee ID unik."),
        ]
    )
    if not unmapped.is_empty():
        st.warning(
            f"{len(unmapped)} Employee ID pada log absensi tidak ditemukan di Employee Master "
            "(memakai fallback nama/departemen dari mesin absensi)."
        )

    st.subheader("Executive Summary")
    working = int(daily["Is_Working_Day"].sum()) if not daily.is_empty() and "Is_Working_Day" in daily.columns else 0
    present = int(daily["Present_Flag"].sum()) if not daily.is_empty() and "Present_Flag" in daily.columns else 0
    attendance_rate = (present / working * 100.0) if working else 100.0
    avg_score = float(employee_summary["Attendance_Score"].mean()) if not employee_summary.is_empty() and "Attendance_Score" in employee_summary.columns else 0.0
    _metric_grid(
        [
            ("Attendance Rate", f"{attendance_rate:.1f}%", "Persentase kehadiran pada hari kerja."),
            ("Attendance Score", f"{avg_score:.1f}%", "Rata-rata skor kehadiran karyawan."),
            ("Terlambat", f"{int(daily['Late_Flag'].sum()):,}" if not daily.is_empty() else "0", "Total kejadian terlambat."),
            ("Total Menit Telat", f"{int(daily['Menit_Telat'].sum()):,}" if not daily.is_empty() else "0", "Akumulasi menit keterlambatan."),
            ("Mangkir", f"{int(daily['Absent_Flag'].sum()):,}" if not daily.is_empty() else "0", "Hari kerja tanpa log dan tanpa cuti/izin."),
            ("Anomali", f"{int(daily['Is_Anomali'].sum()):,}" if not daily.is_empty() else "0", "Hari dengan anomali yang perlu ditinjau HR."),
        ]
    )

    st.subheader("Insights HR")
    for insight in insights:
        st.info(insight)


def _render_payslip_card(slip: PayslipData, company_info: dict):
    """Render interactive document preview of the payslip exactly matching the generated template."""
    c_name = company_info.get("name", "PT. Dealership Maju Bersama")
    c_addr = company_info.get("address", "Surabaya, Indonesia")

    card_html = f"""
    <div style="background-color: #ffffff; color: #1a1a1a; padding: 28px; border-radius: 8px; border: 1px solid #d0d7de; font-family: 'Segoe UI', Inter, -apple-system, sans-serif; max-width: 820px; margin: 0 auto; box-shadow: 0 4px 16px rgba(0,0,0,0.07);">
        <!-- Header -->
        <div style="display: flex; justify-content: space-between; align-items: flex-start; border-bottom: 2.5px solid #1F4E78; padding-bottom: 12px; margin-bottom: 16px;">
            <div>
                <div style="font-size: 16px; font-weight: bold; color: #1F4E78; margin-bottom: 2px;">{c_name}</div>
                <div style="font-size: 12px; color: #595959;">{c_addr}</div>
            </div>
            <div style="font-size: 26px; font-weight: bold; color: #1a1a1a; letter-spacing: 0.5px;">Slip Gaji</div>
        </div>

        <!-- Meta -->
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 13px; margin-bottom: 16px; background-color: #f8fafc; padding: 12px 16px; border-radius: 6px; border: 1px solid #e2e8f0; line-height: 1.6;">
            <div><b>Nama / NIK:</b> {slip.name} ({slip.employee_id})</div>
            <div><b>Tgl Mulai Bekerja:</b> {slip.join_date or '-'}</div>
            <div><b>Dept / Jabatan:</b> {slip.department} / {slip.position}</div>
            <div><b>Periode Gaji:</b> {slip.period_label or '-'}</div>
        </div>

        <!-- Two column financial table -->
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px;">
            <!-- Earnings -->
            <div>
                <div style="background-color: #D9D9D9; color: #1a1a1a; padding: 7px 12px; font-weight: bold; font-size: 13px; border-radius: 4px 4px 0 0;">Pendapatan</div>
                <div style="border: 1px solid #e2e8f0; border-top: none; padding: 12px; font-size: 12.5px; line-height: 1.9; background-color: #ffffff;">
                    <div style="display: flex; justify-content: space-between;"><span>Gaji Pokok</span> <span>Rp {slip.gaji_pokok:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Lembur ({slip.lembur_hours:.1f} Jam)</span> <span>Rp {slip.lembur_pay:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Tunjangan Jabatan</span> <span>Rp {slip.tunjangan_jabatan:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Uang Makan</span> <span>Rp {slip.uang_makan:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Tunjangan Parkir / Transport</span> <span>Rp {slip.tunjangan_transport:,.0f}</span></div>
                    <div style="border-top: 2px solid #1a1a1a; margin-top: 10px; padding-top: 8px; display: flex; justify-content: space-between; font-weight: bold; font-size: 13px;">
                        <span>Total Pendapatan</span> <span>Rp {slip.total_pendapatan:,.0f}</span>
                    </div>
                </div>
            </div>

            <!-- Deductions -->
            <div>
                <div style="background-color: #D9D9D9; color: #1a1a1a; padding: 7px 12px; font-weight: bold; font-size: 13px; border-radius: 4px 4px 0 0;">Potongan</div>
                <div style="border: 1px solid #e2e8f0; border-top: none; padding: 12px; font-size: 12.5px; line-height: 1.9; background-color: #ffffff;">
                    <div style="display: flex; justify-content: space-between;"><span>Potongan Absen (Mangkir)</span> <span>Rp {slip.potongan_mangkir:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Potongan Datang Terlambat</span> <span>Rp {slip.potongan_telat:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Potongan Pulang Cepat</span> <span>Rp {slip.potongan_pulang_cepat:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Potongan Lupa Absen Pulang (Sales)</span> <span>Rp {slip.potongan_lupa_pulang_sales:,.0f}</span></div>
                    <div style="display: flex; justify-content: space-between;"><span>Potongan Lain-lain</span> <span>Rp {slip.potongan_lain:,.0f}</span></div>
                    <div style="border-top: 2px solid #1a1a1a; margin-top: 10px; padding-top: 8px; display: flex; justify-content: space-between; font-weight: bold; font-size: 13px;">
                        <span>Total Potongan</span> <span>Rp {slip.total_potongan:,.0f}</span>
                    </div>
                </div>
            </div>
        </div>

        <!-- Take Home Pay Box -->
        <div style="background-color: #EBF1F5; border: 1.5px solid #1F4E78; border-radius: 6px; padding: 12px 18px; display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
            <span style="font-size: 14px; font-weight: bold; color: #1F4E78;">Gaji Bersih / Take Home Pay</span>
            <span style="font-size: 20px; font-weight: bold; color: #1F4E78;">Rp {slip.take_home_pay:,.0f}</span>
        </div>

        <!-- Attendance Summary -->
        <div style="margin-bottom: 20px;">
            <div style="background-color: #D9D9D9; color: #1a1a1a; padding: 7px 12px; font-weight: bold; font-size: 13px; border-radius: 4px 4px 0 0;">Rangkuman Informasi Kehadiran</div>
            <div style="border: 1px solid #e2e8f0; border-top: none; padding: 12px 16px; font-size: 12.5px; display: grid; grid-template-columns: 1fr 1fr; gap: 16px; background-color: #ffffff; line-height: 1.8;">
                <div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Kehadiran:</span> <b>{slip.hari_kehadiran} Hari</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Ketidak Hadiran (Mangkir):</span> <b>{slip.hari_mangkir} Hari</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Cuti:</span> <b>{slip.hari_cuti} Hari</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Izin:</span> <b>{slip.hari_izin} Hari</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Sakit:</span> <b>{slip.hari_sakit} Hari</b></div>
                </div>
                <div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Terlambat:</span> <b>{slip.kali_terlambat} Kali ({slip.menit_terlambat} Menit)</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Pulang Lebih Dulu:</span> <b>{slip.kali_pulang_cepat} Kali</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Lupa Absen:</span> <b>{slip.kali_lupa_absen} Kali</b></div>
                    <div style="display: flex; justify-content: space-between; margin-bottom: 2px;"><span>Total Jam Lembur:</span> <b>{slip.total_jam_lembur:.1f} Jam</b></div>
                </div>
            </div>
        </div>
    </div>
    """
    st.iframe(card_html, height=800)


def render_tabs(res: dict, company_info: dict, has_sales_file: bool):
    trend = res["trend"]
    department_summary = res["department_summary"]
    employee_summary = res["employee_summary"]
    daily = res["daily"]
    anomalies = res["anomalies"]
    payroll = res["payroll"]
    sales_performance = res["sales_performance"]
    combined_view = res["combined_view"]
    payslips = res.get("payslips", [])
    cfg = res["cfg"]

    tabs = st.tabs(
        ["\U0001F4C8 Trends", "\U0001F3E2 Department", "\U0001F464 Employee", "\U0001F6A8 Anomaly",
         "\U0001F4B0 Payroll & Slip Gaji", "\U0001F697 Sales Performance"]
    )

    with tabs[0]:
        if not trend.is_empty():
            fig = px.line(trend, x="Tanggal", y="Attendance_Rate_%", markers=True, title="Trend Attendance Rate")
            fig.update_yaxes(range=[0, 100], ticksuffix="%")
            st.plotly_chart(fig, width="stretch")
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(px.bar(trend, x="Tanggal", y="Late", title="Keterlambatan Harian"), width="stretch")
            with c2:
                st.plotly_chart(px.bar(trend, x="Tanggal", y="Absent", title="Mangkir Harian"), width="stretch")

    with tabs[1]:
        if not department_summary.is_empty():
            fig = px.bar(
                department_summary.sort("Attendance_Score_%"), x="Department", y="Attendance_Score_%",
                text="Attendance_Score_%", title="Attendance Score per Department",
            )
            fig.update_yaxes(range=[0, 100], ticksuffix="%")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            st.plotly_chart(fig, width="stretch")
            st.dataframe(department_summary, width="stretch", hide_index=True)

    with tabs[2]:
        c1, c2 = st.columns([1, 2])
        with c1:
            depts = ["Semua"] + sorted([d for d in employee_summary["Department"].drop_nulls().unique().to_list()]) if not employee_summary.is_empty() else ["Semua"]
            selected_dept = st.selectbox("Filter departemen", depts)
        with c2:
            search = st.text_input("Cari nama / employee ID")

        view = employee_summary
        if selected_dept != "Semua" and not view.is_empty():
            view = view.filter(pl.col("Department") == selected_dept)
        if search and not view.is_empty():
            s_low = search.lower()
            view = view.filter(
                pl.col("Name").cast(pl.String).str.to_lowercase().str.contains(s_low)
                | pl.col("No.").cast(pl.String).str.to_lowercase().str.contains(s_low)
            )
        st.dataframe(view, width="stretch", hide_index=True)

        if not view.is_empty():
            no_list = view["No."].to_list()
            name_map = dict(zip(view["No."].to_list(), view["Name"].to_list()))
            selected_emp = st.selectbox(
                "Buka histori karyawan", no_list,
                format_func=lambda x: f"{x} \u2014 {name_map.get(x, '')}",
            )
            person = daily.filter(pl.col("No.") == selected_emp) if not daily.is_empty() else pl.DataFrame()
            hist_cols = [
                "Tanggal", "Hari", "Jam_Masuk", "Jam_Pulang", "Status_Masuk", "Menit_Telat",
                "Status_Pulang", "Menit_Pulang_Cepat", "Overtime_Hours", "Group_Event",
                "Scan_Count", "Anomaly_Type", "Severity",
            ]
            person_cols = [c for c in hist_cols if c in person.columns]
            st.dataframe(person.select(person_cols) if not person.is_empty() else pl.DataFrame(), width="stretch", hide_index=True)

    with tabs[3]:
        sev_order = ["HIGH", "LOW"]
        view = anomalies
        if not view.is_empty():
            sev_filter = st.multiselect("Severity", sev_order, default=sev_order)
            if sev_filter:
                view = view.filter(pl.col("Severity").is_in(sev_filter))
        anom_cols = [
            "No.", "Name", "Department", "Tanggal", "Hari", "Jam_Masuk", "Jam_Pulang",
            "Status_Masuk", "Status_Pulang", "Group_Event", "Anomaly_Type", "Severity",
        ]
        view_cols = [c for c in anom_cols if c in view.columns]
        st.dataframe(view.select(view_cols) if not view.is_empty() else pl.DataFrame(), width="stretch", hide_index=True)

    with tabs[4]:
        st.subheader("Rekap Payroll Karyawan")
        if not cfg.payroll.is_configured:
            st.warning("Tarif potongan belum diatur pada sidebar \u2014 kolom potongan akan kosong.")
        else:
            notes = []
            if cfg.payroll.late_deduction_mode == "per_occurrence" and cfg.payroll.late_deduction_per_occurrence:
                notes.append(f"Tarif telat: Rp {cfg.payroll.late_deduction_per_occurrence:,.0f} / 1x kejadian")
            elif cfg.payroll.late_deduction_per_minute:
                notes.append(f"Tarif telat: Rp {cfg.payroll.late_deduction_per_minute:,.0f} / menit")
            if cfg.payroll.sales_no_clock_out_deduction:
                notes.append(f"Potongan tidak absen pulang Sales: Rp {cfg.payroll.sales_no_clock_out_deduction:,.0f} / 1x")
            if notes:
                st.caption(" • ".join(notes))
        st.dataframe(payroll, width="stretch", hide_index=True)

        st.markdown("---")
        st.subheader("\U0001F4C4 Cetak & Ekspor Slip Gaji")

        if not payslips:
            st.info("Data slip gaji belum tersedia.")
        else:
            # 1. Bulk Export Section
            st.markdown("##### \U0001F4E6 Ekspor Massal (Bulk Export)")
            b1, b2 = st.columns(2)
            with b1:
                bulk_pdf = generate_bulk_payslip_pdf(payslips, company_info)
                st.download_button(
                    "\U0001F4E6 Download Semua Slip Gaji (Bulk PDF)",
                    data=bulk_pdf,
                    file_name="Bulk_Slip_Gaji_Karyawan.pdf",
                    mime="application/pdf",
                    width="stretch",
                    help="Mengunduh seluruh slip gaji karyawan dalam 1 file PDF multi-halaman.",
                )
            with b2:
                bulk_excel = generate_bulk_payslip_excel(payslips, company_info)
                st.download_button(
                    "\U0001F4CA Download Semua Slip Gaji (Bulk Excel)",
                    data=bulk_excel,
                    file_name="Bulk_Slip_Gaji_Karyawan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    help="Mengunduh seluruh slip gaji karyawan dalam 1 workbook Excel (1 sheet per karyawan).",
                )

            st.markdown("---")
            st.markdown("##### \U0001F464 Pratinjau & Cetak Slip Gaji Per Karyawan")

            # Selection
            slip_map = {s.employee_id: s for s in payslips}
            selected_slip_id = st.selectbox(
                "Pilih Karyawan",
                list(slip_map.keys()),
                format_func=lambda x: f"{x} \u2014 {slip_map[x].name} ({slip_map[x].department})",
                key="payslip_emp_selector",
            )

            current_slip = slip_map[selected_slip_id]

            # Action Buttons for single employee
            s1, s2 = st.columns(2)
            with s1:
                single_pdf = generate_payslip_pdf(current_slip, company_info)
                st.download_button(
                    f"\U0001F4C4 Download Slip {current_slip.name} (PDF)",
                    data=single_pdf,
                    file_name=f"Slip_Gaji_{current_slip.employee_id}_{current_slip.name.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    width="stretch",
                )
            with s2:
                single_excel = generate_payslip_excel(current_slip, company_info)
                st.download_button(
                    f"\U0001F4CA Download Slip {current_slip.name} (Excel)",
                    data=single_excel,
                    file_name=f"Slip_Gaji_{current_slip.employee_id}_{current_slip.name.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )

            st.write("")
            _render_payslip_card(current_slip, company_info)

    with tabs[5]:
        if not has_sales_file:
            st.info("Upload data Sales Activity pada sidebar untuk melihat performa sales.")
        elif sales_performance.is_empty():
            st.warning("Tidak ada baris valid pada file Sales Activity.")
        else:
            st.dataframe(sales_performance, width="stretch", hide_index=True)
            if not combined_view.is_empty():
                st.markdown("**Attendance vs Sales Performance**")
                st.dataframe(combined_view, width="stretch", hide_index=True)


def render_exports_section(res: dict):
    st.subheader("Export & Sharing")
    cfg = res["cfg"]
    daily = res["daily"]
    master = res["master"]
    clean = res["clean"]
    employee_summary = res["employee_summary"]
    department_summary = res["department_summary"]
    trend = res["trend"]
    anomalies = res["anomalies"]
    rejected = res["rejected"]
    payroll = res["payroll"]
    sales_activity = res["sales_activity"]
    sales_performance = res["sales_performance"]
    insights = res["insights"]
    unmapped = res["unmapped"]

    try:
        period_label = res.get("period_label", "")

        report = ReportData(
            config=cfg,
            employee_master=master,
            raw_log=clean,
            daily_attendance=daily,
            employee_summary=employee_summary,
            department_summary=department_summary,
            daily_trend=trend,
            anomalies=anomalies,
            rejected=rejected,
            payroll=payroll,
            sales_activity=sales_activity,
            sales_performance=sales_performance,
            insights=insights,
            report_period_label=period_label,
        )
        excel_bytes = generate_workbook(report)

        e1, e2 = st.columns(2)
        with e1:
            st.download_button(
                "\U0001F4E5 Download Excel HR Pack", data=excel_bytes,
                file_name="Laporan_HR_Attendance_Payroll_Sales.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
            )
        if REPORTLAB_AVAILABLE:
            pdf_bytes = _export_pdf(employee_summary, department_summary, insights)
            with e2:
                st.download_button(
                    "\U0001F4C4 Download PDF Management Report", data=pdf_bytes,
                    file_name="Laporan_HR_Attendance.pdf", mime="application/pdf", width="stretch",
                )
        else:
            with e2:
                st.warning("PDF export memerlukan package reportlab.")
    except Exception as exc:
        LOGGER.exception("Export failed")
        st.error(f"Export gagal: {exc}")

    if not rejected.is_empty():
        with st.expander(f"Data ditolak saat cleaning ({len(rejected):,} baris)"):
            st.dataframe(rejected, width="stretch", hide_index=True)

    if not unmapped.is_empty():
        with st.expander(f"Employee ID belum ada di Employee Master ({len(unmapped):,})"):
            st.dataframe(unmapped, width="stretch", hide_index=True)


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------


def main() -> None:
    st.set_page_config(
        page_title="HR Attendance & Sales Dashboard",
        page_icon="\U0001F4CA",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(
        """
        <style>
        .block-container { padding-top: 1.4rem; padding-bottom: 2rem; }
        div[data-testid="stMetric"] {
            background: var(--secondary-background-color);
            color: var(--text-color);
            border: 1px solid rgba(128, 128, 128, 0.28);
            padding: 14px;
            border-radius: 12px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("\U0001F4CA HR Attendance, Payroll & Sales Performance Dashboard")
    st.caption("Untuk HR, branch manager, dan supervisor dealer mobil")

    (
        attendance_file,
        master_file,
        leave_file,
        sales_file,
        clock_in_str,
        weekday_out_str,
        saturday_out_str,
        grace_minutes,
        saturday_working,
        sunday_off,
        start_date_val,
        end_date_val,
        holidays_tuple,
        score_thresholds,
        payroll_args,
        sales_weights,
        company_info,
        overtime_hourly_rate,
    ) = render_sidebar()

    if attendance_file is None:
        st.info("Upload export mesin absensi pada sidebar untuk memulai analisis.")
        return

    attendance_bytes = attendance_file.getvalue()
    master_bytes = master_file.getvalue() if master_file else None
    leave_bytes = leave_file.getvalue() if leave_file else None
    leave_is_csv = leave_file.name.lower().endswith(".csv") if leave_file else False
    sales_bytes = sales_file.getvalue() if sales_file else None
    sales_is_csv = sales_file.name.lower().endswith(".csv") if sales_file else False

    try:
        res = run_attendance_pipeline(
            attendance_bytes,
            master_bytes,
            leave_bytes,
            leave_is_csv,
            sales_bytes,
            sales_is_csv,
            clock_in_str,
            weekday_out_str,
            saturday_out_str,
            grace_minutes,
            saturday_working,
            sunday_off,
            start_date_val,
            end_date_val,
            holidays_tuple,
            score_thresholds,
            payroll_args,
            sales_weights,
            overtime_hourly_rate,
        )
    except (AttendanceValidationError, EmployeeMasterError, SalesActivityError) as exc:
        st.error(f"Data tidak dapat diproses: {exc}")
        return
    except Exception as exc:
        LOGGER.exception("Processing failed")
        st.error(f"Data tidak dapat diproses: {exc}")
        return

    render_overview_section(res["quality"], res["unmapped"], res["daily"], res["employee_summary"], res["insights"])
    render_tabs(res, company_info, has_sales_file=sales_file is not None)
    render_exports_section(res)


if __name__ == "__main__":
    main()
