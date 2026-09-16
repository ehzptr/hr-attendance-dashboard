"""
Streamlit UI entrypoint.

Thin presentation layer only — all business logic lives in the other
``hrdash`` modules. Run with:

    streamlit run -m hrdash.app
"""

from __future__ import annotations

import io
import logging
from datetime import date, time

import pandas as pd
import plotly.express as px
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
from .payroll import calculate_payroll
from .reporting.excel_report import ReportData, generate_workbook
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


def _build_sample_attendance() -> bytes:
    sample = pd.DataFrame(
        {
            "Department": ["Sales Mobil Baru", "Sales Mobil Baru", "HRD & Admin", "HRD & Admin"],
            "No.": ["EMP001", "EMP001", "EMP002", "EMP002"],
            "Name": ["Budi", "Budi", "Sari", "Sari"],
            "Date/Time": [
                "15/09/2026 07.54.23",
                "15/09/2026 17.06.11",
                "15/09/2026 08.22.02",
                "15/09/2026 16.30.44",
            ],
        }
    )
    b = io.BytesIO()
    sample.to_excel(b, index=False)
    return b.getvalue()


def _export_pdf(employee_summary: pd.DataFrame, department_summary: pd.DataFrame, insights: list[str]) -> bytes:
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
    top = employee_summary.nsmallest(15, "Attendance_Score")
    data = [["Karyawan", "Departemen", "Hadir", "Terlambat", "Mangkir", "Attendance Score"]]
    for _, r in top.iterrows():
        data.append(
            [str(r["Name"]), str(r["Department"]), str(int(r["Hadir"])), str(int(r["Terlambat"])),
             str(int(r["Mangkir"])), f"{r['Attendance_Score']:.1f}%"]
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

    if not department_summary.empty:
        story.append(Spacer(1, 5 * mm))
        story.append(Paragraph("Ringkasan Departemen", styles["Heading2"]))
        ddata = [["Departemen", "Employee", "Attendance", "Score", "Anomali"]]
        for _, r in department_summary.iterrows():
            ddata.append(
                [str(r["Department"]), str(int(r["Employee"])), f"{r['Attendance_Rate_%']:.1f}%",
                 f"{r['Attendance_Score_%']:.1f}%", str(int(r["Anomalies"]))]
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

    if attendance_file is None:
        st.info("Upload export mesin absensi pada sidebar untuk memulai analisis.")
        return

    cfg = AppConfig(
        attendance=AttendanceConfig(
            clock_in=clock_in,
            weekday_clock_out=weekday_out,
            saturday_clock_out=saturday_out,
            grace_minutes=int(grace),
            saturday_is_working=saturday_working,
            sunday_is_off=sunday_off,
            start_date=start_date_val,
            end_date=end_date_val,
        ),
        score=ScoreConfig(excellent_threshold=excellent_t, good_threshold=good_t, watch_threshold=watch_t),
        payroll=PayrollConfig(
            late_deduction_mode="per_occurrence" if is_per_occ else "per_minute",
            late_deduction_per_minute=late_rate_per_min if (use_late_rate and not is_per_occ) else None,
            late_deduction_per_occurrence=late_rate_per_occ if (use_late_rate and is_per_occ) else None,
            sales_no_clock_out_deduction=sales_no_out_rate if use_sales_no_out_rate else None,
            early_leave_deduction_per_minute=early_rate if use_early_rate else None,
            absence_deduction_per_day=absent_rate if use_absent_rate else None,
        ),
        sales_score=SalesScoreConfig(
            weight_visit=w_visit, weight_test_drive=w_td, weight_spk=w_spk,
            weight_delivery=w_delivery, weight_conversion_rate=w_conv,
        ),
    )

    try:
        raw = pd.read_excel(attendance_file)
        clean, rejected, quality = clean_attendance_data(raw)
        if clean.empty:
            raise AttendanceValidationError("Tidak ada baris valid setelah proses cleaning.")

        master = pd.DataFrame()
        if master_file is not None:
            master = load_employee_master(pd.read_excel(master_file))

        enriched, unmapped = merge_master_into_log(clean, master)

        leave_result = None
        approved_leave = None
        if leave_file is not None:
            if leave_file.name.lower().endswith(".csv"):
                leave_raw = pd.read_csv(leave_file)
                approved_leave = clean_approved_leave(leave_raw)
            else:
                leave_result = load_leave_workbook(leave_file)

        holidays = set()
        for line in holidays_text.splitlines():
            line = line.strip()
            if line:
                holidays.add(pd.Timestamp(line).date())

        daily = att_engine.build_daily_attendance(
            enriched, cfg.attendance, holidays, approved_leave=approved_leave, leave_result=leave_result
        )
        employee_summary = att_analytics.calculate_employee_summary(daily, cfg.score)
        department_summary = att_analytics.calculate_department_summary(daily)
        trend = att_analytics.calculate_daily_trend(daily)
        anomalies = daily[daily["Is_Anomali"] == 1].copy()
        insights = att_analytics.derive_hr_insights(employee_summary, department_summary)
        payroll = calculate_payroll(employee_summary, master, cfg.payroll)

        sales_activity = pd.DataFrame()
        sales_performance = pd.DataFrame()
        combined_view = pd.DataFrame()
        if sales_file is not None:
            sales_raw = (
                pd.read_csv(sales_file) if sales_file.name.lower().endswith(".csv") else pd.read_excel(sales_file)
            )
            sales_activity = clean_sales_activity(sales_raw)
            sales_performance = calculate_sales_performance(sales_activity, master, cfg.sales_score)
            combined_view = combine_attendance_and_sales(employee_summary, sales_performance)

    except (AttendanceValidationError, EmployeeMasterError, SalesActivityError) as exc:
        st.error(f"Data tidak dapat diproses: {exc}")
        return
    except Exception as exc:
        LOGGER.exception("Processing failed")
        st.error(f"Data tidak dapat diproses: {exc}")
        return

    # ------------------------------------------------------------------
    # Data quality
    # ------------------------------------------------------------------
    _metric_grid(
        [
            ("Baris sumber", f"{quality['input_rows']:,}", "Jumlah baris dari file mesin."),
            ("Baris valid", f"{quality['valid_rows']:,}", "Baris yang lolos validasi dan cleaning."),
            ("Baris ditolak", f"{quality['rejected_rows']:,}", "Baris yang tidak dapat dipakai untuk kalkulasi."),
            ("Duplikat dihapus", f"{quality['duplicate_rows_removed']:,}", "Scan duplikat No. + Date/Time."),
            ("Karyawan", f"{quality['employees']:,}", "Jumlah Employee ID unik."),
        ]
    )
    if len(unmapped):
        st.warning(
            f"{len(unmapped)} Employee ID pada log absensi tidak ditemukan di Employee Master "
            "(memakai fallback nama/departemen dari mesin absensi)."
        )

    st.subheader("Executive Summary")
    working = int(daily["Is_Working_Day"].sum())
    present = int(daily["Present_Flag"].sum())
    attendance_rate = (present / working * 100) if working else 100
    avg_score = float(employee_summary["Attendance_Score"].mean()) if not employee_summary.empty else 0
    _metric_grid(
        [
            ("Attendance Rate", f"{attendance_rate:.1f}%", "Persentase kehadiran pada hari kerja."),
            ("Attendance Score", f"{avg_score:.1f}%", "Rata-rata skor kehadiran karyawan."),
            ("Terlambat", f"{int(daily['Late_Flag'].sum()):,}", "Total kejadian terlambat."),
            ("Total Menit Telat", f"{int(daily['Menit_Telat'].sum()):,}", "Akumulasi menit keterlambatan."),
            ("Mangkir", f"{int(daily['Absent_Flag'].sum()):,}", "Hari kerja tanpa log dan tanpa cuti/izin."),
            ("Anomali", f"{int(daily['Is_Anomali'].sum()):,}", "Hari dengan anomali yang perlu ditinjau HR."),
        ]
    )

    st.subheader("Insights HR")
    for insight in insights:
        st.info(insight)

    tabs = st.tabs(
        ["\U0001F4C8 Trends", "\U0001F3E2 Department", "\U0001F464 Employee", "\U0001F6A8 Anomaly",
         "\U0001F4B0 Payroll", "\U0001F697 Sales Performance"]
    )

    with tabs[0]:
        if not trend.empty:
            fig = px.line(trend, x="Tanggal", y="Attendance_Rate_%", markers=True, title="Trend Attendance Rate")
            fig.update_yaxes(range=[0, 100], ticksuffix="%")
            st.plotly_chart(fig, width="stretch")
            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(px.bar(trend, x="Tanggal", y="Late", title="Keterlambatan Harian"), width="stretch")
            with c2:
                st.plotly_chart(px.bar(trend, x="Tanggal", y="Absent", title="Mangkir Harian"), width="stretch")

    with tabs[1]:
        if not department_summary.empty:
            fig = px.bar(
                department_summary.sort_values("Attendance_Score_%"), x="Department", y="Attendance_Score_%",
                text="Attendance_Score_%", title="Attendance Score per Department",
            )
            fig.update_yaxes(range=[0, 100], ticksuffix="%")
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            st.plotly_chart(fig, width="stretch")
            st.dataframe(
                department_summary.style.format({"Attendance_Rate_%": "{:.1f}%", "Attendance_Score_%": "{:.1f}%"}),
                width="stretch", hide_index=True,
            )

    with tabs[2]:
        c1, c2 = st.columns([1, 2])
        with c1:
            depts = ["Semua"] + sorted(employee_summary["Department"].dropna().unique().tolist())
            selected_dept = st.selectbox("Filter departemen", depts)
        with c2:
            search = st.text_input("Cari nama / employee ID")

        view = employee_summary.copy()
        if selected_dept != "Semua":
            view = view[view["Department"] == selected_dept]
        if search:
            mask = view["Name"].str.contains(search, case=False, na=False) | view["No."].astype(str).str.contains(
                search, case=False, na=False
            )
            view = view[mask]
        st.dataframe(
            view.style.format({"Attendance_Rate_%": "{:.1f}%", "Punctuality_Rate_%": "{:.1f}%", "Attendance_Score": "{:.1f}%"}),
            width="stretch", hide_index=True,
        )

        if not view.empty:
            selected_emp = st.selectbox(
                "Buka histori karyawan", view["No."].tolist(),
                format_func=lambda x: f"{x} \u2014 {view.loc[view['No.'].eq(x), 'Name'].iloc[0]}",
            )
            person = daily[daily["No."] == selected_emp]
            hist_cols = [
                "Tanggal", "Hari", "Jam_Masuk", "Jam_Pulang", "Status_Masuk", "Menit_Telat",
                "Status_Pulang", "Menit_Pulang_Cepat", "Overtime_Hours", "Group_Event",
                "Scan_Count", "Anomaly_Type", "Severity",
            ]
            person_cols = [c for c in hist_cols if c in person.columns]
            st.dataframe(person[person_cols], width="stretch", hide_index=True)

    with tabs[3]:
        sev_order = ["HIGH", "LOW"]
        view = anomalies.copy()
        if not view.empty:
            sev_filter = st.multiselect("Severity", sev_order, default=sev_order)
            if sev_filter:
                view = view[view["Severity"].isin(sev_filter)]
        anom_cols = [
            "No.", "Name", "Department", "Tanggal", "Hari", "Jam_Masuk", "Jam_Pulang",
            "Status_Masuk", "Status_Pulang", "Group_Event", "Anomaly_Type", "Severity",
        ]
        view_cols = [c for c in anom_cols if c in view.columns]
        st.dataframe(view[view_cols], width="stretch", hide_index=True)

    with tabs[4]:
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

    with tabs[5]:
        if sales_file is None:
            st.info("Upload data Sales Activity pada sidebar untuk melihat performa sales.")
        elif sales_performance.empty:
            st.warning("Tidak ada baris valid pada file Sales Activity.")
        else:
            st.dataframe(
                sales_performance.style.format(
                    {"Conversion_Rate_%": "{:.1f}%", "Performance_Score": "{:.1f}"}
                ),
                width="stretch", hide_index=True,
            )
            if not combined_view.empty:
                st.markdown("**Attendance vs Sales Performance**")
                st.dataframe(
                    combined_view.style.format(
                        {"Attendance_Rate_%": "{:.1f}%", "Attendance_Score": "{:.1f}%", "Performance_Score": "{:.1f}"}
                    ),
                    width="stretch", hide_index=True,
                )

    # ------------------------------------------------------------------
    # Exports
    # ------------------------------------------------------------------
    st.subheader("Export & Sharing")
    try:
        period_label = ""
        if not daily.empty:
            period_label = f"{daily['Tanggal'].min()} s/d {daily['Tanggal'].max()}"

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

    if not rejected.empty:
        with st.expander(f"Data ditolak saat cleaning ({len(rejected):,} baris)"):
            st.dataframe(rejected, width="stretch", hide_index=True)

    if len(unmapped):
        with st.expander(f"Employee ID belum ada di Employee Master ({len(unmapped):,})"):
            st.dataframe(unmapped, width="stretch", hide_index=True)


if __name__ == "__main__":
    main()
