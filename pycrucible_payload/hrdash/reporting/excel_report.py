"""
Excel workbook generator.

Produces the eleven-sheet workbook described in Section 10:
DASHBOARD, CONFIG, EMPLOYEE_MASTER, RAW_LOG, DAILY_ATTENDANCE,
EMPLOYEE_SUMMARY, ANOMALY, PAYROLL, SALES_PERFORMANCE, Sales_Activity,
README. The dashboard sheet is a decision-making surface (KPI cards +
priority tables + native Excel charts), not a dump of every column.
"""

from __future__ import annotations

import io
from dataclasses import dataclass

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from ..config import AppConfig, config_to_dataframe

NAVY = "1F4E78"
ACCENT_BLUE = "2E75B6"
LIGHT_BLUE = "D9EAF7"
LIGHT_GRAY = "F2F2F2"
WHITE = "FFFFFF"
GREEN = "2E7D32"
AMBER = "B7791F"
RED = "C0392B"

HEADER_FONT = Font(name="Aptos", bold=True, color=WHITE, size=10)
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(name="Aptos", bold=True, size=16, color=NAVY)
SUBTITLE_FONT = Font(name="Aptos", size=10, color="595959", italic=True)


@dataclass
class ReportData:
    config: AppConfig
    employee_master: pd.DataFrame
    raw_log: pd.DataFrame
    daily_attendance: pd.DataFrame
    employee_summary: pd.DataFrame
    department_summary: pd.DataFrame
    daily_trend: pd.DataFrame
    anomalies: pd.DataFrame
    rejected: pd.DataFrame
    payroll: pd.DataFrame
    sales_activity: pd.DataFrame
    sales_performance: pd.DataFrame
    insights: list[str]
    report_period_label: str = ""


# ---------------------------------------------------------------------------
# Generic sheet helpers
# ---------------------------------------------------------------------------


def _write_table(ws: Worksheet, df: pd.DataFrame, start_row: int = 1, start_col: int = 1) -> tuple[int, int]:
    """Write a styled table starting at (start_row, start_col). Returns the
    (last_row, last_col) written."""
    if df is None or df.empty:
        cell = ws.cell(row=start_row, column=start_col, value="(Tidak ada data)")
        cell.font = Font(italic=True, color="7F7F7F")
        return start_row, start_col

    for j, col_name in enumerate(df.columns):
        cell = ws.cell(row=start_row, column=start_col + j, value=str(col_name))
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        for j, value in enumerate(row):
            v = value
            if pd.isna(v):
                v = None
            elif hasattr(v, "isoformat"):
                v = v
            cell = ws.cell(row=start_row + i, column=start_col + j, value=v)
            if i % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=LIGHT_GRAY)

    last_row = start_row + len(df)
    last_col = start_col + len(df.columns) - 1
    ws.auto_filter.ref = (
        f"{get_column_letter(start_col)}{start_row}:{get_column_letter(last_col)}{last_row}"
    )
    ws.freeze_panes = ws.cell(row=start_row + 1, column=start_col).coordinate
    return last_row, last_col


def _autofit(ws: Worksheet, max_scan_row: int = 2000) -> None:
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        longest = 0
        for row_idx in range(1, min(ws.max_row, max_scan_row) + 1):
            value = ws.cell(row_idx, col_idx).value
            if value is not None:
                longest = max(longest, len(str(value)))
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), 42)


def _data_sheet(wb: Workbook, name: str, df: pd.DataFrame) -> Worksheet:
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    _write_table(ws, df)
    _autofit(ws)
    return ws


# ---------------------------------------------------------------------------
# KPI card helper (used on the dashboard)
# ---------------------------------------------------------------------------


def _kpi_card(ws: Worksheet, row: int, col: int, label: str, value: str, accent: str = ACCENT_BLUE) -> None:
    label_cell = ws.cell(row=row, column=col, value=label)
    label_cell.font = Font(name="Aptos", size=9, color="595959")
    label_cell.alignment = Alignment(horizontal="left")

    value_cell = ws.cell(row=row + 1, column=col, value=value)
    value_cell.font = Font(name="Aptos", size=20, bold=True, color=accent)
    value_cell.alignment = Alignment(horizontal="left")

    for r in (row, row + 1):
        c = ws.cell(row=r, column=col)
        c.fill = PatternFill("solid", fgColor=LIGHT_GRAY)
        c.border = Border(
            left=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9") if r == row else None,
            bottom=Side(style="thin", color="D9D9D9") if r == row + 1 else None,
        )
    ws.row_dimensions[row].height = 16
    ws.row_dimensions[row + 1].height = 26


def _section_title(ws: Worksheet, row: int, col: int, text: str) -> None:
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(name="Aptos", bold=True, size=12, color=NAVY)


# ---------------------------------------------------------------------------
# DASHBOARD sheet
# ---------------------------------------------------------------------------


def _build_dashboard(wb: Workbook, data: ReportData) -> None:
    ws = wb.create_sheet("DASHBOARD", 0)
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    ws["A1"] = "HR Attendance, Payroll & Sales Performance Dashboard"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A2:H2")
    period = data.report_period_label or "Seluruh periode data"
    ws["A2"] = f"Periode: {period}"
    ws["A2"].font = SUBTITLE_FONT

    daily = data.daily_attendance
    emp_summary = data.employee_summary
    total_employees = int(daily["No."].nunique()) if not daily.empty else 0
    working = int(daily["Is_Working_Day"].sum()) if not daily.empty else 0
    present = int(daily["Present_Flag"].sum()) if not daily.empty else 0
    attendance_rate = (present / working * 100) if working else 0.0
    total_late = int(daily["Late_Flag"].sum()) if not daily.empty else 0
    total_absent = int(daily["Absent_Flag"].sum()) if not daily.empty else 0

    kpis = [
        ("Total Karyawan", f"{total_employees:,}", ACCENT_BLUE),
        ("Attendance Rate", f"{attendance_rate:.1f}%", GREEN if attendance_rate >= 90 else AMBER),
        ("Total Terlambat", f"{total_late:,}", AMBER if total_late else GREEN),
        ("Total Mangkir", f"{total_absent:,}", RED if total_absent else GREEN),
    ]
    for i, (label, value, accent) in enumerate(kpis):
        _kpi_card(ws, 4, 1 + i * 2, label, value, accent)

    row = 8
    _section_title(ws, row, 1, "Perlu Perhatian HR (Attendance Score Terendah)")
    row += 1
    if not emp_summary.empty:
        watch = emp_summary.nsmallest(10, "Attendance_Score")[
            ["No.", "Name", "Department", "Terlambat", "Mangkir", "Attendance_Score", "HR_Classification"]
        ]
        last_row, last_col = _write_table(ws, watch, start_row=row, start_col=1)
    else:
        ws.cell(row=row, column=1, value="(Tidak ada data)")
        last_row = row

    sales_row = 8
    sales_col = 7
    if not data.sales_performance.empty:
        _section_title(ws, sales_row, sales_col, "Top Sales Performance")
        sales_row += 1
        top_sales = data.sales_performance.head(10)[
            ["No.", "Name", "SPK", "Delivery", "Revenue", "Performance_Score", "Performance_Classification"]
        ]
        _write_table(ws, top_sales, start_row=sales_row, start_col=sales_col)

    insight_row = last_row + 3
    _section_title(ws, insight_row, 1, "Insight HR")
    insight_row += 1
    for text in data.insights:
        ws.cell(row=insight_row, column=1, value=f"• {text}")
        ws.merge_cells(start_row=insight_row, start_column=1, end_row=insight_row, end_column=8)
        ws.cell(row=insight_row, column=1).alignment = Alignment(wrap_text=True)
        insight_row += 1

    # --- Trend chart data block (kept off to the side, used only as the
    # chart's data source) ---
    chart_row = insight_row + 2
    _section_title(ws, chart_row, 1, "Trend Kehadiran Harian")
    data_start = chart_row + 1
    trend = data.daily_trend
    if not trend.empty:
        _write_table(ws, trend[["Tanggal", "Attendance_Rate_%", "Late", "Absent"]], start_row=data_start, start_col=1)
        n = len(trend)

        line = LineChart()
        line.title = "Attendance Rate Harian (%)"
        line.y_axis.title = "%"
        line.y_axis.scaling.min = 0
        line.y_axis.scaling.max = 100
        line.height = 8
        line.width = 22
        cats = Reference(ws, min_col=1, min_row=data_start + 1, max_row=data_start + n)
        vals = Reference(ws, min_col=2, min_row=data_start, max_row=data_start + n)
        line.add_data(vals, titles_from_data=True)
        line.set_categories(cats)
        ws.add_chart(line, f"F{data_start}")

        bar = BarChart()
        bar.type = "col"
        bar.title = "Terlambat vs Mangkir Harian"
        bar.height = 8
        bar.width = 22
        vals2 = Reference(ws, min_col=3, max_col=4, min_row=data_start, max_row=data_start + n)
        bar.add_data(vals2, titles_from_data=True)
        bar.set_categories(cats)
        ws.add_chart(bar, f"F{data_start + 17}")

    if not data.department_summary.empty:
        dept_chart_row = data_start + len(trend) + 35
        _section_title(ws, dept_chart_row, 1, "Attendance Score per Departemen")
        dept_data_row = dept_chart_row + 1
        _write_table(
            ws,
            data.department_summary[["Department", "Attendance_Score_%"]],
            start_row=dept_data_row,
            start_col=1,
        )
        n = len(data.department_summary)
        dept_bar = BarChart()
        dept_bar.type = "bar"
        dept_bar.title = "Attendance Score per Departemen"
        dept_bar.height = 8
        dept_bar.width = 22
        cats = Reference(ws, min_col=1, min_row=dept_data_row + 1, max_row=dept_data_row + n)
        vals = Reference(ws, min_col=2, min_row=dept_data_row, max_row=dept_data_row + n)
        dept_bar.add_data(vals, titles_from_data=True)
        dept_bar.set_categories(cats)
        ws.add_chart(dept_bar, f"F{dept_data_row}")

    for col, width in {"A": 14, "B": 20, "C": 16, "D": 12, "E": 12, "F": 16, "G": 14, "H": 16}.items():
        ws.column_dimensions[col].width = width


# ---------------------------------------------------------------------------
# README sheet
# ---------------------------------------------------------------------------

README_ROWS = [
    ("PANDUAN WORKBOOK", ""),
    ("DASHBOARD", "Ringkasan eksekutif: KPI utama, karyawan yang perlu perhatian HR, top sales, dan trend."),
    ("CONFIG", "Semua aturan bisnis (jadwal, toleransi, opsi potongan telat, potongan sales, bobot sales)."),
    ("EMPLOYEE_MASTER", "Data master karyawan: nama, departemen, tipe, jam masuk & pulang divisi, gaji bulanan."),
    ("RAW_LOG", "Transaksi mentah dari mesin absensi setelah validasi dasar. Tidak diubah/dihapus (audit trail)."),
    ("DAILY_ATTENDANCE", "Hasil perhitungan harian per karyawan: scan pertama/terakhir, status, menit telat, lembur, anomali."),
    ("EMPLOYEE_SUMMARY", "KPI kehadiran per karyawan: attendance rate, punctuality, attendance score, tidak absen pulang, lembur."),
    ("ANOMALY", "Hanya baris dengan anomali (severity HIGH/LOW, lupa scan pulang, scan hari libur) untuk ditelusuri HR."),
    ("PAYROLL", "Estimasi potongan gaji: telat (per menit/kejadian), mangkir, dan potongan 1x tidak absen pulang khusus tim Sales."),
    ("SALES_PERFORMANCE", "KPI dan skor performa sales, dihitung HANYA dari Sales_Activity — bukan dari fingerprint."),
    ("Sales_Activity", "Input aktivitas sales (prospect, visit, test drive, SPK, delivery, revenue) per tanggal."),
    ("", ""),
    ("CATATAN PENTING", ""),
    (
        "Aturan First/Last Scan",
        "Scan pertama = absen masuk, scan terakhir = absen pulang. Scan di antaranya TIDAK ditafsirkan "
        "sebagai istirahat/keluar-masuk kantor, dicatat sebagai potensi anomali untuk audit.",
    ),
    (
        "Sales & Canvassing",
        "Mesin absensi tidak mencatat aktivitas canvassing sales. Jumlah scan TIDAK digunakan sebagai KPI sales. "
        "Performa sales murni berasal dari sheet Sales_Activity.",
    ),
    (
        "Data Ditolak",
        "Baris raw yang tidak valid (ID kosong / tanggal tidak terbaca) tidak dihitung tetapi tetap disimpan "
        "terpisah untuk audit (lihat sheet Data_Ditolak jika ada).",
    ),
]


def _build_readme(wb: Workbook) -> None:
    ws = wb.create_sheet("README")
    ws.sheet_view.showGridLines = False
    for i, (label, desc) in enumerate(README_ROWS, start=1):
        c1 = ws.cell(row=i, column=1, value=label)
        c2 = ws.cell(row=i, column=2, value=desc)
        if desc == "" and label:
            c1.font = Font(bold=True, size=13, color=WHITE)
            c1.fill = PatternFill("solid", fgColor=NAVY)
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=2)
        else:
            c1.font = Font(bold=True, color=NAVY)
            c2.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110


# ---------------------------------------------------------------------------
# Public entrypoint
# ---------------------------------------------------------------------------


def generate_workbook(data: ReportData) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    _build_dashboard(wb, data)
    _data_sheet(wb, "CONFIG", config_to_dataframe(data.config))
    _data_sheet(wb, "EMPLOYEE_MASTER", data.employee_master)
    _data_sheet(wb, "RAW_LOG", data.raw_log)
    _data_sheet(wb, "DAILY_ATTENDANCE", data.daily_attendance)
    _data_sheet(wb, "EMPLOYEE_SUMMARY", data.employee_summary)
    _data_sheet(wb, "ANOMALY", data.anomalies)
    _data_sheet(wb, "PAYROLL", data.payroll)
    _data_sheet(wb, "SALES_PERFORMANCE", data.sales_performance)
    _data_sheet(wb, "Sales_Activity", data.sales_activity)
    if data.rejected is not None and not data.rejected.empty:
        _data_sheet(wb, "Data_Ditolak", data.rejected)
    _build_readme(wb)

    for ws in wb.worksheets:
        if ws.title != "DASHBOARD":
            ws.freeze_panes = ws.freeze_panes or "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
