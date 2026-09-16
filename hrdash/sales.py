"""
Sales performance module (Section 4).

Sales performance is deliberately kept as a *separate* data source and
pipeline from attendance. Nothing here reads scan counts, first/last scan
times, or any other fingerprint-derived value — canvassing is not observable
on the attendance machine (Section 3) and this module never pretends it is.
"""

from __future__ import annotations

import io

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .config import SalesScoreConfig

SALES_ACTIVITY_COLUMNS = ["No.", "Tanggal", "Prospect", "Visit", "Test_Drive", "SPK", "Delivery", "Revenue"]

SALES_PERFORMANCE_COLUMNS = [
    "No.", "Name", "Department",
    "Prospect", "Visit", "Test_Drive", "SPK", "Delivery", "Revenue",
    "Conversion_Rate_%", "Performance_Score", "Performance_Classification",
]

NAVY = "1F4E78"
WHITE = "FFFFFF"


class SalesActivityError(ValueError):
    pass


def clean_sales_activity(raw_df: pd.DataFrame | None) -> pd.DataFrame:
    """Validate and normalize an uploaded Sales Activity file.

    Required columns: employee ID and date. All KPI columns are optional
    and default to 0 when absent, since a sales rep may have zero of a
    given activity on a given day — that is data, not a missing value.
    """
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=SALES_ACTIVITY_COLUMNS)

    df = raw_df.copy()
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    id_col = next((c for c in ["No.", "Employee ID", "ID", "NIK"] if c in df.columns), None)
    date_col = next((c for c in ["Tanggal", "Date"] if c in df.columns), None)
    if not id_col or not date_col:
        raise SalesActivityError(
            "File Sales Activity harus memiliki kolom employee ID (No./Employee ID) "
            "dan tanggal (Tanggal/Date)."
        )

    out = pd.DataFrame()
    out["No."] = df[id_col].astype("string").str.strip()
    out["Tanggal"] = pd.to_datetime(df[date_col], format="mixed", dayfirst=True, errors="coerce").dt.date

    kpi_map = {
        "Prospect": ["Prospect", "Prospek"],
        "Visit": ["Visit", "Customer Visit"],
        "Test_Drive": ["Test Drive", "Test_Drive"],
        "SPK": ["SPK"],
        "Delivery": ["Delivery"],
        "Revenue": ["Revenue"],
    }
    for target, candidates in kpi_map.items():
        col = next((c for c in candidates if c in df.columns), None)
        out[target] = pd.to_numeric(df[col], errors="coerce").fillna(0) if col else 0.0

    out = out.dropna(subset=["No.", "Tanggal"])
    return out[SALES_ACTIVITY_COLUMNS].reset_index(drop=True)


def _normalize_0_100(series: pd.Series) -> pd.Series:
    lo, hi = series.min(), series.max()
    if pd.isna(lo) or pd.isna(hi) or hi == lo:
        return pd.Series(np.where(series > 0, 100.0, 0.0), index=series.index)
    return (series - lo) / (hi - lo) * 100.0


def calculate_sales_performance(
    sales_activity: pd.DataFrame,
    master: pd.DataFrame,
    cfg: SalesScoreConfig,
) -> pd.DataFrame:
    """Aggregate Sales Activity into per-employee sales KPIs (Section 4).

    Performance_Score blends KPI volume (Visit/Test Drive/SPK/Delivery) and
    Conversion Rate using HR-configurable weights (``SalesScoreConfig``),
    each metric relatively normalised (0-100) across the current sales team
    for the selected period, since no fixed company-wide sales target was
    supplied.
    """
    if sales_activity.empty:
        return pd.DataFrame(columns=SALES_PERFORMANCE_COLUMNS)

    agg = sales_activity.groupby("No.", as_index=False).agg(
        Prospect=("Prospect", "sum"),
        Visit=("Visit", "sum"),
        Test_Drive=("Test_Drive", "sum"),
        SPK=("SPK", "sum"),
        Delivery=("Delivery", "sum"),
        Revenue=("Revenue", "sum"),
    )
    agg["Conversion_Rate_%"] = np.where(
        agg["Prospect"] > 0, agg["SPK"] / agg["Prospect"] * 100, 0.0
    )

    if master is not None and not master.empty:
        m = master.set_index("No.")
        agg["Name"] = agg["No."].map(m["Name"]).fillna(agg["No."].map(lambda x: f"Sales {x}"))
        agg["Department"] = agg["No."].map(m["Department"]).fillna("Belum Dipetakan")
    else:
        agg["Name"] = agg["No."].map(lambda x: f"Sales {x}")
        agg["Department"] = "Belum Dipetakan"

    weights = {
        "Visit": cfg.weight_visit,
        "Test_Drive": cfg.weight_test_drive,
        "SPK": cfg.weight_spk,
        "Delivery": cfg.weight_delivery,
        "Conversion_Rate_%": cfg.weight_conversion_rate,
    }
    total_weight = sum(weights.values()) or 1.0

    score = pd.Series(0.0, index=agg.index)
    for col, w in weights.items():
        score = score + _normalize_0_100(agg[col]) * (w / total_weight)
    agg["Performance_Score"] = score.round(1)
    agg["Performance_Classification"] = agg["Performance_Score"].map(cfg.classify)

    return agg[SALES_PERFORMANCE_COLUMNS].sort_values(
        "Performance_Score", ascending=False, kind="stable"
    ).reset_index(drop=True)


def combine_attendance_and_sales(
    employee_summary: pd.DataFrame, sales_performance: pd.DataFrame
) -> pd.DataFrame:
    """Management view (Section 4 example table): attendance discipline
    next to sales outcomes, joined on Employee ID only — the two data
    sources are never blended at the raw-data level.
    """
    if employee_summary.empty or sales_performance.empty:
        return pd.DataFrame()

    att = employee_summary[
        ["No.", "Name", "Department", "Attendance_Rate_%", "Terlambat", "Attendance_Score", "HR_Classification"]
    ]
    combined = sales_performance.merge(att, on="No.", suffixes=("", "_att"), how="left")
    combined["Name"] = combined["Name"].fillna(combined.get("Name_att"))
    if "Name_att" in combined.columns:
        combined = combined.drop(columns=["Name_att"])
    if "Department_att" in combined.columns:
        combined = combined.drop(columns=["Department_att"])
    return combined


def build_sales_activity_template() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Activity"
    guide = wb.create_sheet("Petunjuk")

    ws.append(["No.", "Tanggal", "Prospect", "Visit", "Test Drive", "SPK", "Delivery", "Revenue"])
    ws.append(["EMP001", "2026-09-01", 8, 4, 2, 1, 0, 0])
    ws.append(["EMP001", "2026-09-02", 6, 3, 1, 1, 1, 250000000])

    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
    body_font = Font(name="Segoe UI", size=10, color="1F2937")
    thin = Side(style="thin", color="D1D5DB")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)
    for row in ws.iter_rows(min_row=2, max_row=1000, max_col=8):
        for cell in row:
            cell.font = body_font
            cell.border = Border(bottom=thin)
    ws["B2"].number_format = "yyyy-mm-dd"
    ws["B3"].number_format = "yyyy-mm-dd"

    widths = {"A": 12, "B": 14, "C": 10, "D": 10, "E": 12, "F": 8, "G": 10, "H": 14}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:H1000"

    guide_rows = [
        ["PETUNJUK SALES ACTIVITY"],
        ["Tujuan", "Mencatat aktivitas sales harian, TERPISAH dari data absensi fingerprint."],
        ["No.", "Employee ID sales, harus sama dengan No. pada Employee Master / mesin absensi."],
        ["Tanggal", "Tanggal aktivitas."],
        ["Prospect / Visit / Test Drive / SPK / Delivery", "Jumlah kejadian pada tanggal tersebut."],
        ["Revenue", "Nilai revenue dari delivery pada tanggal tersebut (opsional)."],
        ["Catatan", "Canvassing TIDAK tercatat pada mesin absensi. Data ini adalah satu-satunya sumber untuk menilai performa sales."],
    ]
    for row in guide_rows:
        guide.append(row)
    guide["A1"].font = Font(name="Segoe UI", size=14, bold=True, color=WHITE)
    guide["A1"].fill = header_fill
    guide.merge_cells("A1:B1")
    for row in guide.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    guide.column_dimensions["A"].width = 20
    guide.column_dimensions["B"].width = 90

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
