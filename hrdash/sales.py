"""
Sales performance module (Section 4).

Sales performance is deliberately kept as a *separate* data source and
pipeline from attendance. Nothing here reads scan counts, first/last scan
times, or any other fingerprint-derived value — canvassing is not observable
on the attendance machine (Section 3) and this module never pretends it is.
"""

from __future__ import annotations

import io
from typing import Optional

import polars as pl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .config import SalesScoreConfig
from .attendance.parser import parse_datetime_expr

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


def _clean_col(name: object) -> str:
    return str(name).replace("\n", " ").strip()


def clean_sales_activity(raw_input: Optional[pl.DataFrame]) -> pl.DataFrame:
    """Validate and normalize an uploaded Sales Activity file.

    Required columns: employee ID and date. All KPI columns are optional
    and default to 0 when absent, since a sales rep may have zero of a
    given activity on a given day — that is data, not a missing value.
    """
    if raw_input is None or raw_input.is_empty():
        return pl.DataFrame(
            schema={
                "No.": pl.String,
                "Tanggal": pl.Date,
                "Prospect": pl.Float64,
                "Visit": pl.Float64,
                "Test_Drive": pl.Float64,
                "SPK": pl.Float64,
                "Delivery": pl.Float64,
                "Revenue": pl.Float64,
            }
        )

    df = raw_input
    rename_map = {c: _clean_col(c) for c in df.columns}
    df = df.rename(rename_map)

    id_col = next((c for c in ["No.", "Employee ID", "ID", "NIK"] if c in df.columns), None)
    date_col = next((c for c in ["Tanggal", "Date"] if c in df.columns), None)
    if not id_col or not date_col:
        raise SalesActivityError(
            "File Sales Activity harus memiliki kolom employee ID (No./Employee ID) "
            "dan tanggal (Tanggal/Date)."
        )

    kpi_map = {
        "Prospect": ["Prospect", "Prospek"],
        "Visit": ["Visit", "Customer Visit"],
        "Test_Drive": ["Test Drive", "Test_Drive"],
        "SPK": ["SPK"],
        "Delivery": ["Delivery"],
        "Revenue": ["Revenue"],
    }

    select_exprs = [
        pl.col(id_col).cast(pl.String).str.strip_chars().alias("No."),
        parse_datetime_expr(date_col).dt.date().alias("Tanggal"),
    ]

    for target, candidates in kpi_map.items():
        matched = next((c for c in candidates if c in df.columns), None)
        if matched:
            select_exprs.append(pl.col(matched).cast(pl.Float64, strict=False).fill_null(0.0).alias(target))
        else:
            select_exprs.append(pl.lit(0.0, dtype=pl.Float64).alias(target))

    out = df.select(select_exprs)
    out = out.filter(pl.col("No.").is_not_null() & (pl.col("No.") != "") & pl.col("Tanggal").is_not_null())
    return out.select(SALES_ACTIVITY_COLUMNS)


def _normalize_series(s: pl.Series) -> list[float]:
    vals = s.to_list()
    valid_vals = [v for v in vals if v is not None]
    if not valid_vals:
        return [0.0] * len(vals)
    lo = min(valid_vals)
    hi = max(valid_vals)
    if hi == lo:
        return [100.0 if (v is not None and v > 0) else 0.0 for v in vals]
    return [round((v - lo) / (hi - lo) * 100.0, 4) if v is not None else 0.0 for v in vals]


def calculate_sales_performance(
    sales_activity: Optional[pl.DataFrame],
    master: Optional[pl.DataFrame],
    cfg: SalesScoreConfig,
) -> pl.DataFrame:
    """Aggregate Sales Activity into per-employee sales KPIs (Section 4).

    Performance_Score blends KPI volume (Visit/Test Drive/SPK/Delivery) and
    Conversion Rate using HR-configurable weights (``SalesScoreConfig``),
    each metric relatively normalised (0-100) across the current sales team
    for the selected period, since no fixed company-wide sales target was
    supplied.
    """
    if sales_activity is None or sales_activity.is_empty():
        return pl.DataFrame(
            schema={
                "No.": pl.String,
                "Name": pl.String,
                "Department": pl.String,
                "Prospect": pl.Float64,
                "Visit": pl.Float64,
                "Test_Drive": pl.Float64,
                "SPK": pl.Float64,
                "Delivery": pl.Float64,
                "Revenue": pl.Float64,
                "Conversion_Rate_%": pl.Float64,
                "Performance_Score": pl.Float64,
                "Performance_Classification": pl.String,
            }
        )

    agg = sales_activity.group_by("No.", maintain_order=True).agg(
        Prospect=pl.col("Prospect").cast(pl.Float64).sum(),
        Visit=pl.col("Visit").cast(pl.Float64).sum(),
        Test_Drive=pl.col("Test_Drive").cast(pl.Float64).sum(),
        SPK=pl.col("SPK").cast(pl.Float64).sum(),
        Delivery=pl.col("Delivery").cast(pl.Float64).sum(),
        Revenue=pl.col("Revenue").cast(pl.Float64).sum(),
    )

    agg = agg.with_columns(
        **{
            "Conversion_Rate_%": pl.when(pl.col("Prospect") > 0)
            .then(pl.col("SPK") / pl.col("Prospect") * 100.0)
            .otherwise(0.0)
        }
    )

    if master is not None and not master.is_empty() and "No." in master.columns:
        m_sub = master.select(
            [c for c in ["No.", "Name", "Department"] if c in master.columns]
        ).unique(subset=["No."])
        agg = agg.join(m_sub, on="No.", how="left")
        agg = agg.with_columns(
            Name=pl.when(pl.col("Name").is_not_null() & (pl.col("Name") != ""))
            .then(pl.col("Name"))
            .otherwise(pl.concat_str([pl.lit("Sales "), pl.col("No.")])),
            Department=pl.col("Department").fill_null("Belum Dipetakan"),
        )
    else:
        agg = agg.with_columns(
            Name=pl.concat_str([pl.lit("Sales "), pl.col("No.")]),
            Department=pl.lit("Belum Dipetakan", dtype=pl.String),
        )

    weights = {
        "Visit": cfg.weight_visit,
        "Test_Drive": cfg.weight_test_drive,
        "SPK": cfg.weight_spk,
        "Delivery": cfg.weight_delivery,
        "Conversion_Rate_%": cfg.weight_conversion_rate,
    }
    total_weight = sum(weights.values()) or 1.0

    scores = [0.0] * len(agg)
    for col, w in weights.items():
        norm_list = _normalize_series(agg[col])
        factor = w / total_weight
        for i, val in enumerate(norm_list):
            scores[i] += val * factor

    scores = [round(s, 1) for s in scores]
    classifications = [cfg.classify(s) for s in scores]

    agg = agg.with_columns(
        Performance_Score=pl.Series("Performance_Score", scores, dtype=pl.Float64),
        Performance_Classification=pl.Series("Performance_Classification", classifications, dtype=pl.String),
    )

    return agg.select(SALES_PERFORMANCE_COLUMNS).sort("Performance_Score", descending=True)


def combine_attendance_and_sales(
    employee_summary: Optional[pl.DataFrame], sales_performance: Optional[pl.DataFrame]
) -> pl.DataFrame:
    """Management view (Section 4 example table): attendance discipline
    next to sales outcomes, joined on Employee ID only — the two data
    sources are never blended at the raw-data level.
    """
    if employee_summary is None or employee_summary.is_empty() or sales_performance is None or sales_performance.is_empty():
        return pl.DataFrame()

    df_summary = employee_summary
    df_sales = sales_performance

    att_cols = [
        "No.", "Name", "Department", "Attendance_Rate_%", "Terlambat", "Attendance_Score", "HR_Classification"
    ]
    avail_att = [c for c in att_cols if c in df_summary.columns]
    att = df_summary.select(avail_att)

    att_renamed = att.rename({c: f"{c}_att" for c in avail_att if c != "No."})
    combined = df_sales.join(att_renamed, on="No.", how="left")

    if "Name_att" in combined.columns:
        combined = combined.with_columns(
            Name=pl.coalesce([pl.col("Name"), pl.col("Name_att")])
        ).drop("Name_att")

    if "Department_att" in combined.columns:
        combined = combined.with_columns(
            Department=pl.coalesce([pl.col("Department"), pl.col("Department_att")])
        ).drop("Department_att")

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
