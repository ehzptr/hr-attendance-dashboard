"""
hrdash.leave
============
Modul terintegrasi untuk menangani template Excel dan pemrosesan data Cuti,
Lembur & Kegiatan Bersama (Lembur Bersama / Libur Bersama).
"""

from __future__ import annotations

import datetime as dt
import io
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
import polars as pl

NAVY = "1F4E78"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"

# --------------------------------------------------------------------------
# Constant & Enum Definitions
# --------------------------------------------------------------------------

SHEET_CUTI_IZIN = "Cuti_Izin"
SHEET_LEMBUR_KARYAWAN = "Lembur_Karyawan"
SHEET_KEGIATAN_BERSAMA = "Kegiatan_Bersama"


class ApprovalStatus(str, Enum):
    APPROVED = "Approved"
    PENDING = "Pending"
    REJECTED = "Rejected"

    @classmethod
    def parse(cls, raw: object) -> Optional["ApprovalStatus"]:
        if raw is None:
            return None
        text = str(raw).strip().lower()
        for member in cls:
            if member.value.lower() == text:
                return member
        if text in ("ya", "yes", "setuju", "disetujui", "ok", "1", "true"):
            return cls.APPROVED
        if text in ("tidak", "no", "ditolak", "reject"):
            return cls.REJECTED
        return None


class GroupEventType(str, Enum):
    LIBUR_BERSAMA = "Libur Bersama"
    LEMBUR_BERSAMA = "Lembur Bersama"

    @classmethod
    def parse(cls, raw: object) -> Optional["GroupEventType"]:
        if raw is None:
            return None
        text = str(raw).strip().lower()
        if "lembur" in text or "event" in text or "kegiatan" in text or "pameran" in text:
            return cls.LEMBUR_BERSAMA
        if "libur" in text or "cuti" in text:
            return cls.LIBUR_BERSAMA
        for member in cls:
            if member.value.lower() == text:
                return member
        return None


# --------------------------------------------------------------------------
# Dataclasses
# --------------------------------------------------------------------------

@dataclass(frozen=True)
class LeaveRecord:
    employee_id: str
    name: Optional[str]
    start_date: dt.date
    end_date: dt.date
    leave_type: str
    status: ApprovalStatus
    note: Optional[str]
    source_row: int

    def dates(self) -> list[dt.date]:
        return _date_range(self.start_date, self.end_date)


@dataclass(frozen=True)
class OvertimeRecord:
    employee_id: str
    name: Optional[str]
    date: dt.date
    start_time: dt.time
    end_time: dt.time
    duration_hours: float
    status: ApprovalStatus
    note: Optional[str]
    source_row: int


@dataclass(frozen=True)
class GroupEvent:
    start_date: dt.date
    end_date: dt.date
    name: str
    event_type: GroupEventType
    duration_hours: float = 8.0
    scope: str = "Semua Karyawan"
    note: Optional[str] = None
    source_row: int = 0

    def dates(self) -> list[dt.date]:
        return _date_range(self.start_date, self.end_date)


@dataclass(frozen=True)
class RejectedRow:
    sheet: str
    row: int
    reason: str
    raw: dict


@dataclass
class LeaveLoadResult:
    leaves: list[LeaveRecord] = field(default_factory=list)
    overtimes: list[OvertimeRecord] = field(default_factory=list)
    group_events: list[GroupEvent] = field(default_factory=list)
    rejected: list[RejectedRow] = field(default_factory=list)


# --------------------------------------------------------------------------
# Template Generator
# --------------------------------------------------------------------------

def build_leave_template() -> bytes:
    """Membuat template Excel 3 sheet profesional:
    1. Cuti_Izin (per karyawan)
    2. Lembur_Karyawan (per karyawan)
    3. Kegiatan_Bersama (global: Lembur Bersama / Libur Bersama)
    4. Petunjuk
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    header_fill = PatternFill("solid", fgColor=NAVY)
    header_font = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
    body_font = Font(name="Segoe UI", size=10, color="1F2937")
    thin = Side(style="thin", color="D1D5DB")

    def _style_sheet(ws, widths: dict[str, int]):
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(bottom=thin)
        for row in ws.iter_rows(min_row=2, max_row=500, max_col=len(widths)):
            for cell in row:
                cell.font = body_font
                cell.border = Border(bottom=thin)
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w
        ws.freeze_panes = "A2"
        last_col_letter = list(widths.keys())[-1]
        ws.auto_filter.ref = f"A1:{last_col_letter}500"

    # 1. Cuti_Izin
    ws_cuti = wb.create_sheet(title=SHEET_CUTI_IZIN)
    ws_cuti.append(["No.", "Nama", "Tanggal Mulai", "Tanggal Selesai", "Jumlah Hari", "Leave Type", "Status Approval", "Keterangan"])
    ws_cuti.append(["EMP001", "Budi Santoso", "2026-09-10", "2026-09-11", 2, "Cuti Tahunan", "Approved", "Cuti tahunan keperluan keluarga"])
    ws_cuti.append(["EMP002", "Sari Wijaya", "2026-09-18", "2026-09-18", 1, "Izin Sakit", "Approved", "Surat dokter terlampir"])
    _style_sheet(ws_cuti, {"A": 14, "B": 22, "C": 16, "D": 16, "E": 14, "F": 18, "G": 16, "H": 30})

    dv_status = DataValidation(type="list", formula1='"Approved,Pending,Rejected"', allow_blank=True)
    ws_cuti.add_data_validation(dv_status)
    dv_status.add("G2:G500")

    # 2. Lembur_Karyawan
    ws_lembur = wb.create_sheet(title=SHEET_LEMBUR_KARYAWAN)
    ws_lembur.append(["No.", "Nama", "Tanggal", "Jam Mulai Lembur", "Jam Selesai Lembur", "Durasi (Jam)", "Status Approval", "Keterangan / Alasan Lembur"])
    ws_lembur.append(["EMP001", "Budi Santoso", "2026-09-12", "17:00", "20:00", 3.0, "Approved", "Closing rekap akhir pekan"])
    ws_lembur.append(["EMP002", "Sari Wijaya", "2026-09-21", "17:00", "19:30", 2.5, "Approved", "Rekap payroll bulanan"])
    _style_sheet(ws_lembur, {"A": 14, "B": 22, "C": 16, "D": 18, "E": 18, "F": 14, "G": 16, "H": 35})

    dv_lembur_status = DataValidation(type="list", formula1='"Approved,Pending,Rejected"', allow_blank=True)
    ws_lembur.add_data_validation(dv_lembur_status)
    dv_lembur_status.add("G2:G500")

    # 3. Kegiatan_Bersama
    ws_kegiatan = wb.create_sheet(title=SHEET_KEGIATAN_BERSAMA)
    ws_kegiatan.append(["Tanggal Mulai", "Tanggal Selesai", "Nama Kegiatan / Event", "Jenis Pencatatan", "Durasi Lembur (Jam)", "Cakupan Karyawan", "Keterangan"])
    ws_kegiatan.append(["2026-09-15", "2026-09-15", "Pameran Launching Mall", "Lembur Bersama", 8.0, "Semua Karyawan", "Event pameran mobil di mall, karyawan tidak scan di kantor"])
    ws_kegiatan.append(["2026-12-24", "2026-12-24", "Cuti Bersama Nasional", "Libur Bersama", 0.0, "Semua Karyawan", "Hari libur / cuti bersama perusahaan"])
    _style_sheet(ws_kegiatan, {"A": 16, "B": 16, "C": 28, "D": 20, "E": 20, "F": 20, "G": 40})

    dv_jenis = DataValidation(type="list", formula1='"Lembur Bersama,Libur Bersama"', allow_blank=True)
    ws_kegiatan.add_data_validation(dv_jenis)
    dv_jenis.add("D2:D500")

    # 4. Petunjuk
    guide = wb.create_sheet("Petunjuk")
    guide.append(["PANDUAN TEMPLATE CUTI, LEMBUR & KEGIATAN BERSAMA"])
    guide_rows = [
        ["1. Sheet Cuti_Izin", "Digunakan untuk perizinan dan cuti individual per karyawan berdasarkan No. (Employee ID). Status 'Approved' akan membebaskan karyawan dari status Mangkir dan potongan mangkir."],
        ["2. Sheet Lembur_Karyawan", "Digunakan untuk pencatatan lembur individual di luar jam kantor normal per karyawan berdasarkan No. (Employee ID). Durasi lembur otomatis diakumulasikan."],
        ["3. Sheet Kegiatan_Bersama", "Berlaku global untuk seluruh karyawan (atau divisi terkait) tanpa memerlukan scan absensi fingerprint di kantor."],
        ["   - Jenis 'Lembur Bersama'", "Contoh: Pameran launching di Mall / pameran otomotif. Seluruh karyawan yang bertugas TIDAK dianggap mangkir meskipun tidak ada scan di kantor, dan jam lembur akan ditambahkan."],
        ["   - Jenis 'Libur Bersama'", "Contoh: Cuti bersama perusahaan atau libur khusus. Seluruh karyawan otomatis berstatus Libur Bersama dan tidak dianggap mangkir."],
        ["Format Tanggal & Jam", "Gunakan format YYYY-MM-DD untuk tanggal (contoh: 2026-09-15) dan HH:MM untuk jam (contoh: 17:00)."],
    ]
    for row in guide_rows:
        guide.append(row)

    guide["A1"].font = Font(name="Segoe UI", size=13, bold=True, color=WHITE)
    guide["A1"].fill = header_fill
    guide.merge_cells("A1:B1")
    for row in guide.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    guide.column_dimensions["A"].width = 28
    guide.column_dimensions["B"].width = 85

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# File Parser / Loader
# --------------------------------------------------------------------------

def _find_sheet(wb: openpyxl.Workbook, aliases: list[str]) -> Optional[str]:
    cleaned = {s.lower().replace(" ", "").replace("_", "").replace("/", ""): s for s in wb.sheetnames}
    for alias in aliases:
        key = alias.lower().replace(" ", "").replace("_", "").replace("/", "")
        if key in cleaned:
            return cleaned[key]
    for alias in aliases:
        key = alias.lower().replace(" ", "").replace("_", "").replace("/", "")
        for k, orig in cleaned.items():
            if key in k or k in key:
                return orig
    return None


def load_leave_workbook(path: str | Path | object) -> LeaveLoadResult:
    """Membaca file Excel Cuti, Lembur & Kegiatan Bersama dengan toleransi variasi nama sheet."""
    if hasattr(path, "seek"):
        path.seek(0)

    result = LeaveLoadResult()
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:
        result.rejected.append(RejectedRow("Workbook", 1, f"File tidak dapat dibaca sebagai Excel: {exc}", {}))
        return result

    cuti_sheet = _find_sheet(wb, ["Cuti_Izin", "Cuti/Izin", "Cuti", "Data Cuti", "Izin", "Leave"])
    lembur_sheet = _find_sheet(wb, ["Lembur_Karyawan", "Lembur Karyawan", "Lembur", "Overtime"])
    kegiatan_sheet = _find_sheet(
        wb,
        [
            "Kegiatan_Bersama",
            "Kegiatan/Event/Lembur Bersama/Cuti Bersama",
            "Kegiatan Bersama",
            "Event",
            "Kegiatan",
            "Lembur Bersama",
            "Cuti Bersama",
        ],
    )

    # Fallback: if none matched by alias, check if active sheet is single-sheet cuti/leave table
    if not cuti_sheet and not lembur_sheet and not kegiatan_sheet and wb.sheetnames:
        first_ws = wb.active
        first_row = [str(c).lower() for c in next(first_ws.iter_rows(min_row=1, max_row=1, values_only=True), []) if c]
        if any("cuti" in c or "leave" in c or "tanggal" in c or "date" in c for c in first_row):
            cuti_sheet = first_ws.title

    if cuti_sheet and cuti_sheet in wb.sheetnames:
        _load_cuti_izin(wb[cuti_sheet], result)
    if lembur_sheet and lembur_sheet in wb.sheetnames:
        _load_lembur_karyawan(wb[lembur_sheet], result)
    if kegiatan_sheet and kegiatan_sheet in wb.sheetnames:
        _load_kegiatan_bersama(wb[kegiatan_sheet], result)

    return result


def _load_cuti_izin(ws, result: LeaveLoadResult) -> None:
    headers = [str(c).strip() if c else "" for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if _row_is_blank(row):
            continue
        emp_id, name, start_raw, end_raw, _, leave_type, status_raw, note = _pad(row, 8)
        raw = dict(zip(headers or ["No.", "Nama", "Mulai", "Selesai", "Jumlah", "Type", "Status", "Ket"], row))

        if _is_blank(emp_id):
            result.rejected.append(RejectedRow(ws.title, row_idx, "No. (Employee ID) kosong", raw))
            continue

        start_date = _to_date(start_raw)
        end_date = _to_date(end_raw) or start_date
        if start_date is None:
            result.rejected.append(RejectedRow(ws.title, row_idx, "Tanggal Mulai tidak valid", raw))
            continue

        status = ApprovalStatus.parse(status_raw) or ApprovalStatus.APPROVED
        if status is not ApprovalStatus.APPROVED:
            result.rejected.append(RejectedRow(ws.title, row_idx, f"Status '{status.value}' (Bukan Approved)", raw))
            continue

        result.leaves.append(
            LeaveRecord(
                employee_id=str(emp_id).strip(),
                name=str(name).strip() if not _is_blank(name) else None,
                start_date=start_date,
                end_date=end_date,
                leave_type=str(leave_type).strip() if not _is_blank(leave_type) else "Cuti/Izin",
                status=status,
                note=str(note).strip() if not _is_blank(note) else None,
                source_row=row_idx,
            )
        )


def _load_lembur_karyawan(ws, result: LeaveLoadResult) -> None:
    headers = [str(c).strip() if c else "" for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if _row_is_blank(row):
            continue
        emp_id, name, date_raw, start_raw, end_raw, dur_raw, status_raw, note = _pad(row, 8)
        raw = dict(zip(headers or ["No.", "Nama", "Tanggal", "Mulai", "Selesai", "Durasi", "Status", "Ket"], row))

        if _is_blank(emp_id):
            result.rejected.append(RejectedRow(ws.title, row_idx, "No. (Employee ID) kosong", raw))
            continue

        the_date = _to_date(date_raw)
        start_time = _to_time(start_raw)
        end_time = _to_time(end_raw)
        if the_date is None:
            result.rejected.append(RejectedRow(ws.title, row_idx, "Tanggal tidak valid", raw))
            continue

        # Parse duration
        duration = 0.0
        if dur_raw is not None and not _is_blank(dur_raw):
            try:
                duration = float(dur_raw)
            except (ValueError, TypeError):
                duration = 0.0

        if duration <= 0.0 and start_time and end_time:
            duration = _hours_between(start_time, end_time)

        if duration <= 0.0:
            duration = 1.0  # fallback 1 jam jika terisi

        status = ApprovalStatus.parse(status_raw) or ApprovalStatus.APPROVED
        if status is not ApprovalStatus.APPROVED:
            result.rejected.append(RejectedRow(ws.title, row_idx, f"Status '{status.value}' (Bukan Approved)", raw))
            continue

        result.overtimes.append(
            OvertimeRecord(
                employee_id=str(emp_id).strip(),
                name=str(name).strip() if not _is_blank(name) else None,
                date=the_date,
                start_time=start_time or dt.time(17, 0),
                end_time=end_time or dt.time(20, 0),
                duration_hours=round(duration, 2),
                status=status,
                note=str(note).strip() if not _is_blank(note) else None,
                source_row=row_idx,
            )
        )


def _load_kegiatan_bersama(ws, result: LeaveLoadResult) -> None:
    headers = [str(c).strip() if c else "" for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])]
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if _row_is_blank(row):
            continue
        start_raw, end_raw, name, jenis_raw, dur_raw, scope, note = _pad(row, 7)
        raw = dict(zip(headers or ["Mulai", "Selesai", "Event", "Jenis", "Durasi", "Cakupan", "Ket"], row))

        start_date = _to_date(start_raw)
        end_date = _to_date(end_raw) or start_date
        event_type = GroupEventType.parse(jenis_raw)

        if start_date is None or event_type is None or _is_blank(name):
            result.rejected.append(RejectedRow(ws.title, row_idx, "Data kegiatan tidak lengkap/valid", raw))
            continue

        duration_hours = 8.0 if event_type == GroupEventType.LEMBUR_BERSAMA else 0.0
        if dur_raw is not None and not _is_blank(dur_raw):
            try:
                duration_hours = float(dur_raw)
            except (ValueError, TypeError):
                pass

        result.group_events.append(
            GroupEvent(
                start_date=start_date,
                end_date=end_date,
                name=str(name).strip(),
                event_type=event_type,
                duration_hours=round(duration_hours, 2),
                scope=str(scope).strip() if not _is_blank(scope) else "Semua Karyawan",
                note=str(note).strip() if not _is_blank(note) else None,
                source_row=row_idx,
            )
        )


# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# Enrichment Engine Integrator
# --------------------------------------------------------------------------

def _ensure_polars(df: object) -> pl.DataFrame:
    if isinstance(df, pl.DataFrame):
        return df
    if hasattr(df, "to_dict"):
        return pl.from_pandas(df)
    if isinstance(df, list):
        return pl.DataFrame(df)
    return pl.DataFrame(df)


def apply_leave_kegiatan(daily_input: pl.DataFrame | object, leave_result: LeaveLoadResult | None) -> pl.DataFrame:
    """Mengintegrasikan data Cuti, Lembur, dan Kegiatan Bersama ke Polars DataFrame kehadiran utama."""
    daily = _ensure_polars(daily_input)
    if daily.is_empty() or leave_result is None:
        return daily

    # Lookups
    leave_lookup = {(l.employee_id, d): l for l in leave_result.leaves for d in l.dates()}
    ot_lookup = {}
    for ot in leave_result.overtimes:
        ot_lookup.setdefault((ot.employee_id, ot.date), []).append(ot)

    holiday_events = {
        d: e
        for e in leave_result.group_events
        if e.event_type == GroupEventType.LIBUR_BERSAMA
        for d in e.dates()
    }
    overtime_events = {
        d: e
        for e in leave_result.group_events
        if e.event_type == GroupEventType.LEMBUR_BERSAMA
        for d in e.dates()
    }

    id_col = "No." if "No." in daily.columns else "Employee_ID"
    date_col = "Tanggal" if "Tanggal" in daily.columns else "Date"

    rows = daily.to_dicts()

    for row in rows:
        emp_id = str(row.get(id_col, "")).strip()
        the_date = _to_date(row.get(date_col))
        if not emp_id or the_date is None:
            continue

        has_scan = int(row.get("Scan_Count", 0) or 0) > 0

        # Overtime Individual
        ot_entries = ot_lookup.get((emp_id, the_date), [])
        if ot_entries:
            row["Overtime_Hours"] = round(sum(o.duration_hours for o in ot_entries), 2)
            row["Overtime_Status"] = "Approved"

        # 1. Libur Bersama (Global)
        if the_date in holiday_events:
            event = holiday_events[the_date]
            row["Group_Event"] = event.name
            if "Absent_Flag" in row:
                row["Absent_Flag"] = 0
            if not has_scan:
                if "Status_Masuk" in row:
                    row["Status_Masuk"] = "Libur Bersama"
                if "Status_Pulang" in row:
                    row["Status_Pulang"] = "Libur Bersama"
                if "Menit_Telat" in row:
                    row["Menit_Telat"] = 0
                if "Menit_Pulang_Cepat" in row:
                    row["Menit_Pulang_Cepat"] = 0
                if "Anomaly_Type" in row:
                    row["Anomaly_Type"] = ""
                if "Severity" in row:
                    row["Severity"] = ""
                if "Is_Anomali" in row:
                    row["Is_Anomali"] = 0
                if "Late_Flag" in row:
                    row["Late_Flag"] = 0
                if "Early_Leave_Flag" in row:
                    row["Early_Leave_Flag"] = 0
                if "Forgot_Punch_Flag" in row:
                    row["Forgot_Punch_Flag"] = 0
                if "Tidak_Absen_Pulang_Flag" in row:
                    row["Tidak_Absen_Pulang_Flag"] = 0
            continue

        # 2. Approved Cuti (Individual)
        leave = leave_lookup.get((emp_id, the_date))
        if leave:
            row["Leave_Type"] = leave.leave_type
            if "Absent_Flag" in row:
                row["Absent_Flag"] = 0
            if not has_scan:
                if "Status_Masuk" in row:
                    row["Status_Masuk"] = "Cuti Disetujui"
                if "Status_Pulang" in row:
                    row["Status_Pulang"] = leave.leave_type
                if "Menit_Telat" in row:
                    row["Menit_Telat"] = 0
                if "Menit_Pulang_Cepat" in row:
                    row["Menit_Pulang_Cepat"] = 0
                if "Anomaly_Type" in row:
                    row["Anomaly_Type"] = ""
                if "Severity" in row:
                    row["Severity"] = ""
                if "Is_Anomali" in row:
                    row["Is_Anomali"] = 0
                if "Late_Flag" in row:
                    row["Late_Flag"] = 0
                if "Early_Leave_Flag" in row:
                    row["Early_Leave_Flag"] = 0
                if "Forgot_Punch_Flag" in row:
                    row["Forgot_Punch_Flag"] = 0
                if "Tidak_Absen_Pulang_Flag" in row:
                    row["Tidak_Absen_Pulang_Flag"] = 0
            continue

        # 3. Lembur Bersama (Global - e.g. pameran launching mall)
        if the_date in overtime_events:
            event = overtime_events[the_date]
            row["Group_Event"] = event.name
            current_ot = float(row.get("Overtime_Hours") or 0.0)
            row["Overtime_Hours"] = round(current_ot + (event.duration_hours or 8.0), 2)
            row["Overtime_Status"] = "Approved"

            # Karyawan ikut pameran/event sehingga tidak absen di kantor -> TIDAK MANGKIR!
            if not has_scan:
                if "Absent_Flag" in row:
                    row["Absent_Flag"] = 0
                if "Status_Masuk" in row:
                    row["Status_Masuk"] = "Lembur Bersama"
                if "Status_Pulang" in row:
                    row["Status_Pulang"] = "Lembur Bersama"
                if "Menit_Telat" in row:
                    row["Menit_Telat"] = 0
                if "Menit_Pulang_Cepat" in row:
                    row["Menit_Pulang_Cepat"] = 0
                if "Anomaly_Type" in row:
                    row["Anomaly_Type"] = ""
                if "Severity" in row:
                    row["Severity"] = ""
                if "Is_Anomali" in row:
                    row["Is_Anomali"] = 0
                if "Late_Flag" in row:
                    row["Late_Flag"] = 0
                if "Early_Leave_Flag" in row:
                    row["Early_Leave_Flag"] = 0
                if "Forgot_Punch_Flag" in row:
                    row["Forgot_Punch_Flag"] = 0
                if "Tidak_Absen_Pulang_Flag" in row:
                    row["Tidak_Absen_Pulang_Flag"] = 0
                if "Present_Flag" in row:
                    row["Present_Flag"] = 1

    return pl.DataFrame(rows, schema=daily.schema)


# --------------------------------------------------------------------------
# Helper Functions
# --------------------------------------------------------------------------

def _pad(row: tuple, length: int) -> tuple:
    return row[:length] if len(row) >= length else row + (None,) * (length - len(row))

def _is_blank(val: object) -> bool:
    return val is None or (isinstance(val, str) and val.strip() == "")

def _row_is_blank(row: tuple) -> bool:
    return all(_is_blank(v) for v in row)

def _to_date(val: object) -> Optional[dt.date]:
    if val is None:
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", "nat", "<null>"):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None

def _to_time(val: object) -> Optional[dt.time]:
    if val is None:
        return None
    if isinstance(val, dt.datetime):
        return val.time()
    if isinstance(val, dt.time):
        return val
    s = str(val).strip().replace(".", ":")
    if not s or s.lower() in ("nan", "none", "nat", "<null>"):
        return None
    parts = s.split(":")
    if len(parts) >= 2:
        try:
            return dt.time(int(parts[0]), int(parts[1]))
        except ValueError:
            return None
    return None

def _hours_between(start: dt.time, end: dt.time) -> float:
    s = start.hour + start.minute / 60.0
    e = end.hour + end.minute / 60.0
    return e - s if e >= s else (e + 24) - s

def _date_range(start: dt.date, end: dt.date) -> list[dt.date]:
    return [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]