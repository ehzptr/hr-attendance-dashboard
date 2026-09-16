"""
Configuration objects for the HR dashboard.

Every business rule that can plausibly change (schedules, tolerance, payroll
rates, score thresholds, sales scoring weights) lives here instead of being
hard-coded inside the calculation engines. Nothing in this module invents a
monetary/business rule that HR has not supplied: payroll rates default to
``None`` and the payroll engine refuses to fabricate a deduction amount when
a rate is missing (see ``hrdash.payroll``).
"""

from __future__ import annotations

from dataclasses import dataclass, field, fields
from datetime import date, time
from typing import Optional

import pandas as pd

CONFIG_SHEET_NAME = "CONFIG"

# ---------------------------------------------------------------------------
# Attendance schedule configuration
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class AttendanceConfig:
    """Working-hours / tolerance rules. Preserves the original semantics:
    first scan of the day = clock-in, last scan of the day = clock-out.
    """

    clock_in: time = time(7, 55)
    weekday_clock_out: time = time(17, 0)
    saturday_clock_out: time = time(14, 0)
    grace_minutes: int = 10
    single_scan_cutoff: time = time(12, 0)
    sunday_is_off: bool = True
    saturday_is_working: bool = True
    # Used only to flag an obvious missing-punch case where the *only* scan
    # of the day happens to have fallen before the cutoff (i.e. read as an
    # unmatched clock-in). Kept configurable rather than assumed.
    default_late_minutes_when_only_clock_out: int = 60
    start_date: Optional[date] = None
    end_date: Optional[date] = None

    @property
    def working_days(self) -> set[str]:
        days = {"Senin", "Selasa", "Rabu", "Kamis", "Jumat"}
        if self.saturday_is_working:
            days.add("Sabtu")
        if not self.sunday_is_off:
            days.add("Minggu")
        return days


# ---------------------------------------------------------------------------
# HR classification / scoring thresholds (Section 6)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ScoreConfig:
    """Attendance-score classification bands. Configurable, not hard-coded."""

    excellent_threshold: float = 90.0
    good_threshold: float = 80.0
    watch_threshold: float = 70.0

    def classify(self, score: float) -> str:
        if score >= self.excellent_threshold:
            return "EXCELLENT"
        if score >= self.good_threshold:
            return "GOOD"
        if score >= self.watch_threshold:
            return "WATCH"
        return "ATTENTION"


# ---------------------------------------------------------------------------
# Payroll configuration (Section 7)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PayrollConfig:
    """Payroll deduction rules.

    All rates default to ``None`` on purpose: the spec explicitly forbids
    assuming a deduction nominal that HR has not defined. When a rate is
    ``None`` the payroll engine reports the deduction as "Belum Diatur"
    (not configured) and excludes it from the total instead of guessing.
    """

    late_deduction_mode: str = "per_minute"  # "per_minute" or "per_occurrence"
    late_deduction_per_minute: Optional[float] = None
    late_deduction_per_occurrence: Optional[float] = None
    sales_no_clock_out_deduction: Optional[float] = None
    early_leave_deduction_per_minute: Optional[float] = None
    absence_deduction_per_day: Optional[float] = None
    other_deduction_note: str = ""

    @property
    def is_configured(self) -> bool:
        return any(
            v is not None
            for v in (
                self.late_deduction_per_minute,
                self.late_deduction_per_occurrence,
                self.sales_no_clock_out_deduction,
                self.early_leave_deduction_per_minute,
                self.absence_deduction_per_day,
            )
        )


# ---------------------------------------------------------------------------
# Sales performance scoring configuration (Section 4)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SalesScoreConfig:
    """Weights used to blend sales KPIs into a single Performance Score.

    Weights must sum to a positive number; they are normalised internally.
    Defaults give equal-ish emphasis to funnel volume and outcomes and are
    clearly documented in the README sheet so HR can tune them without
    touching code.
    """

    weight_visit: float = 0.15
    weight_test_drive: float = 0.15
    weight_spk: float = 0.30
    weight_delivery: float = 0.25
    weight_conversion_rate: float = 0.15

    excellent_threshold: float = 85.0
    good_threshold: float = 70.0
    attention_threshold: float = 50.0

    def classify(self, score: float) -> str:
        if score >= self.excellent_threshold:
            return "Excellent"
        if score >= self.good_threshold:
            return "Good"
        if score >= self.attention_threshold:
            return "Watch"
        return "Attention"


@dataclass(frozen=True)
class AppConfig:
    attendance: AttendanceConfig = field(default_factory=AttendanceConfig)
    score: ScoreConfig = field(default_factory=ScoreConfig)
    payroll: PayrollConfig = field(default_factory=PayrollConfig)
    sales_score: SalesScoreConfig = field(default_factory=SalesScoreConfig)


# ---------------------------------------------------------------------------
# CONFIG sheet <-> dataclasses round-trip
# ---------------------------------------------------------------------------

_TIME_FIELDS = {"clock_in", "weekday_clock_out", "saturday_clock_out", "single_scan_cutoff"}
_DATE_FIELDS = {"start_date", "end_date"}
_BOOL_FIELDS = {"sunday_is_off", "saturday_is_working"}

_SECTION_LABELS = {
    "attendance": "Jadwal & Toleransi",
    "score": "Klasifikasi Skor Kehadiran",
    "payroll": "Payroll",
    "sales_score": "Bobot Skor Performa Sales",
}

_FIELD_DESCRIPTIONS = {
    "clock_in": "Jam masuk resmi",
    "weekday_clock_out": "Jam pulang Senin-Jumat",
    "saturday_clock_out": "Jam pulang Sabtu",
    "grace_minutes": "Toleransi keterlambatan (menit)",
    "single_scan_cutoff": "Batas waktu untuk membedakan scan tunggal pagi/sore",
    "sunday_is_off": "Minggu libur? (Ya/Tidak)",
    "saturday_is_working": "Sabtu hari kerja? (Ya/Tidak)",
    "default_late_minutes_when_only_clock_out": "Estimasi menit telat saat hanya ada scan pulang",
    "start_date": "Tanggal mulai periode perhitungan (kosongkan = otomatis dari data)",
    "end_date": "Tanggal akhir periode perhitungan (kosongkan = otomatis dari data)",
    "excellent_threshold": "Ambang batas skor EXCELLENT / Good",
    "good_threshold": "Ambang batas skor GOOD",
    "watch_threshold": "Ambang batas skor WATCH (di bawah ini = ATTENTION)",
    "late_deduction_mode": "Metode potongan telat (per_minute / per_occurrence)",
    "late_deduction_per_minute": "Tarif potongan per menit terlambat (kosongkan jika belum ditentukan)",
    "late_deduction_per_occurrence": "Tarif potongan per 1x terlambat (kosongkan jika belum ditentukan)",
    "sales_no_clock_out_deduction": "Tarif potongan per 1x tidak absen pulang tim Sales (kosongkan jika belum ditentukan)",
    "early_leave_deduction_per_minute": "Tarif potongan per menit pulang cepat (kosongkan jika belum ditentukan)",
    "absence_deduction_per_day": "Potongan per hari mangkir (kosongkan jika belum ditentukan)",
    "other_deduction_note": "Catatan potongan lain-lain",
    "weight_visit": "Bobot Customer Visit",
    "weight_test_drive": "Bobot Test Drive",
    "weight_spk": "Bobot SPK",
    "weight_delivery": "Bobot Delivery",
    "weight_conversion_rate": "Bobot Conversion Rate",
    "attention_threshold": "Ambang batas skor Watch/Attention",
}


def _fmt_value(name: str, value) -> object:
    if value is None:
        return ""
    if name in _TIME_FIELDS and isinstance(value, time):
        return value.strftime("%H:%M")
    if name in _DATE_FIELDS and isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if name in _BOOL_FIELDS:
        return "Ya" if value else "Tidak"
    return value


def config_to_dataframe(cfg: AppConfig) -> pd.DataFrame:
    """Flatten AppConfig into a human-editable Parameter/Value table."""
    rows: list[dict] = []
    for section_name in ("attendance", "score", "payroll", "sales_score"):
        section_obj = getattr(cfg, section_name)
        for f in fields(section_obj):
            rows.append(
                {
                    "Section": _SECTION_LABELS[section_name],
                    "Parameter": f.name,
                    "Value": _fmt_value(f.name, getattr(section_obj, f.name)),
                    "Description": _FIELD_DESCRIPTIONS.get(f.name, ""),
                }
            )
    return pd.DataFrame(rows, columns=["Section", "Parameter", "Value", "Description"])


def _parse_value(name: str, raw: object):
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return None
    if name in _TIME_FIELDS:
        if isinstance(raw, time):
            return raw
        text = str(raw).strip()
        parts = text.replace(".", ":").split(":")
        return time(int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)
    if name in _DATE_FIELDS:
        if isinstance(raw, date):
            return raw
        return pd.Timestamp(str(raw)).date()
    if name in _BOOL_FIELDS:
        if isinstance(raw, bool):
            return raw
        return str(raw).strip().lower() in {"ya", "yes", "true", "1"}
    if isinstance(raw, str):
        raw = raw.strip()
        if raw == "":
            return None
    try:
        return float(raw) if "." in str(raw) else int(raw)
    except (ValueError, TypeError):
        return raw


def config_from_dataframe(df: pd.DataFrame) -> AppConfig:
    """Rebuild an AppConfig from a previously exported CONFIG sheet.

    Unknown or malformed rows are ignored; missing rows fall back to
    dataclass defaults so a partial CONFIG sheet never crashes the app.
    """
    values: dict[str, dict] = {"attendance": {}, "score": {}, "payroll": {}, "sales_score": {}}
    section_by_label = {v: k for k, v in _SECTION_LABELS.items()}

    for _, row in df.iterrows():
        section_key = section_by_label.get(str(row.get("Section", "")).strip())
        param = str(row.get("Parameter", "")).strip()
        if not section_key or not param:
            continue
        try:
            values[section_key][param] = _parse_value(param, row.get("Value"))
        except Exception:
            continue

    def build(cls, overrides: dict):
        valid = {f.name for f in fields(cls)}
        kwargs = {k: v for k, v in overrides.items() if k in valid and v is not None}
        return cls(**kwargs)

    return AppConfig(
        attendance=build(AttendanceConfig, values["attendance"]),
        score=build(ScoreConfig, values["score"]),
        payroll=build(PayrollConfig, values["payroll"]),
        sales_score=build(SalesScoreConfig, values["sales_score"]),
    )
