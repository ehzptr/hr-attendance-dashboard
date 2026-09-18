import random
from datetime import datetime, timedelta
import polars as pl
import xlsxwriter

# =====================================================================
# 1. KONFIGURASI PARAMETER GLOBAL & PROBABILITAS (DAPAT DIATUR)
# =====================================================================

# --- General & Periode Waktu ---
NUM_EMPLOYEES = 50
START_DATE = datetime(2026, 8, 21)
END_DATE = datetime(2026, 9, 13)

# --- Probabilitas Master Karyawan ---
PROB_ACTIVE_EMPLOYEE = 0.98    # 98% karyawan aktif ('Y'), 2% resign ('N')
PROB_MIDDLE_NAME = 0.40        # 40% karyawan memiliki 3 kata nama (middle name)

# --- Probabilitas Log Absensi ---
ABSENCE_RATE = 0.05            # 5% Tidak hadir per hari (Sakit/Cuti/Alpha)
FORGOT_IN_RATE = 0.02          # 2% Lupa absen masuk
FORGOT_OUT_RATE = 0.03         # 3% Lupa absen pulang
DOUBLE_SCAN_RATE = 0.03        # 3% Double scan (mesin baca 2x dalam jeda singkat)

# --- Probabilitas Sales Activity ---
PROB_SALES_ACTIVE_DAY = 0.70   # 70% peluang sales mencatat aktivitas pada hari kerja
PROB_SPK_SINGLE = 0.35         # 35% peluang dapat 1 SPK per hari aktif
PROB_SPK_DOUBLE = 0.05         # 5% peluang dapat 2 SPK per hari aktif
PROB_DELIVERY = 0.20           # 20% peluang ada penyerahan unit (Delivery) per hari aktif

# --- Probabilitas Template Cuti / Lembur / Event ---
PROB_CUTI_PER_EMP_DAY = 0.005  # ~0.5% peluang karyawan mengajukan Cuti/Izin per hari
PROB_LEMBUR_PER_EMP_DAY = 0.010 # ~1.0% peluang karyawan mengajukan Lembur per hari
NUM_KEGIATAN_BERSAMA = 3       # Jumlah event / kegiatan bersama yang di-generate

# --- Nama File Output ---
FILE_MASTER = "dummy_data-master_karyawan_dealership.xlsx"
FILE_LOG_ABSENSI = "dummy_data-log_absensi_dealership_2026.xlsx"
FILE_SALES_ACTIVITY = "dummy_data-sales_activity_dealership.xlsx"
FILE_TEMPLATE = "dummy_data-template_cuti_lembur_event.xlsx"

# =====================================================================
# DATA REFERENSI PERUSAHAAN DEALERSHIP
# =====================================================================
DEPARTMENTS_CONFIG = [
    ("Sales Mobil Baru", ["Sales Executive", "Senior Sales Executive", "Sales Supervisor", "Branch Manager"], "SALES", "08:30", "17:30"),
    ("Sales Mobil Bekas", ["Sales Executive", "Assessor Mobil Bekas", "Sales Supervisor"], "SALES", "08:30", "17:30"),
    ("Service & Workshop", ["Service Advisor", "Teknisi / Mekanik", "Kepala Bengkel"], "OFFICE", "08:00", "17:00"),
    ("HRD & Admin", ["Staff HRD", "Staff Admin", "HR Manager", "General Affairs"], "OFFICE", "08:00", "17:00"),
    ("Finance & Accounting", ["Staff Finance", "Staff Accounting", "Kasir", "Finance Manager"], "OFFICE", "08:00", "17:00"),
    ("Customer Relations", ["CRM Officer", "Telemarketing", "CRO Manager"], "OFFICE", "08:00", "17:00"),
    ("Spareparts & Accessories", ["Staff Sparepart", "Warehouse Staff", "Supervisor Sparepart"], "OFFICE", "08:00", "17:00")
]

FIRST_NAMES = ["Budi", "Sari", "Agus", "Dewi", "Eko", "Ani", "Rudi", "Rina", "Dedi", "Maya", "Ahmad", "Dian", "Fajar", "Indah", "Heri", "Joko", "Fitri", "Taufik", "Wulan", "Doni"]
MIDDLE_NAMES = ["Pratama", "Kusuma", "Santoso", "Putra", "Putri", "Wibowo", "Kurniawan", "Suryadi", "Firmansyah", "Permana", "Lestari", "Setiawan", "Hidayat", "Saputra"]
LAST_NAMES = ["Wijaya", "Santoso", "Hidayat", "Saputra", "Utami", "Subakti", "Susanto", "Gunawan", "Nugroho", "Suryani", "Permana", "Kusuma", "Wibowo"]

CAR_PRICES = [250000000, 300000000, 350000000, 450000000, 600000000]

LEAVE_TYPES = ["Cuti Tahunan", "Izin Sakit", "Izin Alasan Penting", "Cuti Melahirkan", "Cuti Menikah"]
APPROVAL_STATUSES = ["Approved", "Approved", "Approved", "Pending", "Rejected"]

LEAVE_REASONS = {
    "Cuti Tahunan": ["Cuti tahunan keperluan keluarga", "Acara keluarga di luar kota", "Liburan keluarga", "Urusan pribadi"],
    "Izin Sakit": ["Surat dokter terlampir", "Demam dan flu", "Pemeriksaan kesehatan ke rumah sakit", "Istirahat sakit dokter"],
    "Izin Alasan Penting": ["Urusan administrasi perbankan/pemerintah", "Renovasi rumah emergency", "Acara keagamaan keluarga"],
    "Cuti Melahirkan": ["Persalinan dan perawatan bayi", "Istirahat melahirkan"],
    "Cuti Menikah": ["Pernikahan pribadi"]
}

OVERTIME_REASONS = [
    "Closing rekap akhir pekan", "Rekap payroll bulanan", "Stock opname suku cadang dealer",
    "Persiapan unit delivery ke konsumen", "Maintenance server & database kantor",
    "Persiapan event pameran mobil dealer", "Layanan emergency road service 24 jam"
]

EVENT_POOL = [
    ("Pameran Launching Mobil Baru Mall", "Lembur Bersama", 8, "Semua Karyawan", "Event pameran mobil di mall, karyawan tidak scan di kantor"),
    ("Cuti Bersama Nasional", "Libur Bersama", 0, "Semua Karyawan", "Hari libur / cuti bersama perusahaan"),
    ("Training Product Knowledge & Sales Pitch", "Lembur Bersama", 4, "Sales Mobil Baru & Sales Mobil Bekas", "Pelatihan produk baru pada akhir pekan"),
    ("Stock Opname Massal Kuartal III", "Lembur Bersama", 6, "Service & Spareparts & Warehouse", "Perhitungan fisik persediaan suku cadang"),
    ("Customer Gathering & Night Sale", "Lembur Bersama", 5, "Sales & Customer Relations", "Acara khusus konsumen reguler dealership")
]


def generate_random_name():
    """Membuat nama karyawan dengan variasi 2 kata atau 3 kata (middle name)."""
    if random.random() < PROB_MIDDLE_NAME:
        return f"{random.choice(FIRST_NAMES)} {random.choice(MIDDLE_NAMES)} {random.choice(LAST_NAMES)}"
    else:
        return f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"


def generate_master_employees(num_emp):
    """Membuat master data karyawan lengkap."""
    random.seed(42)
    employees = []
    for i in range(1, num_emp + 1):
        emp_id = f"EMP{i:04d}" if i % 2 != 0 else f"{12020000 + i}"
        name = generate_random_name()
        
        dept_info = random.choice(DEPARTMENTS_CONFIG)
        dept_name = dept_info[0]
        position = random.choice(dept_info[1])
        emp_type = dept_info[2]
        jam_masuk = dept_info[3]
        jam_pulang = dept_info[4]
        
        salary = random.choice([4500000, 5000000, 5500000, 6000000, 7500000, 9000000, 12000000, 15000000])
        active = "Y" if random.random() < PROB_ACTIVE_EMPLOYEE else "N"

        employees.append({
            "No.": emp_id,
            "Name": name,
            "Department": dept_name,
            "Position": position,
            "Employee_Type": emp_type,
            "Jam_Masuk": jam_masuk,
            "Jam_Pulang": jam_pulang,
            "Monthly_Salary": salary,
            "Active": active
        })
    return employees


def build_all_files():
    random.seed(42)
    
    # List Tanggal Periode
    current_date = START_DATE
    dates = []
    while current_date <= END_DATE:
        dates.append(current_date)
        current_date += timedelta(days=1)

    # -----------------------------------------------------------------
    # 1. GENERATE MASTER KARYAWAN
    # -----------------------------------------------------------------
    master_list = generate_master_employees(NUM_EMPLOYEES)
    df_master = pl.DataFrame(master_list)
    df_master.write_excel(FILE_MASTER)
    print(f"File '{FILE_MASTER}' berhasil dibuat ({df_master.height} baris).")

    # -----------------------------------------------------------------
    # 2. GENERATE LOG ABSENSI
    # -----------------------------------------------------------------
    raw_logs = []
    for emp in master_list:
        if emp["Active"] == "N":
            continue

        for d in dates:
            if random.random() < ABSENCE_RATE:
                continue

            # Absen Masuk
            if random.random() >= FORGOT_IN_RATE:
                minute_offset = int(random.gauss(0, 15))
                in_time = d.replace(hour=7) + timedelta(minutes=minute_offset, seconds=random.randint(0, 59))
                raw_logs.append({
                    "Department": emp["Department"],
                    "Name": emp["Name"],
                    "No.": emp["No."],
                    "timestamp": in_time
                })
                if random.random() < DOUBLE_SCAN_RATE:
                    raw_logs.append({
                        "Department": emp["Department"],
                        "Name": emp["Name"],
                        "No.": emp["No."],
                        "timestamp": in_time + timedelta(seconds=random.randint(5, 120))
                    })

            # Absen Pulang
            if random.random() >= FORGOT_OUT_RATE:
                minute_offset = int(random.gauss(15, 20))
                out_time = d.replace(hour=17) + timedelta(minutes=minute_offset, seconds=random.randint(0, 59))
                raw_logs.append({
                    "Department": emp["Department"],
                    "Name": emp["Name"],
                    "No.": emp["No."],
                    "timestamp": out_time
                })
                if random.random() < DOUBLE_SCAN_RATE:
                    raw_logs.append({
                        "Department": emp["Department"],
                        "Name": emp["Name"],
                        "No.": emp["No."],
                        "timestamp": out_time + timedelta(seconds=random.randint(5, 120))
                    })

    df_logs = pl.DataFrame(raw_logs).sort("timestamp")
    df_logs = df_logs.with_columns(
        pl.col("timestamp").dt.strftime("%d/%m/%Y %H.%M.%S").alias("Date/Time")
    ).select(["Department", "Name", "No.", "Date/Time"])
    df_logs.write_excel(FILE_LOG_ABSENSI)
    print(f"File '{FILE_LOG_ABSENSI}' berhasil dibuat ({df_logs.height} baris).")

    # -----------------------------------------------------------------
    # 3. GENERATE SALES ACTIVITY
    # -----------------------------------------------------------------
    sales_employees = [e for e in master_list if e["Employee_Type"] == "SALES" and e["Active"] == "Y"]
    sales_activity_rows = []

    for emp in sales_employees:
        for d in dates:
            if random.random() < PROB_SALES_ACTIVE_DAY:
                prospect = random.randint(3, 10)
                visit = random.randint(1, min(prospect, 6))
                test_drive = random.randint(0, min(visit, 3))
                
                spk = 0
                r_spk = random.random()
                if r_spk < PROB_SPK_DOUBLE:
                    spk = 2
                elif r_spk < (PROB_SPK_SINGLE + PROB_SPK_DOUBLE):
                    spk = 1

                delivery = 1 if random.random() < PROB_DELIVERY else 0
                revenue = delivery * random.choice(CAR_PRICES) if delivery > 0 else 0

                sales_activity_rows.append({
                    "No.": emp["No."],
                    "Tanggal": d.strftime("%Y-%m-%d"),
                    "Prospect": prospect,
                    "Visit": visit,
                    "Test Drive": test_drive,
                    "SPK": spk,
                    "Delivery": delivery,
                    "Revenue": revenue
                })

    df_sales_activity = pl.DataFrame(sales_activity_rows)
    df_sales_activity.write_excel(FILE_SALES_ACTIVITY)
    print(f"File '{FILE_SALES_ACTIVITY}' berhasil dibuat ({df_sales_activity.height} baris).")

    # -----------------------------------------------------------------
    # 4. GENERATE TEMPLATE CUTI / LEMBUR / EVENT (RANDOMIZED & PARAMETRIC)
    # -----------------------------------------------------------------
    active_employees = [e for e in master_list if e["Active"] == "Y"]

    # Sheet 1: Cuti_Izin
    cuti_rows = []
    for emp in active_employees:
        for d in dates:
            if random.random() < PROB_CUTI_PER_EMP_DAY:
                leave_type = random.choice(LEAVE_TYPES)
                days_count = 1 if leave_type != "Cuti Tahunan" else random.choice([1, 2, 3])
                end_d = d + timedelta(days=days_count - 1)
                
                cuti_rows.append({
                    "No.": emp["No."],
                    "Nama": emp["Name"],
                    "Tanggal Mulai": d.strftime("%Y-%m-%d"),
                    "Tanggal Selesai": end_d.strftime("%Y-%m-%d"),
                    "Jumlah Hari": days_count,
                    "Leave Type": leave_type,
                    "Status Approval": random.choice(APPROVAL_STATUSES),
                    "Keterangan": random.choice(LEAVE_REASONS[leave_type])
                })

    # Sheet 2: Lembur_Karyawan
    lembur_rows = []
    for emp in active_employees:
        for d in dates:
            if random.random() < PROB_LEMBUR_PER_EMP_DAY:
                duration_hours = random.choice([1.5, 2.0, 2.5, 3.0, 4.0])
                start_hour = 17
                start_mins = random.choice([0, 30])
                start_dt = d.replace(hour=start_hour, minute=start_mins)
                end_dt = start_dt + timedelta(hours=duration_hours)

                lembur_rows.append({
                    "No.": emp["No."],
                    "Nama": emp["Name"],
                    "Tanggal": d.strftime("%Y-%m-%d"),
                    "Jam Mulai Lembur": start_dt.strftime("%H:%M"),
                    "Jam Selesai Lembur": end_dt.strftime("%H:%M"),
                    "Durasi (Jam)": str(duration_hours).replace(".", ","),
                    "Status Approval": random.choice(APPROVAL_STATUSES),
                    "Keterangan / Alasan Lembur": random.choice(OVERTIME_REASONS)
                })

    # Sheet 3: Kegiatan_Bersama
    kegiatan_rows = []
    selected_events = random.sample(EVENT_POOL, min(NUM_KEGIATAN_BERSAMA, len(EVENT_POOL)))
    selected_dates = random.sample(dates, len(selected_events))
    
    for ev, ev_date in zip(selected_events, selected_dates):
        kegiatan_rows.append({
            "Tanggal Mulai": ev_date.strftime("%Y-%m-%d"),
            "Tanggal Selesai": ev_date.strftime("%Y-%m-%d"),
            "Nama Kegiatan / Event": ev[0],
            "Jenis Pencatatan": ev[1],
            "Durasi Lembur (Jam)": ev[2],
            "Cakupan Karyawan": ev[3],
            "Keterangan": ev[4]
        })

    df_cuti = pl.DataFrame(cuti_rows)
    df_lembur = pl.DataFrame(lembur_rows)
    df_kegiatan = pl.DataFrame(kegiatan_rows)

    # Menulis multi-sheet ke File Template Excel
    with xlsxwriter.Workbook(FILE_TEMPLATE) as workbook:
        df_cuti.write_excel(workbook=workbook, worksheet="Cuti_Izin")
        df_lembur.write_excel(workbook=workbook, worksheet="Lembur_Karyawan")
        df_kegiatan.write_excel(workbook=workbook, worksheet="Kegiatan_Bersama")

    print(f"File '{FILE_TEMPLATE}' berhasil dibuat (Sheet Cuti: {df_cuti.height} baris, Sheet Lembur: {df_lembur.height} baris, Sheet Event: {df_kegiatan.height} baris).")


if __name__ == "__main__":
    build_all_files()