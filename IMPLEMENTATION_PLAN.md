# Implementation Plan: Enhancing HR Dashboard Attendance, Overtime, Multi-Schedule & Payroll Rules

Rencana implementasi untuk mengakomodasi 4 kebutuhan baru dari HR:
1. **Multi-sheet Upload Data Cuti, Lembur & Kegiatan Bersama (Global)**
2. **Opsi Potongan Keterlambatan: Hitung Per Menit vs Per 1x Telat (Kejadian)**
3. **Employee Master dengan Range Jam Masuk & Jam Pulang per Divisi/Karyawan**
4. **Kolom Khusus Jumlah Tidak Absen Pulang & Potongan Gaji Tim Sales per 1x Tidak Absen Pulang**

---

## User Review Required

> [!IMPORTANT]
> - **Template Cuti/Lembur/Kegiatan**: Template Excel baru akan menyertakan 3 sheet (`Cuti_Izin`, `Lembur_Karyawan`, `Kegiatan_Bersama`) dengan fleksibilitas pembacaan variasi nama sheet (`Cuti/Izin`, `Lembur`, `Event`, dll). Format lama (single sheet cuti / CSV) tetap didukung untuk *backward compatibility*.
> - **Lembur Bersama (Event Global)**: Jika tercatat event "Lembur Bersama" pada suatu tanggal (misalnya pameran launching di mall), seluruh karyawan (atau sesuai cakupan event) **tidak akan dianggap mangkir** walaupun tidak ada scan fingerprint di kantor, dan jam kerja/lembur dicatat sesuai durasi kegiatan.
> - **Potongan Telat**: Sekarang tersedia 2 pilihan mode: `per_minute` (Rp per menit telat) atau `per_occurrence` (Rp per 1x telat).
> - **Potongan Tidak Absen Pulang**: Khusus untuk karyawan bertipe **SALES** (`Employee_Type == "SALES"`), setiap kejadian tidak absen pulang (`Lupa Absen Pulang`) dapat dikenakan potongan gaji nominal per 1x kejadian sesuai konfigurasi HR. Karyawan tipe **OFFICE** tidak dikenakan potongan anomali sales ini.

---

## Proposed Changes

### 1. Multi-Sheet Leave, Overtime & Global Events
#### [MODIFY] [leave.py](file:///f:/hr-dashboard-refactor-1/hrdash/leave.py)
- Tingkatkan fungsi `load_leave_workbook` agar mendukung pencocokan nama sheet yang fleksibel:
  1. Sheet Cuti/Izin: `Cuti_Izin`, `Cuti/Izin`, `Cuti`, `Data Cuti`, `Izin`
  2. Sheet Lembur Karyawan: `Lembur_Karyawan`, `Lembur Karyawan`, `Lembur`
  3. Sheet Kegiatan/Event: `Kegiatan_Bersama`, `Kegiatan/Event/Lembur Bersama/Cuti Bersama`, `Kegiatan Bersama`, `Event`, `Kegiatan`
- Tambahkan parsing kolom durasi lembur pada Sheet Kegiatan Bersama (`Durasi Lembur (Jam)`).
- Di `apply_leave_kegiatan`:
  - **Lembur Bersama**: Jika ada event lembur bersama pada tanggal bersangkutan, karyawan tanpa scan tidak dianggap mangkir (`Absent_Flag = 0`, `Status_Masuk = "Lembur Bersama"`, `Status_Pulang = "Lembur Bersama"`, `Group_Event = Nama Event`), serta diakumulasikan jam lemburnya.
  - **Libur Bersama**: Ditandai libur bersama, `Absent_Flag = 0`.
  - **Cuti/Izin**: Perkaryawan berdasarkan NIK/No.
  - **Lembur Karyawan**: Perkaryawan berdasarkan NIK/No.
- Perbarui `build_leave_template()` agar menghasilkan workbook Excel lengkap dengan 3 sheet berdesain profesional, validasi data dropdown, dan sheet "Petunjuk".

---

### 2. Employee Master Range Jam Masuk & Jam Pulang
#### [MODIFY] [employee_master.py](file:///f:/hr-dashboard-refactor-1/hrdash/employee_master.py)
- Tambahkan kolom `Jam_Masuk` dan `Jam_Pulang` ke `EMPLOYEE_MASTER_COLUMNS`.
- Di `load_employee_master`:
  - Dukung input terpisah (`Jam_Masuk`, `Jam_Pulang` / `Jam Masuk`, `Jam Pulang`) maupun format rentang (`Jam_Kerja` e.g. `"08:30 - 17:30"`).
  - Parse ke format waktu yang valid (`HH:MM`).
  - Jika kosong, fallback ke `pd.NA` (akan menggunakan jam kerja default dashboard).
- Di `merge_master_into_log`:
  - Meneruskan `Jam_Masuk` dan `Jam_Pulang` ke dataframe kehadiran.
- Di `build_employee_master_template()`:
  - Sertakan kolom `Jam_Masuk` dan `Jam_Pulang` serta contoh jadwal divisi Sales vs Office pada sheet data dan sheet Petunjuk.

---

### 3. Attendance Calculation Engine & Analytics
#### [MODIFY] [engine.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/engine.py)
- Di `build_employee_directory` & `build_daily_attendance`:
  - Bawa `Jam_Masuk` dan `Jam_Pulang` spesifik karyawan dari Employee Master.
  - Saat evaluasi baris kehadiran harian:
    - Jika karyawan memiliki `Jam_Masuk` master, gunakan nilai tersebut (ditambah `grace_minutes`) untuk mengecek keterlambatan.
    - Jika karyawan memiliki `Jam_Pulang` master, gunakan nilai tersebut untuk mengecek pulang cepat.
    - Jika kosong, gunakan default global `AttendanceConfig`.
- Buat flag khusus:
  - `Tidak_Absen_Pulang_Flag = (Status_Pulang == "Lupa Absen Pulang").astype(int)`
- Integrasikan hasil pemrosesan `hrdash.leave` (Cuti, Lembur, dan Kegiatan Lembur/Libur Bersama).

#### [MODIFY] [analytics.py](file:///f:/hr-dashboard-refactor-1/hrdash/attendance/analytics.py)
- Di `calculate_employee_summary`:
  - Tambahkan agregasi kolom khusus:
    - `Tidak_Absen_Pulang = ("Tidak_Absen_Pulang_Flag", "sum")`
    - `Total_Jam_Lembur = ("Overtime_Hours", "sum")` (jika ada data lembur)
- Tambahkan indikator insight HR jika ada karyawan Sales yang memiliki frekuensi tinggi tidak absen pulang.

---

### 4. Konfigurasi & Perhitungan Payroll
#### [MODIFY] [config.py](file:///f:/hr-dashboard-refactor-1/hrdash/config.py)
- Perluas `PayrollConfig`:
  - `late_deduction_mode: str = "per_minute"` (pilihan: `"per_minute"` atau `"per_occurrence"`)
  - `late_deduction_per_minute: Optional[float] = None`
  - `late_deduction_per_occurrence: Optional[float] = None` (Rp per 1x terlambat)
  - `sales_no_clock_out_deduction: Optional[float] = None` (Rp per 1x tidak absen pulang khusus Sales)
- Perbarui `config_to_dataframe` dan `config_from_dataframe` untuk mendukung parameter baru ini secara dua arah (round-trip).

#### [MODIFY] [payroll.py](file:///f:/hr-dashboard-refactor-1/hrdash/payroll.py)
- Perbarui `PAYROLL_COLUMNS` dengan menambahkan:
  - `Jumlah_Tidak_Absen_Pulang`
  - `Potongan_Lupa_Pulang_Sales`
- Logika perhitungan:
  - **Potongan Telat**:
    - Jika `mode == "per_occurrence"`: `Jumlah_Terlambat * late_deduction_per_occurrence`
    - Jika `mode == "per_minute"`: `Total_Menit_Telat * late_deduction_per_minute`
  - **Potongan Tidak Absen Pulang Sales**:
    - Jika `Employee_Type == "SALES"` dan tarif diisi: `Jumlah_Tidak_Absen_Pulang * sales_no_clock_out_deduction`
    - Untuk selain SALES: tidak dikenakan potongan ini (Rp 0).
  - Masukkan komponen baru ini ke dalam `Total_Potongan` dan kurangkan dari `Estimasi_Take_Home`.

---

### 5. Streamlit UI & Excel Report
#### [MODIFY] [app.py](file:///f:/hr-dashboard-refactor-1/hrdash/app.py)
- **Sidebar 1 (Data)**:
  - File uploader: `"Upload daftar cuti/lembur/kegiatan bersama (opsional)"`
  - Tombol download: `📥 Template Cuti, Lembur & Kegiatan Bersama` (menggunakan `build_leave_template()`).
  - Tombol download template Employee Master (yang sudah ada `Jam_Masuk` & `Jam_Pulang`).
- **Sidebar 4 (Tarif Potongan Payroll)**:
  - Opsi potongan telat: Radio / Selectbox `Metode Potongan Telat`:
    - `Per Menit (Rp / menit)` -> input nominal per menit.
    - `Per 1x Telat (Rp / kejadian)` -> input nominal per 1x telat.
  - Opsi potongan tidak absen pulang Sales:
    - Checkbox & input nominal: `Rp / 1x tidak absen pulang tim Sales`.
- **Tab Tampilan**:
  - Employee Tab & Payroll Tab: tampilkan kolom `Tidak_Absen_Pulang`, `Potongan_Lupa_Pulang_Sales`, dan lembur jika ada.
  - Anomaly Tab: filter dan penandaan jelas untuk anomali Lupa Absen Pulang.
- Hubungkan `load_leave_workbook` ke alur kalkulasi.

#### [MODIFY] [excel_report.py](file:///f:/hr-dashboard-refactor-1/hrdash/reporting/excel_report.py)
- Perbarui definisi tabel sheet PAYROLL dan EMPLOYEE_SUMMARY agar mencakup kolom baru `Jumlah_Tidak_Absen_Pulang`, `Potongan_Lupa_Pulang_Sales`, dan jadwal jam kerja di EMPLOYEE_MASTER.

---

## Verification Plan

### Automated Verification
- Menjalankan skrip validasi python mandiri untuk memverifikasi:
  1. Parsing file Excel 3-sheet (Cuti/Izin, Lembur Karyawan, Kegiatan Bersama).
  2. Verifikasi Lembur Bersama: memastikan karyawan tanpa scan fingerprint pada hari event launching mall tidak tercatat sebagai mangkir.
  3. Verifikasi jam masuk/pulang kustom per karyawan di Employee Master.
  4. Verifikasi kedua opsi potongan telat (per menit vs per kejadian).
  5. Verifikasi potongan gaji khusus tim sales untuk 1x tidak absen pulang, dan memastikan non-sales tidak terpotong.
  6. Verifikasi round-trip AppConfig (export & reload dataframe).

### Manual UI Verification
- Validasi visual di UI dashboard:
  - Sidebar inputs (upload 3-sheet, template downloads, radio pilihan potongan telat, input tarif sales).
  - Tab Employee, Anomaly, dan Payroll menampilkan data dan kalkulasi dengan benar.
