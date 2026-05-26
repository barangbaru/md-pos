Update POS System — Master Terapis, Input Point, & Konsolidasi Superadmin
==========================================================================

FITUR BARU:
-----------
1. **Master Terapis** — Tab baru untuk mengelola data terapis (ID, nama, HP, nilai default point, status aktif/non-aktif).
2. **Input Point Terapis** — Kasir dapat mencatat hasil layanan terapis: nama layanan, qty, point per unit, total tarif (opsional).
3. **Kasbon Terapis** — Pengeluaran dengan type 'Therapist' untuk mencatat kasbon yang diberikan ke terapis.
4. **Konsolidasi Superadmin** — Tab khusus superadmin untuk settlement: hitung total point × % bagi hasil, kurangi kasbon, hasilkan net payout per terapis.
5. **Filter Range Tanggal** — Dashboard & konsolidasi punya filter range tanggal (hari ini, 7 hari, 30 hari, bulan ini, custom).

SHEET BARU:
-----------
- THERAPISTS: master terapis (Therapist_ID, Therapist_Name, Phone, Default_Point_Value, Active, Notes, Created_At, Updated_At)
- THERAPIST_POINTS: log point (Point_ID, Date, Time, Therapist_ID, Therapist_Name, Service_Name, Qty, Point_Per_Unit, Total_Point, Amount, Cashier_Name, Notes, Created_At)

CONFIG BARU:
------------
- THERAPIST_PASSCODE (default: 111111) — untuk buka tab Master Terapis
- SUPERADMIN_PASSCODE (default: 777777) — untuk buka tab Konsolidasi
- THERAPIST_SHARE_PERCENT (default: 50) — persentase bagi hasil terapis dari total point

ROLE AKSES:
-----------
- **Kasir** (CASHIER_PASSCODE): login, POS, stock, pengeluaran
- **Terapis Manager** (THERAPIST_PASSCODE): master terapis + input point
- **Dashboard** (DASHBOARD_PASSCODE): lihat dashboard penjualan
- **Superadmin** (SUPERADMIN_PASSCODE): konsolidasi settlement terapis
- **Config** (CONFIG_PASSCODE): ubah setting

CARA PAKAI:
-----------
1. **Setup awal**: Jalankan `setupPOSDatabase()` (atau `upgradePOSDatabaseSchema()` jika sudah punya data lama).
2. **Tambah terapis**: Buka tab Terapis (passcode 111111), isi form master terapis, klik Simpan.
3. **Input point**: Di tab Terapis, bagian Input Point, pilih terapis, isi layanan & point, klik Simpan Point.
4. **Kasbon terapis**: Bisa lewat tab Pengeluaran biasa (pilih kategori kasbon + nama terapis), atau nanti bisa dibuat UI khusus di tab Terapis.
5. **Lihat settlement**: Buka tab Konsolidasi (passcode 777777), pilih range tanggal, lihat net payout per terapis (share - kasbon).

FORMULA KONSOLIDASI:
--------------------
Per terapis:
- total_point_nominal = SUM(Total_Point) — anggap Total_Point sudah ekuivalen rupiah
- share = total_point_nominal × (THERAPIST_SHARE_PERCENT / 100)
- kasbon = SUM(EXPENSES.Amount) WHERE Expense_Type='Therapist' AND Personal_Cashier=therapistName
- net_payout = share - kasbon

FILE YANG BERUBAH:
------------------
- Setup.gs: tambah setupTherapistsSheet_, setupTherapistPointsSheet_, config baru, upgradePOSDatabaseSchema extended
- Utils.gs: tambah POS_SHEET.THERAPISTS & THERAPIST_POINTS
- Backend.gs: validateProtectedTabPasscode (therapist, superadmin), getConfig/updateConfig (share %), addExpense (targetType='therapist'), API baru (saveTherapist, getTherapists, saveTherapistPoint, getTherapistPoints, getTherapistConsolidation, dll)
- Index.html: tab Terapis (master form, tabel, input point form, log point), tab Konsolidasi (filter range, summary cards, tabel settlement, chart harian), Config form (passcode baru, share %)
- Script.html: handler semua tab baru, switchTab extended, protectedModal extended, lockProtectedArea extended, helper rupiah input
- Style.html: status badges (aktif/non-aktif)

CARA PASANG:
------------
1. Backup Apps Script lama.
2. Replace semua file di project Apps Script dengan file dari ZIP ini.
3. Jalankan `upgradePOSDatabaseSchema()` dari Apps Script Editor (Tools > Script editor > pilih function > Run) jika sudah punya data lama. Kalau baru, jalankan `setupPOSDatabase()`.
4. Deploy ulang Web App (Deploy > Manage deployments > Edit > Version: New version > Deploy).
5. Buka web app, test role akses:
   - Login kasir (passcode kasir lama)
   - Buka tab Terapis (passcode: 111111)
   - Buka tab Konsolidasi (passcode: 777777)
   - Buka Config, ubah passcode kalau mau

CATATAN:
--------
- Kasbon terapis tersimpan di sheet EXPENSES dengan Expense_Type='Therapist' dan Personal_Cashier=nama terapis.
- Konsolidasi auto-link kasbon ke terapis via nama (case-insensitive). Pastikan nama terapis di master sama dengan nama di kasbon.
- Bagi hasil % bisa diubah di Config atau langsung di tab Konsolidasi saat filter.
- Range maksimal 365 hari untuk menjaga performa.
