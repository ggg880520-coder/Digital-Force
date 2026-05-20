SANGKULIRANG SERVICE TRACKING ANALYTICS - FINAL LIVE SA UPDATE

Perubahan pada versi ini:
1. Menu Live SA menampilkan kolom Mekanik/Nama Mekanik selain NRP.
2. Nama mekanik diambil dari field nama pada tabel sa_jobs jika ada, atau dicocokkan dari NRP pada data TimeSheet.
3. Badge status Live SA sudah diberi warna berbeda:
   - Dispatch: merah
   - Travel: orange
   - In Progress: kuning
   - Completed: hijau
4. Data Live SA yang sudah masuk status POST / POST_ACTIVITY tidak tampil lagi di tabel Live SA.
5. Data POST / POST_ACTIVITY akan digabung ke data Detail SA per mekanik pada dashboard.
6. Supabase tidak perlu diubah jika tabel sa_jobs masih menyimpan baris dengan status POST / POST_ACTIVITY dan memiliki field NRP/nama.

Catatan:
- Jika nama mekanik tidak muncul, pastikan tabel sa_jobs menyimpan salah satu kolom berikut:
  mechanic_name, name, nama, full_name, employee_name, telegram_name, atau minimal NRP yang sama dengan data TimeSheet.
- Jika data POST ingin benar-benar tersimpan permanen ke tabel sa_raw, proses tersebut sebaiknya dibuat di backend Telegram bot atau Supabase trigger, bukan dari GitHub Pages/frontend.
- Untuk update GitHub Pages, replace file index.html lama dengan index.html versi ini, lalu Commit changes.
