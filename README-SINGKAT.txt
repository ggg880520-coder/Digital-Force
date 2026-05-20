PANDUAN SINGKAT - SANGKULIRANG SERVICE TRACKING ANALYTICS

1. File utama website:
   - index.html
   - Background.png
   - Logo.png
   - Moving As One.png
   - Sangkulirang-Stronger.jpeg

2. Jalankan dashboard:
   - Upload semua file di atas ke hosting yang sama.
   - Buka index.html.

3. Sinkronisasi Google Sheet ke Supabase:
   - Install Node.js minimal versi 20.
   - Jalankan: npm install
   - Set environment variable SUPABASE_SERVICE_ROLE_KEY.
   - Jalankan: npm run sync

4. Catatan keamanan:
   - index.html hanya membaca data dari Supabase.
   - Proses insert/delete data Supabase dipindahkan ke sync.js/server.
   - Pastikan RLS Supabase aktif dan policy frontend dibatasi untuk SELECT saja.

5. Nama tabel yang dipakai:
   - sa_raw
   - timesheet_raw
   - sa_jobs
   - fast_update_log
