Aplikasi Web Visit Magang Agronomi
Offline-first + Sinkronisasi Google Sheet

Update versi modifikasi:
1. Spreadsheet ID sudah di-hardcode di backend GAS.
2. Endpoint GAS Web App sudah di-hardcode di frontend agar user tidak perlu input manual lagi.
3. Saat aplikasi dibuka akan muncul notifikasi jika masih ada antrian sinkron atau data gagal sinkron.
4. Ditambahkan indikator konflik data antar perangkat.
5. Ditambahkan halaman "Gagal Sinkron" untuk retry sinkron ulang.
6. Ditambahkan progress bar visual saat proses sinkronisasi push/pull/full sync.

Spreadsheet online yang dipakai (hardcode backend GAS):
1czxW4yH9OZv0ag9GnO4E-EqY8aNuU_nD6KxMmhs2zZg

Endpoint GAS Web App yang sudah dipasang di app.js:
https://script.google.com/macros/s/AKfycbxTOPuN4FJTedfU_hpEbxVcB1XH6cAUVUABr5wAc0WI0O5AnvXukCETb8pwdB19y8p-/exec

File tambahan:
- gas/Code.gs  -> script backend Google Apps Script yang harus Anda paste/update di project GAS lalu deploy ulang sebagai Web App.

Cara setup/update backend Google Apps Script:
1. Buka project GAS Anda.
2. Hapus kode lama, lalu copy seluruh isi file gas/Code.gs.
3. Simpan project.
4. Deploy > Manage deployments.
5. Edit deployment Web App yang sudah ada atau buat deployment baru.
6. Execute as: Me.
7. Who has access: Anyone.
8. Deploy.
9. Karena endpoint sudah di-hardcode di frontend, pastikan URL deployment yang aktif sama dengan URL di atas. Jika Anda membuat URL deployment baru yang berbeda, ganti konstanta GAS_WEBAPP_URL pada app.js.

Cara pakai aplikasi:
1. Ekstrak file ZIP.
2. Buka index.html di browser modern (Chrome / Edge / Firefox).
3. Isi nama perangkat pada dashboard sinkronisasi.
4. Simpan pengaturan sinkronisasi.
5. Jalankan Tes Koneksi.
6. Gunakan Kirim Antrian ke Online / Tarik Data dari Online / Sinkron Penuh sesuai kebutuhan.
7. Jika ada kegagalan, buka halaman Gagal Sinkron untuk retry.
8. Jika ada benturan data antar perangkat, buka halaman Konflik Data.

Catatan sinkronisasi:
- ID laporan menjadi kunci utama overwrite di Google Sheet.
- NIK peserta menjadi kunci utama overwrite di Google Sheet.
- Konflik ditandai ketika versi online dari perangkat lain lebih baru daripada data lokal yang masih pending.
- Pull dari online akan memasukkan data yang lebih baru ke local storage selama tidak bentrok dengan perubahan lokal pending.
- Jika terjadi benturan, data akan ditandai sebagai konflik dan masuk juga ke daftar gagal sinkron agar mudah ditindaklanjuti.
- Aplikasi tetap dapat dipakai offline menggunakan IndexedDB.

Fitur utama:
- Offline-first.
- Master peserta dapat diunggah dari file .xlsx.
- Autosuggest peserta berdasarkan nama / NIK / unit.
- Satu laporan visit bisa memuat lebih dari 1 mentee.
- Laporan bisa diedit dan disimpan overwrite.
- Export laporan ke Excel .xlsx.
- Preview & kirim ringkasan visit ke WhatsApp.
- Dark mode.
- Sinkronisasi online ke Google Sheet.
- Progress bar sinkronisasi.
- Indikator konflik antar perangkat.
- Retry data gagal sinkron.


Perbaikan Chrome Mobile:
- Request Apps Script sekarang memprioritaskan POST via iframe + postMessage.
- JSONP hanya dipakai sebagai cadangan.
- Backend doPost diperkuat agar mengirim respon ke parent/top/opener dan mengizinkan iframe (ALLOWALL).
