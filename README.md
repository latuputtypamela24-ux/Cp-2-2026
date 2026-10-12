# Cp-2-2026
Presentasi ini memaparkan aplikasi CRUD berbasis terminal yang dibuat dengan Python untuk mengelola data pegawai. Case study-nya disesuaikan dengan Excel data pegawai Kanwil DJPb Maluku dan 4 KPPN, dengan 18 data dummy fiktif dari 5 unit kerja dan 11 field. Data disimpan dalam dictionary of dictionary dengan NIP sebagai key unik, sementara usia dan masa kerja dihitung otomatis dari tanggal.

Program berjalan dari satu menu utama dengan empat fitur, semuanya berupa regular function:

Create: tambah pegawai dengan validasi NIP (18 digit, tidak duplikat), tanggal, dan usia minimal 18, lalu konfirmasi sebelum simpan.
Read: lihat, cari, filter, urutkan, dan statistik data.
Update: ubah kolom satu per satu atau lewat fitur mutasi, dengan NIP yang tidak bisa diubah.
Delete: hapus per NIP dengan konfirmasi, atau per unit dengan mengetik “HAPUS”.

Nilai tambahnya adalah validasi input, statistik otomatis, fitur mutasi, dan double confirm saat hapus. Aplikasi ini siap dikembangkan dengan penyimpanan ke file atau database.


