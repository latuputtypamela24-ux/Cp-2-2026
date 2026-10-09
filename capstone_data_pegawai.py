"""
CAPSTONE PROJECT MODULE 2
Case Study : Data Karyawan Perusahaan -> Data Pegawai Lingkup Kanwil DJPb Prov. Maluku
Fitur      : Create, Read, Update, Delete (semua menu dibuat dengan regular function)
Data       : Dictionary di dalam dictionary (key = NIP, kolom yang unik)

Struktur field & pilihan nilai (golongan, eselon, pendidikan, jabatan, unit kerja,
bidang/bagian) disesuaikan dengan file Excel "Data Pegawai Lingkup Kanwil DJPb Prov. Maluku
ISI DATA adalah DATA DUMMY FIKTIF (nama & NIP rekaan), bukan data pegawai sebenarnya.

Field tiap pegawai (11 field, 1 unik):
    nip (unik / key), nama, tgl_lahir, jk, golongan, eselon, pendidikan,
    jabatan, unit_kerja, bidang, tmt_unit
Field turunan (dihitung otomatis, tidak disimpan): usia dan masa kerja di unit.
"""

# ======================================================================
# KONSTANTA & DATA DUMMY (collection data type)
# ======================================================================
TANGGAL_ACUAN = "2026-10-08"     # tanggal acuan untuk menghitung usia & masa kerja
LEBAR = 115

DAFTAR_UNIT = ["Kanwil", "KPPN Ambon", "KPPN Masohi", "KPPN Tual", "KPPN Saumlaki"]
DAFTAR_GOLONGAN = ["II/b", "II/c", "II/d", "III/a", "III/b", "III/c", "III/d", "IV/a", "IV/b"]
DAFTAR_ESELON = ["II.A", "III.A", "IV.A", "Plks.", "Fungsional"]
DAFTAR_PENDIDIKAN = ["SMA", "D1", "D3", "D4", "S1", "S2"]
DAFTAR_JABATAN = ["Kepala Kantor Wilayah", "Kepala Kantor", "Kepala Bidang", "Kepala Bagian",
                  "Kepala Seksi", "Kepala Subbagian", "Pelaksana", "Jabatan Fungsional"]
DAFTAR_BIDANG = ["Umum", "PD", "PPA I", "PPA II", "MSKI", "PAPK", "SKKI", "Bank", "PDMS", "Vera", "-"]
DAFTAR_JK = ["L", "P"]

# Urutan kolom saat input (kolom nip diminta terpisah karena merupakan key)
KOLOM_INPUT = ["nama", "tgl_lahir", "jk", "golongan", "eselon", "pendidikan",
               "jabatan", "unit_kerja", "bidang", "tmt_unit"]

LABEL = {
    "nama": "Nama lengkap",
    "tgl_lahir": "Tanggal lahir",
    "jk": "Jenis kelamin",
    "golongan": "Golongan",
    "eselon": "Eselon",
    "pendidikan": "Pendidikan",
    "jabatan": "Jabatan",
    "unit_kerja": "Unit kerja",
    "bidang": "Bidang/Bagian",
    "tmt_unit": "TMT di unit (tgl mulai bertugas)",
}

data_pegawai = {
    "197402141998031009": {"nama": "Hendra Wijaksana", "tgl_lahir": "1974-02-14", "jk": "L",
        "golongan": "IV/b", "eselon": "II.A", "pendidikan": "S2", "jabatan": "Kepala Kantor Wilayah",
        "unit_kerja": "Kanwil", "bidang": "-", "tmt_unit": "2024-06-10"},
    "1971032002121006": {"nama": "Rizal Fahmi", "tgl_lahir": "1971-03-20", "jk": "L",
        "golongan": "IV/a", "eselon": "III.A", "pendidikan": "S2", "jabatan": "Kepala Bidang",
        "unit_kerja": "Kanwil", "bidang": "PPA I", "tmt_unit": "2023-02-01"},
    "197504272001032006": {"nama": "Dewi Anggraini", "tgl_lahir": "1975-04-27", "jk": "P",
        "golongan": "IV/a", "eselon": "III.A", "pendidikan": "S1", "jabatan": "Kepala Bagian",
        "unit_kerja": "Kanwil", "bidang": "Umum", "tmt_unit": "2022-09-15"},
    "198605092012102001": {"nama": "Selfi Pattiasina", "tgl_lahir": "1986-05-09", "jk": "P",
        "golongan": "III/c", "eselon": "IV.A", "pendidikan": "S1", "jabatan": "Kepala Subbagian",
        "unit_kerja": "Kanwil", "bidang": "Umum", "tmt_unit": "2021-07-05"},
    "199401182015011001": {"nama": "Yosef Tehuayo", "tgl_lahir": "1994-01-18", "jk": "L",
        "golongan": "II/c", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "Kanwil", "bidang": "PPA I", "tmt_unit": "2019-03-11"},
    "199509252019032002": {"nama": "Nurul Hidayah", "tgl_lahir": "1995-09-25", "jk": "P",
        "golongan": "II/c", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "Kanwil", "bidang": "MSKI", "tmt_unit": "2020-01-20"},
    "198212022005121003": {"nama": "Frans Kaihena", "tgl_lahir": "1982-12-02", "jk": "L",
        "golongan": "III/b", "eselon": "Plks.", "pendidikan": "S1", "jabatan": "Pelaksana",
        "unit_kerja": "Kanwil", "bidang": "PD", "tmt_unit": "2018-08-27"},
    "199104302013032004": {"nama": "Tiara Salampessy", "tgl_lahir": "1991-04-30", "jk": "P",
        "golongan": "III/c", "eselon": "Fungsional", "pendidikan": "S2", "jabatan": "Jabatan Fungsional",
        "unit_kerja": "Kanwil", "bidang": "SKKI", "tmt_unit": "2022-02-14"},
    "197407161997031005": {"nama": "Joni Matulessy", "tgl_lahir": "1974-07-16", "jk": "L",
        "golongan": "IV/a", "eselon": "III.A", "pendidikan": "S2", "jabatan": "Kepala Kantor",
        "unit_kerja": "KPPN Ambon", "bidang": "-", "tmt_unit": "2024-09-02"},
    "198705052006121006": {"nama": "Ahmad Rumakat", "tgl_lahir": "1987-05-05", "jk": "L",
        "golongan": "III/c", "eselon": "IV.A", "pendidikan": "S1", "jabatan": "Kepala Seksi",
        "unit_kerja": "KPPN Ambon", "bidang": "Bank", "tmt_unit": "2021-11-08"},
    "199808082018072007": {"nama": "Lusiana Pelupessy", "tgl_lahir": "1998-08-08", "jk": "P",
        "golongan": "II/c", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "KPPN Ambon", "bidang": "Vera", "tmt_unit": "2020-03-16"},
    "199410212011011008": {"nama": "Rudi Hehanussa", "tgl_lahir": "1994-10-21", "jk": "L",
        "golongan": "II/d", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "KPPN Ambon", "bidang": "PDMS", "tmt_unit": "2017-05-22"},
    "197403282000031009": {"nama": "Sapto Nugroho", "tgl_lahir": "1974-03-28", "jk": "L",
        "golongan": "IV/a", "eselon": "III.A", "pendidikan": "S1", "jabatan": "Kepala Kantor",
        "unit_kerja": "KPPN Masohi", "bidang": "-", "tmt_unit": "2023-09-18"},
    "199402142015092010": {"nama": "Wa Ode Sarifa", "tgl_lahir": "1994-02-14", "jk": "P",
        "golongan": "II/c", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "KPPN Masohi", "bidang": "Umum", "tmt_unit": "2019-10-07"},
    "198301092007101011": {"nama": "Lukas Rahanra", "tgl_lahir": "1983-01-09", "jk": "L",
        "golongan": "III/c", "eselon": "IV.A", "pendidikan": "S1", "jabatan": "Kepala Seksi",
        "unit_kerja": "KPPN Tual", "bidang": "Bank", "tmt_unit": "2023-07-13"},
    "199306172019112012": {"nama": "Hasan Renwarin", "tgl_lahir": "1993-06-17", "jk": "L",
        "golongan": "II/b", "eselon": "Plks.", "pendidikan": "D1", "jabatan": "Pelaksana",
        "unit_kerja": "KPPN Tual", "bidang": "Vera", "tmt_unit": "2021-01-25"},
    "197209032001121013": {"nama": "Petrus Lololuan", "tgl_lahir": "1972-09-03", "jk": "L",
        "golongan": "IV/a", "eselon": "III.A", "pendidikan": "S2", "jabatan": "Kepala Kantor",
        "unit_kerja": "KPPN Saumlaki", "bidang": "-", "tmt_unit": "2024-09-19"},
    "199511302016032014": {"nama": "Maya Batmomolin", "tgl_lahir": "1995-11-30", "jk": "P",
        "golongan": "II/d", "eselon": "Plks.", "pendidikan": "D3", "jabatan": "Pelaksana",
        "unit_kerja": "KPPN Saumlaki", "bidang": "PDMS", "tmt_unit": "2022-06-06"},
}


# ======================================================================
# FUNGSI TANGGAL (tanpa import, hanya string, split, dan perhitungan biasa)
# ======================================================================
def pecah_tanggal(teks):
    """'2026-10-08' -> (2026, 10, 8)."""
    bagian = teks.split("-")
    return int(bagian[0]), int(bagian[1]), int(bagian[2])


def hari_dalam_bulan(tahun, bulan):
    if bulan == 2:
        if (tahun % 4 == 0 and tahun % 100 != 0) or tahun % 400 == 0:
            return 29
        return 28
    if bulan in [4, 6, 9, 11]:
        return 30
    return 31


def valid_tanggal(teks):
    """Cek format YYYY-MM-DD dan tanggal yang benar-benar ada di kalender."""
    bagian = teks.split("-")
    if len(bagian) != 3:
        return False
    for b in bagian:
        if not b.isdigit():
            return False
    if len(bagian[0]) != 4 or len(bagian[1]) != 2 or len(bagian[2]) != 2:
        return False
    tahun, bulan, hari = pecah_tanggal(teks)
    if bulan < 1 or bulan > 12:
        return False
    return 1 <= hari <= hari_dalam_bulan(tahun, bulan)


def hitung_usia(tgl_lahir):
    """Usia (tahun penuh) pada tanggal acuan."""
    th, bl, hr = pecah_tanggal(tgl_lahir)
    ath, abl, ahr = pecah_tanggal(TANGGAL_ACUAN)
    usia = ath - th
    if (abl, ahr) < (bl, hr):
        usia -= 1
    return usia


def hitung_masa_kerja_bulan(tmt):
    """Masa kerja di unit dalam satuan bulan penuh pada tanggal acuan."""
    th, bl, hr = pecah_tanggal(tmt)
    ath, abl, ahr = pecah_tanggal(TANGGAL_ACUAN)
    total = (ath - th) * 12 + (abl - bl)
    if ahr < hr:
        total -= 1
    return total


def format_masa_kerja(tmt):
    bulan = hitung_masa_kerja_bulan(tmt)
    return f"{bulan // 12} Thn {bulan % 12} Bln"


# ======================================================================
# FUNGSI INPUT (dengan validasi)
# ======================================================================
def garis(karakter="="):
    print(karakter * LEBAR)


def input_teks(pesan):
    """Input teks tidak boleh kosong. Mengembalikan None jika user mengetik 'batal'."""
    while True:
        isi = input(pesan).strip()
        if isi.lower() == "batal":
            return None
        if isi == "":
            print("  [!] Input tidak boleh kosong. (ketik 'batal' untuk membatalkan)")
        else:
            return isi


def input_nip(pesan):
    """NIP harus 18 digit angka. Mengembalikan None jika 'batal'."""
    while True:
        isi = input(pesan).strip()
        if isi.lower() == "batal":
            return None
        if isi.isdigit() and len(isi) == 18:
            return isi
        print("  [!] NIP harus terdiri dari 18 digit angka. (ketik 'batal' untuk membatalkan)")


def input_tanggal(pesan):
    """Input tanggal YYYY-MM-DD yang valid dan tidak melewati tanggal acuan."""
    while True:
        isi = input(pesan).strip()
        if isi.lower() == "batal":
            return None
        if not valid_tanggal(isi):
            print("  [!] Format harus YYYY-MM-DD dan tanggalnya harus ada, contoh: 1990-05-17.")
        elif pecah_tanggal(isi) > pecah_tanggal(TANGGAL_ACUAN):
            print(f"  [!] Tanggal tidak boleh melewati tanggal acuan ({TANGGAL_ACUAN}).")
        else:
            return isi


def pilih_dari_daftar(judul, daftar):
    """Tampilkan pilihan bernomor, kembalikan item terpilih. None jika 'batal'."""
    print(judul)
    nomor = 1
    for item in daftar:
        print(f"  {nomor:>2}. {item}")
        nomor += 1
    while True:
        isi = input(f"Pilih [1-{len(daftar)}]: ").strip()
        if isi.lower() == "batal":
            return None
        if isi.isdigit() and 1 <= int(isi) <= len(daftar):
            return daftar[int(isi) - 1]
        print("  [!] Pilihan tidak valid.")


def input_konfirmasi(pesan):
    """True jika y, False jika n."""
    while True:
        jawab = input(pesan + " (y/n): ").strip().lower()
        if jawab == "y":
            return True
        if jawab == "n":
            return False
        print("  [!] Jawab dengan 'y' atau 'n'.")


def minta_nilai(kolom):
    """Meminta nilai baru untuk satu kolom sesuai jenisnya. None jika dibatalkan."""
    if kolom == "nama":
        nilai = input_teks("Nama lengkap: ")
        if nilai is not None:
            nilai = nilai.title()
        return nilai
    if kolom == "tgl_lahir":
        while True:
            nilai = input_tanggal("Tanggal lahir (YYYY-MM-DD): ")
            if nilai is None:
                return None
            if hitung_usia(nilai) < 18:
                print("  [!] Pegawai harus berusia minimal 18 tahun.")
            else:
                return nilai
    if kolom == "tmt_unit":
        return input_tanggal("TMT di unit / tanggal mulai bertugas (YYYY-MM-DD): ")
    if kolom == "jk":
        return pilih_dari_daftar("Jenis kelamin:", DAFTAR_JK)
    if kolom == "golongan":
        return pilih_dari_daftar("Golongan:", DAFTAR_GOLONGAN)
    if kolom == "eselon":
        return pilih_dari_daftar("Eselon:", DAFTAR_ESELON)
    if kolom == "pendidikan":
        return pilih_dari_daftar("Pendidikan:", DAFTAR_PENDIDIKAN)
    if kolom == "jabatan":
        return pilih_dari_daftar("Jabatan:", DAFTAR_JABATAN)
    if kolom == "unit_kerja":
        return pilih_dari_daftar("Unit kerja:", DAFTAR_UNIT)
    if kolom == "bidang":
        return pilih_dari_daftar("Bidang/Bagian:", DAFTAR_BIDANG)
    return None


# ======================================================================
# FUNGSI TAMPILAN
# ======================================================================
def tampilkan_tabel(data):
    """Menampilkan data (dictionary) dalam bentuk tabel."""
    if len(data) == 0:
        print("\n  Data tidak ditemukan / belum ada pegawai.")
        return
    print()
    garis()
    print(f"{'No':<4}{'NIP':<20}{'Nama':<26}{'Usia':>4} {'JK':<3}{'Gol.':<7}{'Pdd':<5}"
          f"{'Jabatan':<22}{'Unit':<15}{'Bidang':<8}")
    garis("-")
    nomor = 1
    for nip, p in data.items():
        print(f"{nomor:<4}{nip:<20}{p['nama']:<26}{hitung_usia(p['tgl_lahir']):>4} {p['jk']:<3}"
              f"{p['golongan']:<7}{p['pendidikan']:<5}{p['jabatan']:<22}{p['unit_kerja']:<15}{p['bidang']:<8}")
        nomor += 1
    garis()
    print(f"Total pegawai: {len(data)}   (usia dihitung per {TANGGAL_ACUAN})")


def tampilkan_detail(nip):
    """Menampilkan detail lengkap 1 pegawai."""
    p = data_pegawai[nip]
    jk_lengkap = "Laki-laki" if p["jk"] == "L" else "Perempuan"
    print(f"\n  NIP                 : {nip}")
    print(f"  Nama                : {p['nama']}")
    print(f"  Tanggal lahir       : {p['tgl_lahir']} (usia {hitung_usia(p['tgl_lahir'])} tahun)")
    print(f"  Jenis kelamin       : {jk_lengkap}")
    print(f"  Golongan            : {p['golongan']}")
    print(f"  Eselon              : {p['eselon']}")
    print(f"  Pendidikan          : {p['pendidikan']}")
    print(f"  Jabatan             : {p['jabatan']}")
    print(f"  Unit kerja          : {p['unit_kerja']}")
    print(f"  Bidang/Bagian       : {p['bidang']}")
    print(f"  TMT di unit         : {p['tmt_unit']} (masa kerja {format_masa_kerja(p['tmt_unit'])})")


def cari_nip(pesan="Masukkan NIP (ketik 'batal' untuk kembali): "):
    """Minta NIP sampai ditemukan di data. None jika batal."""
    while True:
        nip = input_teks(pesan)
        if nip is None:
            return None
        if nip in data_pegawai:
            return nip
        print(f"  [!] NIP '{nip}' tidak ditemukan.")


def tampilkan_hitungan(judul, kolom, daftar):
    """Mencetak jumlah pegawai per kategori beserta batang sederhana."""
    print(f"\n  {judul}:")
    for nilai in daftar:
        jumlah = 0
        for p in data_pegawai.values():
            if p[kolom] == nilai:
                jumlah += 1
        if jumlah > 0:
            print(f"    - {nilai:<22}: {jumlah:>3}  {'#' * jumlah}")


# ----- fungsi kunci untuk sorting -----
def ambil_nama(pasangan):
    return pasangan[1]["nama"].lower()


def ambil_usia(pasangan):
    return hitung_usia(pasangan[1]["tgl_lahir"])


def ambil_masa_kerja(pasangan):
    return hitung_masa_kerja_bulan(pasangan[1]["tmt_unit"])


def ambil_golongan(pasangan):
    return DAFTAR_GOLONGAN.index(pasangan[1]["golongan"])


# ======================================================================
# FITUR READ
# ======================================================================
def tampilkan_semua():
    print("\n=== DAFTAR SEMUA PEGAWAI ===")
    tampilkan_tabel(data_pegawai)


def tampilkan_berdasarkan_nip():
    print("\n=== CARI BERDASARKAN NIP ===")
    nip = cari_nip()
    if nip is not None:
        tampilkan_detail(nip)


def cari_berdasarkan_nama():
    print("\n=== CARI BERDASARKAN NAMA ===")
    kata = input_teks("Masukkan kata kunci nama (ketik 'batal' untuk kembali): ")
    if kata is None:
        return
    hasil = {}
    for nip, p in data_pegawai.items():
        if kata.lower() in p["nama"].lower():
            hasil[nip] = p
    tampilkan_tabel(hasil)


def filter_data(judul, kolom, daftar):
    """Filter umum: tampilkan pegawai yang nilai kolomnya sama dengan pilihan user."""
    print(f"\n=== FILTER BERDASARKAN {judul.upper()} ===")
    pilihan = pilih_dari_daftar(f"Pilih {judul.lower()}:", daftar)
    if pilihan is None:
        return
    hasil = {}
    for nip, p in data_pegawai.items():
        if p[kolom] == pilihan:
            hasil[nip] = p
    print(f"\nPegawai dengan {judul.lower()} '{pilihan}':")
    tampilkan_tabel(hasil)


def urutkan_data():
    print("\n=== URUTKAN DATA ===")
    print("1. Nama (A-Z)")
    print("2. Usia (termuda -> tertua)")
    print("3. Usia (tertua -> termuda)")
    print("4. Masa kerja di unit (terlama -> tersingkat)")
    print("5. Golongan (tertinggi -> terendah)")
    print("6. Kembali")
    pilihan = input("Pilih urutan [1-6]: ").strip()

    if pilihan == "1":
        terurut = sorted(data_pegawai.items(), key=ambil_nama)
    elif pilihan == "2":
        terurut = sorted(data_pegawai.items(), key=ambil_usia)
    elif pilihan == "3":
        terurut = sorted(data_pegawai.items(), key=ambil_usia, reverse=True)
    elif pilihan == "4":
        terurut = sorted(data_pegawai.items(), key=ambil_masa_kerja, reverse=True)
    elif pilihan == "5":
        terurut = sorted(data_pegawai.items(), key=ambil_golongan, reverse=True)
    elif pilihan == "6":
        return
    else:
        print("  [!] Pilihan tidak valid.")
        return

    hasil = {}
    for nip, p in terurut:
        hasil[nip] = p
    tampilkan_tabel(hasil)


def tampilkan_statistik():
    print("\n=== STATISTIK PEGAWAI ===")
    if len(data_pegawai) == 0:
        print("  Belum ada data pegawai.")
        return

    total_usia = 0
    termuda_nip = ""
    tertua_nip = ""
    usia_termuda = 999
    usia_tertua = 0
    for nip, p in data_pegawai.items():
        usia = hitung_usia(p["tgl_lahir"])
        total_usia += usia
        if usia < usia_termuda:
            usia_termuda = usia
            termuda_nip = nip
        if usia > usia_tertua:
            usia_tertua = usia
            tertua_nip = nip

    print(f"  Total pegawai   : {len(data_pegawai)}")
    print(f"  Rata-rata usia  : {total_usia / len(data_pegawai):.1f} tahun")
    print(f"  Pegawai termuda : {data_pegawai[termuda_nip]['nama']} ({usia_termuda} tahun)")
    print(f"  Pegawai tertua  : {data_pegawai[tertua_nip]['nama']} ({usia_tertua} tahun)")

    tampilkan_hitungan("Jumlah pegawai per unit kerja", "unit_kerja", DAFTAR_UNIT)
    tampilkan_hitungan("Jumlah pegawai per jenis kelamin", "jk", DAFTAR_JK)
    tampilkan_hitungan("Jumlah pegawai per pendidikan", "pendidikan", DAFTAR_PENDIDIKAN)
    tampilkan_hitungan("Jumlah pegawai per golongan", "golongan", DAFTAR_GOLONGAN)
    tampilkan_hitungan("Jumlah pegawai per jabatan", "jabatan", DAFTAR_JABATAN)


def menu_read():
    """Menu utama fitur READ."""
    while True:
        print("\n" + "=" * 48)
        print("            MENU READ (LIHAT DATA)")
        print("=" * 48)
        print(" 1. Tampilkan semua pegawai")
        print(" 2. Tampilkan pegawai berdasarkan NIP")
        print(" 3. Cari pegawai berdasarkan nama")
        print(" 4. Filter berdasarkan unit kerja")
        print(" 5. Filter berdasarkan golongan")
        print(" 6. Filter berdasarkan pendidikan")
        print(" 7. Filter berdasarkan jabatan")
        print(" 8. Filter berdasarkan bidang/bagian")
        print(" 9. Filter berdasarkan jenis kelamin")
        print("10. Urutkan data")
        print("11. Statistik pegawai")
        print("12. Kembali ke menu utama")
        pilihan = input("Pilih menu [1-12]: ").strip()

        if pilihan == "1":
            tampilkan_semua()
        elif pilihan == "2":
            tampilkan_berdasarkan_nip()
        elif pilihan == "3":
            cari_berdasarkan_nama()
        elif pilihan == "4":
            filter_data("Unit kerja", "unit_kerja", DAFTAR_UNIT)
        elif pilihan == "5":
            filter_data("Golongan", "golongan", DAFTAR_GOLONGAN)
        elif pilihan == "6":
            filter_data("Pendidikan", "pendidikan", DAFTAR_PENDIDIKAN)
        elif pilihan == "7":
            filter_data("Jabatan", "jabatan", DAFTAR_JABATAN)
        elif pilihan == "8":
            filter_data("Bidang/Bagian", "bidang", DAFTAR_BIDANG)
        elif pilihan == "9":
            filter_data("Jenis kelamin", "jk", DAFTAR_JK)
        elif pilihan == "10":
            urutkan_data()
        elif pilihan == "11":
            tampilkan_statistik()
        elif pilihan == "12":
            break
        else:
            print("  [!] Pilihan tidak valid, silakan pilih 1-12.")


# ======================================================================
# FITUR CREATE
# ======================================================================
def menu_create():
    print("\n" + "=" * 48)
    print("        MENU CREATE (TAMBAH PEGAWAI)")
    print("=" * 48)
    print("(Ketik 'batal' kapan saja untuk kembali ke menu utama)")

    # NIP = kolom unik, tidak boleh sama dengan data yang sudah ada
    while True:
        nip = input_nip("Masukkan NIP baru (18 digit): ")
        if nip is None:
            print("  Penambahan data dibatalkan.")
            return
        if nip in data_pegawai:
            print(f"  [!] NIP sudah terdaftar atas nama {data_pegawai[nip]['nama']}.")
        else:
            break

    data_baru = {}
    for kolom in KOLOM_INPUT:
        nilai = minta_nilai(kolom)
        if nilai is None:
            print("  Penambahan data dibatalkan.")
            return
        data_baru[kolom] = nilai

    # Validasi silang: TMT di unit tidak boleh sebelum pegawai berusia 18 tahun
    th_lahir = pecah_tanggal(data_baru["tgl_lahir"])[0]
    th_tmt = pecah_tanggal(data_baru["tmt_unit"])[0]
    if th_tmt < th_lahir + 18:
        print("  [!] TMT di unit tidak masuk akal (pegawai belum berusia 18 tahun). Data dibatalkan.")
        return

    print("\nData yang akan disimpan:")
    print(f"  NIP : {nip}")
    for kolom in KOLOM_INPUT:
        print(f"  {LABEL[kolom]:<34}: {data_baru[kolom]}")

    if input_konfirmasi("Simpan data pegawai ini?"):
        data_pegawai[nip] = data_baru
        print("\n  [OK] Data pegawai berhasil ditambahkan!")
        tampilkan_detail(nip)
    else:
        print("  Data tidak jadi disimpan.")


# ======================================================================
# FITUR UPDATE
# ======================================================================
def ubah_satu_kolom(nip, kolom):
    """Mengubah satu kolom milik pegawai dengan konfirmasi."""
    nilai_baru = minta_nilai(kolom)
    if nilai_baru is None:
        print("  Perubahan dibatalkan.")
        return
    nilai_lama = data_pegawai[nip][kolom]
    print(f"\n  Perubahan {LABEL[kolom]}: {nilai_lama}  ->  {nilai_baru}")
    if input_konfirmasi("Yakin ingin mengubah data ini?"):
        data_pegawai[nip][kolom] = nilai_baru
        print("  [OK] Data berhasil diperbarui!")
    else:
        print("  Perubahan dibatalkan.")


def mutasi_pegawai(nip):
    """Mengubah unit kerja, bidang, jabatan, dan TMT sekaligus (mutasi)."""
    print("\n  --- MUTASI PEGAWAI ---")
    kolom_mutasi = ["unit_kerja", "bidang", "jabatan", "tmt_unit"]
    nilai_baru = {}
    for kolom in kolom_mutasi:
        nilai = minta_nilai(kolom)
        if nilai is None:
            print("  Mutasi dibatalkan.")
            return
        nilai_baru[kolom] = nilai

    print("\n  Ringkasan mutasi:")
    for kolom in kolom_mutasi:
        print(f"    {LABEL[kolom]:<34}: {data_pegawai[nip][kolom]}  ->  {nilai_baru[kolom]}")
    if input_konfirmasi("Proses mutasi ini?"):
        for kolom in kolom_mutasi:
            data_pegawai[nip][kolom] = nilai_baru[kolom]
        print("  [OK] Mutasi berhasil diproses!")
    else:
        print("  Mutasi dibatalkan.")


def menu_update():
    print("\n" + "=" * 48)
    print("          MENU UPDATE (UBAH DATA)")
    print("=" * 48)
    tampilkan_tabel(data_pegawai)
    if len(data_pegawai) == 0:
        return

    nip = cari_nip("Masukkan NIP pegawai yang akan diubah (ketik 'batal' untuk kembali): ")
    if nip is None:
        return

    while True:
        print("\nData saat ini:")
        tampilkan_detail(nip)
        print("\nKolom yang dapat diubah (NIP tidak dapat diubah karena menjadi kunci data):")
        nomor = 1
        for kolom in KOLOM_INPUT:
            print(f"  {nomor:>2}. {LABEL[kolom]}")
            nomor += 1
        print(f"  {nomor:>2}. Mutasi (unit, bidang, jabatan, TMT sekaligus)")
        print(f"  {nomor + 1:>2}. Selesai")
        pilihan = input(f"Pilih [1-{nomor + 1}]: ").strip()

        if not pilihan.isdigit():
            print("  [!] Masukkan angka sesuai pilihan.")
            continue
        pilihan = int(pilihan)

        if pilihan == nomor + 1:
            print("  Selesai mengubah data.")
            break
        elif pilihan == nomor:
            mutasi_pegawai(nip)
        elif 1 <= pilihan <= len(KOLOM_INPUT):
            ubah_satu_kolom(nip, KOLOM_INPUT[pilihan - 1])
        else:
            print(f"  [!] Pilihan tidak valid, silakan pilih 1-{nomor + 1}.")


# ======================================================================
# FITUR DELETE
# ======================================================================
def hapus_satu_pegawai():
    tampilkan_tabel(data_pegawai)
    if len(data_pegawai) == 0:
        return
    nip = cari_nip("Masukkan NIP pegawai yang akan dihapus (ketik 'batal' untuk kembali): ")
    if nip is None:
        return

    print("\nPegawai yang akan dihapus:")
    tampilkan_detail(nip)
    if data_pegawai[nip]["jabatan"].startswith("Kepala"):
        print(f"\n  [PERINGATAN] Pegawai ini menjabat sebagai {data_pegawai[nip]['jabatan']}!")

    if input_konfirmasi("Yakin ingin menghapus data pegawai ini?"):
        nama = data_pegawai[nip]["nama"]
        del data_pegawai[nip]
        print(f"\n  [OK] Data '{nama}' berhasil dihapus.")
        tampilkan_tabel(data_pegawai)
    else:
        print("  Penghapusan dibatalkan.")


def hapus_per_unit():
    unit = pilih_dari_daftar("Pilih unit kerja yang SELURUH pegawainya akan dihapus:", DAFTAR_UNIT)
    if unit is None:
        return

    # Kumpulkan dulu NIP yang akan dihapus (dictionary tidak boleh diubah saat di-loop)
    daftar_hapus = []
    target = {}
    for nip, p in data_pegawai.items():
        if p["unit_kerja"] == unit:
            daftar_hapus.append(nip)
            target[nip] = p

    if len(daftar_hapus) == 0:
        print(f"  Tidak ada pegawai di unit {unit}.")
        return

    print(f"\nPegawai berikut akan DIHAPUS dari unit {unit}:")
    tampilkan_tabel(target)
    konfirmasi = input(f"Ketik HAPUS (huruf besar) untuk menghapus {len(daftar_hapus)} data, "
                       "atau tekan Enter untuk membatalkan: ").strip()
    if konfirmasi == "HAPUS":
        for nip in daftar_hapus:
            del data_pegawai[nip]
        print(f"\n  [OK] {len(daftar_hapus)} data pegawai unit {unit} berhasil dihapus.")
    else:
        print("  Penghapusan dibatalkan.")


def menu_delete():
    while True:
        print("\n" + "=" * 48)
        print("          MENU DELETE (HAPUS DATA)")
        print("=" * 48)
        print("1. Hapus satu pegawai berdasarkan NIP")
        print("2. Hapus semua pegawai pada satu unit kerja")
        print("3. Kembali ke menu utama")
        pilihan = input("Pilih menu [1-3]: ").strip()

        if pilihan == "1":
            hapus_satu_pegawai()
        elif pilihan == "2":
            hapus_per_unit()
        elif pilihan == "3":
            break
        else:
            print("  [!] Pilihan tidak valid, silakan pilih 1-3.")


# ======================================================================
# MENU UTAMA
# ======================================================================
def menu_utama():
    while True:
        print("\n" + "=" * 48)
        print("   SISTEM DATA PEGAWAI KANWIL DJPb MALUKU")
        print("          (Data Dummy / Fiktif)")
        print("=" * 48)
        print("1. Lihat Data Pegawai    (Read)")
        print("2. Tambah Pegawai        (Create)")
        print("3. Ubah Data Pegawai     (Update)")
        print("4. Hapus Data Pegawai    (Delete)")
        print("5. Keluar")
        pilihan = input("Pilih menu [1-5]: ").strip()

        if pilihan == "1":
            menu_read()
        elif pilihan == "2":
            menu_create()
        elif pilihan == "3":
            menu_update()
        elif pilihan == "4":
            menu_delete()
        elif pilihan == "5":
            print("\nTerima kasih telah menggunakan program ini. Sampai jumpa!")
            break
        else:
            print("  [!] Pilihan tidak valid, silakan pilih 1-5.")


menu_utama()
