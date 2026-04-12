STYLES = """
<style>
/* ── Page title (st.title → h1) ── */
h1 {
    color: #1C549D !important;
    font-weight: 700;
}

/* ── Metric value & container ── */
[data-testid="stMetricValue"] {
    font-size: 2rem;
    font-weight: bold;
    color: #2E9D32;
}
[data-testid="stMetric"] {
    border-left: 4px solid #2E9D32;
    padding-left: 8px;
}

/* ── Success alert: green left border ── */
[data-testid="stNotificationContentSuccess"],
div[data-testid="stAlert"] > div[role="alert"][data-baseweb="notification"][kind="positive"] {
    border-left: 4px solid #2E9D32 !important;
}

/* ── Primary & default action buttons ── */
[data-testid="stBaseButton-primary"] > button,
[data-testid="stBaseButton-secondary"] > button,
.stButton > button {
    background-color: #2E9D32 !important;
    color: #FFFFFF !important;
    border: 1px solid #236E26 !important;
    font-weight: 600;
}
[data-testid="stBaseButton-primary"] > button:hover,
[data-testid="stBaseButton-secondary"] > button:hover,
.stButton > button:hover {
    background-color: #236E26 !important;
    border-color: #1A5220 !important;
}

/* ── Sidebar title ── */
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] .stMarkdown h1 {
    color: #1C549D !important;
}

/* ── Footer ── */
.footer {
    position: fixed;
    left: 0; bottom: 0;
    width: 100%;
    background-color: #f8f9fa;
    color: #6c757d;
    text-align: center;
    padding: 10px;
    font-size: 13px;
    border-top: 2px solid #2E9D32;
    z-index: 1000;
}
.stApp { margin-bottom: 80px; }

/* ── Checkbox ── */
.stCheckbox {
    background-color: #e2e3e5;
    padding: 10px;
    border-radius: 5px;
    border: 1px solid #ced4da;
}
.stCheckbox label p,
.stCheckbox label span {
    color: #000000 !important;
    font-weight: bold;
}
</style>
<div class="footer">
    Dikembangkan secara internal untuk mendukung kerja data <strong>Antasena</strong>
</div>
"""

COLOR_MAP = {
    "UNIK": "#28a745",
    "GANDA": "#dc3545",
    "BUKAN ANGKA": "#ffc107",
    "TIDAK 16 DIGIT": "#fd7e14",
    "TERKONVERSI (000)": "#17a2b8",
    "KOSONG": "#6c757d",
    "SUDAH SALUR 2026": "#6f42c1",
}

COLOR_MAP_KATEGORI = {
    "ANAK": "#4fc3f7",
    "DEWASA": "#66bb6a",
    "LANSIA": "#ffa726",
    "TIDAK VALID": "#ef5350",
}

TEXT_COLUMNS_KEYWORDS = [
    "NIK", "KK", "NO KK", "NO. HP", "NOMOR HP",
    "TANGGAL", "TGL", "SK", "NOMOR SK",
    "BAST", "NOMOR BAST", "NOMOR", "NO ",
]

SPLIT_HELP_TEXT = {
    "upload": "Upload file Excel/CSV. File akan diproses di memory server dan tidak disimpan secara permanen.",
    "header_row": "Baris yang berisi nama kolom header. Default: 1 (baris pertama). Ubah jika header tidak di baris pertama.",
    "kolom_split": "Pilih kolom yang nilainya akan digunakan untuk memecah file. Setiap unique value di kolom ini akan menjadi 1 file output terpisah.",
    "kolom_split_bertingkat": "Pilih satu atau beberapa kolom. Urutan pilihan menjadi level split, misalnya Provinsi > Kabupaten > Kecamatan.",
    "preview_stats": "Menampilkan statistik dasar: jumlah baris data, jumlah unique value di kolom split, dan estimasi jumlah file yang akan dihasilkan.",
    "preview_stats_bertingkat": "Menampilkan statistik dasar split bertingkat: total baris, jumlah level, estimasi output, dan jumlah sel split yang kosong.",
    "auto_clean": "Gabungkan nilai yang serupa menjadi satu. Contoh: 'kab. boyolali', 'kabupaten boyolali', 'KABUPATEN BOYOLALI' akan digabung menjadi satu nilai. Menggunakan fuzzy matching dengan threshold 80%.",
    "text_format": "Aktifkan untuk menjaga format teks di kolom tertentu agar tidak terkonversi menjadi angka oleh Excel. Contoh: NIK '0012345678901234' tidak berubah jadi '1.2345E+15'.",
    "select_columns": "Pilih kolom yang perlu diformat sebagai teks. Sistem akan otomatis mendeteksi kolom yang mengandung keyword seperti: NIK, KK, NO. HP, TANGGAL, SK, BAST.",
    "select_all": "Centang untuk memilih semua kolom. Batalkan centang untuk membatalkan semua pilihan.",
    "process": "Memulai proses pemisahan file. File output akan dibundle dalam 1 file ZIP untuk download.",
    "cancel": "Membatalkan proses. Semua file yang sudah dibuat akan dihapus.",
    "merge_upload": "Upload workbook yang berisi tabel master. Sheet sumber bisa berasal dari sheet lain di workbook ini atau dari workbook sumber tambahan.",
}
