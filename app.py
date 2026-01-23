import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Cek Validitas Internal Antasena", layout="wide")

# --- 2. STYLE & FOOTER (CSS) ---
st.markdown("""
<style>
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #f8f9fa;
    color: #6c757d;
    text-align: center;
    padding: 10px;
    font-size: 13px;
    border-top: 1px solid #dee2e6;
    z-index: 1000;
}
.stApp {
    margin-bottom: 80px; /* Jarak aman agar konten tidak tertutup footer */
}
</style>
<div class="footer">
    Dibuat oleh <strong>RBKA</strong> untuk digunakan internal <strong>Antasena</strong>
</div>
""", unsafe_allow_html=True)

# --- 3. FUNGSI LOGGING ---
def catat_log(nama_file, nama_sheet, rincian_per_kolom):
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary_text = ""
    for col, stats in rincian_per_kolom.items():
        stat_str = ", ".join([f"{k}:{v}" for k, v in stats.items()])
        summary_text += f"[{col}: {stat_str}] "
    pesan = f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | DETAIL: {summary_text}\n"
    with open("activity_log.txt", "a") as f:
        f.write(pesan)

# --- 4. LOGIKA UTAMA ---
st.title("🛡️ Validasi Data - Internal Antasena")
st.info("Fitur Lengkap: Atur Posisi Header, Multi-Kolom, Multi-Sheet, & Auto-Format Text.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- LANGKAH 1: BACA STRUKTUR FILE ---
        xls = pd.ExcelFile(uploaded_file)
        daftar_sheet = xls.sheet_names
        
        # --- LANGKAH 2: PILIH SHEET ---
        st.subheader("1. Konfigurasi File")
        col_sheet, col_header_row = st.columns([2, 1])
        
        with col_sheet:
            selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
        
        # --- FITUR BARU: PREVIEW RAW DATA ---
        # Baca 10 baris pertama TANPA header untuk membantu user melihat posisi header
        df_preview_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, nrows=10)
        df_preview_raw = df_preview_raw.fillna('') # Biar rapi di tampilan
        
        with st.expander("🔍 Klik untuk melihat Preview Data Mentah (Cek posisi Header disini)", expanded=True):
            st.caption("Lihat tabel di bawah ini. Di baris nomor berapakah nama kolom (Header) Anda berada?")
            # Trik agar index mulai dari 1 di tampilan, bukan 0
            df_preview_raw.index += 1 
            st.dataframe(df_preview_raw, use_container_width=True)

        # --- FITUR BARU: PILIH URUTAN ROW ---
        with col_header_row:
            header_row_input = st.number_input(
                "Header Table ada di baris ke:", 
                min_value=1, 
                value=1, 
                help="Jika judul kolom ada di baris 3, isi angka 3."
            )

        # --- LANGKAH 3: BACA DATA SESUNGGUHNYA ---
        # header=header_row_input - 1 (karena Python mulai hitung dari 0, manusia dari 1)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row_input - 1)
        
        # Hapus baris yang kosong semua (opsional, untuk kebersihan data)
        df.dropna(how='all', inplace=True)
        
        # Paksa string
        df = df.astype(str) 
        
        # --- LANGKAH