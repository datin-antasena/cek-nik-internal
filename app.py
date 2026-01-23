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
        
        # --- LANGKAH 4: PILIH KOLOM ---
        st.divider()
        st.subheader("2. Pilih Kolom Data")
        cols = df.columns.tolist()
        
        # Validasi jika kolom kosong (biasanya karena salah pilih baris header)
        if len(cols) == 0:
            st.error("⚠️ Tidak ditemukan nama kolom. Coba cek kembali nomor baris Header di atas.")
        else:
            target_cols = st.multiselect(
                "Pilih Kolom yang akan dicek (Contoh: NIK, No. KK):", 
                cols,
                placeholder="Klik untuk memilih kolom..."
            )
            
            if st.button("🚀 Proses Cek Data") and target_cols:
                with st.spinner('Sedang memproses...'):
                    df_result = df.copy()
                    log_data_all = {}
                    
                    # --- LOOPING TIAP KOLOM ---
                    for col_name in target_cols:
                        
                        df_result[col_name] = df_result[col_name].replace('nan', '')
                        
                        temp_count_col = f"__temp_count_{col_name}"
                        df_result[temp_count_col] = df_result.groupby(col_name).cumcount() + 1
                        
                        def cek_validitas(row, c_name, c_temp):
                            val = row[c_name]
                            count = row[c_temp]
                            val = val.replace('.0', '').strip()
                            
                            if len(val) != 16:
                                return "TIDAK 16 DIGIT"
                            elif not val.isdigit():
                                return "BUKAN ANGKA"
                            elif val.endswith("00"):
                                return "TERKONVENSI (00)"
                            elif count == 1:
                                return "UNIK"
                            else:
                                return f"GANDA {count}"
                        
                        result_col_name = f"STATUS_{col_name}"
                        df_result[result_col_name] = df_result.apply(
                            lambda row: cek_validitas(row, col_name, temp_count_col), 
                            axis=1
                        )
                        
                        df_result.drop(columns=[temp_count_col], inplace=True)
                        
                        # Logging Stats
                        counts = df_result[result_col_name].value_counts()
                        col_log = {}
                        ganda_sum = 0
                        for k, v in counts.items():
                            if str(k).startswith("GANDA"):
                                ganda_sum += v
                            else:
                                col_log[k] = v
                        if ganda_sum > 0: col_log['GANDA'] = ganda_sum
                        log_data_all[col_name] = col_log

                    catat_log(uploaded_file.name, selected_sheet, log_data_all)

                    # --- HASIL & DOWNLOAD ---
                    st.success("✅ Pemeriksaan Selesai!")
                    st.dataframe(df_result, use_container_width=True)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        sheet_export = f"Cek_{selected_sheet}"[:30]
                        df_result.to_excel(writer, index=False, sheet_name=sheet_export)
                        
                        wb = writer.book
                        ws = writer.sheets[sheet_export]
                        txt_fmt = wb.add_format({'num_format': '@'})
                        
                        for idx, col in enumerate(df_result.columns):
                            ws.set_column(idx, idx, 25, txt_fmt)

                    buffer.seek(0)
                    st.download_button(
                        label="📥 Download Hasil Lengkap",
                        data=buffer,
                        file_name=f"MultiCheck_{selected_sheet}_{uploaded_file.name}",
                        mime="application/vnd.ms-excel"
                    )
            
            elif not target_cols and uploaded_file:
                st.warning("⚠️ Silakan pilih minimal 1 kolom dulu.")
                
    except Exception as e:
        st.error(f"Terjadi kesalahan pembacaan file: {e}")
        st.warning("Tips: Pastikan 'Header ada di baris ke' sudah sesuai dengan file Excel Anda.")

# Admin Sidebar
with st.sidebar:
    st.header("⚙️ Admin Panel")
    if st.checkbox("Lihat Log Aktivitas"):
        try:
            with open("activity_log.txt", "r") as f:
                st.text(f.read())
        except: st.text("Log kosong.")
    if st.button("Hapus Log"):
        try: open("activity_log.txt", "w").close(); st.rerun()
        except: pass

st.write("<br><br><br>", unsafe_allow_html=True)