import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Cek NIK Multi-Sheet", layout="wide")

# --- 2. FUNGSI LOGGING ---
def catat_log(nama_file, nama_sheet, total_baris, rincian_status):
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary = ", ".join([f"{k}: {v}" for k, v in rincian_status.items()])
    # Tambahkan info Sheet di log
    pesan = f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | TOTAL: {total_baris} | RESULT: {summary}\n"
    
    with open("activity_log.txt", "a") as f:
        f.write(pesan)

# --- 3. LOGIKA UTAMA ---
st.title("🛡️ Validasi NIK (Multi-Sheet Support)")
st.info("Fitur Baru: Anda sekarang bisa memilih Sheet mana yang ingin dicek sebelum data diproses.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- LANGKAH 1: BACA METADATA EXCEL (DAFTAR SHEET) ---
        # Kita gunakan pd.ExcelFile untuk mengintip isi file tanpa memuat semua data ke RAM dulu
        xls = pd.ExcelFile(uploaded_file)
        daftar_sheet = xls.sheet_names
        
        # --- LANGKAH 2: PILIH SHEET ---
        st.subheader("1. Konfigurasi File")
        col_sheet, col_dummy = st.columns([1, 1]) # Layout kolom agar rapi
        
        with col_sheet:
            selected_sheet = st.selectbox("Pilih Sheet yang akan dicek:", daftar_sheet)
        
        # --- LANGKAH 3: BACA DATA DARI SHEET TERPILIH ---
        # Baca spesifik sheet yang dipilih user
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        # Pastikan semua data dianggap string (Text)
        df = df.astype(str)
        
        # --- LANGKAH 4: PREVIEW & PILIH KOLOM ---
        st.caption(f"Menampilkan 5 baris pertama dari sheet: **{selected_sheet}**")
        st.dataframe(df.head(), use_container_width=True)
        
        cols = df.columns.tolist()
        target_col = st.selectbox("Pilih Kolom NIK:", cols)
        
        if st.button("🚀 Proses Cek Data"):
            with st.spinner('Sedang memproses...'):
                df_result = df.copy()
                
                # Pre-processing
                df_result[target_col] = df_result[target_col].replace('nan', '')
                
                # Hitung Running Count
                df_result['__temp_count'] = df_result.groupby(target_col).cumcount() + 1
                
                # --- LOGIKA VALIDASI ---
                def cek_rumus(row):
                    val = row[target_col]
                    count = row['__temp_count']
                    # Bersihkan .0 dan spasi
                    val = val.replace('.0', '').strip()
                    
                    if len(val) != 16:
                        return "NIK TIDAK 16 DIGIT"
                    elif not val.isdigit():
                        return "BUKAN ANGKA (ADA HURUF/SIMBOL)"
                    elif val.endswith("00"):
                        return "NIK TERKONVENSI"
                    elif count == 1:
                        return "UNIK"
                    else:
                        return f"GANDA {count}"

                df_result['STATUS_CEK'] = df_result.apply(cek_rumus, axis=1)
                df_result.drop(columns=['__temp_count'], inplace=True)
                
                # Logging (Sekarang mencatat nama sheet juga)
                status_counts = df_result['STATUS_CEK'].value_counts()
                log_summary = {}
                ganda_sum = 0
                for status, jumlah in status_counts.items():
                    if str(status).startswith("GANDA"):
                        ganda_sum += jumlah
                    else:
                        log_summary[status] = jumlah
                if ganda_sum > 0:
                    log_summary['GANDA (TOTAL)'] = ganda_sum
                
                catat_log(uploaded_file.name, selected_sheet, len(df_result), log_summary)

                # --- OUTPUT HASIL ---
                st.success("Selesai! Format data telah diamankan.")
                st.dataframe(df_result, use_container_width=True)
                
                # --- EXPORT KE EXCEL (TEXT FORMAT) ---
                buffer = io.BytesIO()
                
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    # Gunakan nama sheet yang sama dengan aslinya atau nama baru
                    sheet_export_name = f"Hasil_{selected_sheet}"[:30] # Excel max sheet name 31 chars
                    
                    df_result.to_excel(writer, index=False, sheet_name=sheet_export_name)
                    
                    workbook  = writer.book
                    worksheet = writer.sheets[sheet_export_name]
                    
                    text_format = workbook.add_format({'num_format': '@'})
                    
                    for idx, col in enumerate(df_result.columns):
                        worksheet.set_column(idx, idx, 25, text_format)

                buffer.seek(0)
                st.download_button(
                    label="📥 Download Hasil",
                    data=buffer,
                    file_name=f"Checked_{selected_sheet}_{uploaded_file.name}",
                    mime="application/vnd.ms-excel"
                )
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Admin
with st.sidebar:
    st.header("⚙️ Admin Log")
    if st.checkbox("Tampilkan Log"):
        try:
            with open("activity_log.txt", "r") as f:
                st.text(f.read())
        except:
            st.info("Log kosong.")
    if st.button("Hapus Log"):
        try:
            open("activity_log.txt", "w").close()
            st.rerun()
        except: pass