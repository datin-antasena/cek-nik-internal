import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Cek Validitas Multi-Kolom", layout="wide")

# --- 2. FUNGSI LOGGING (Support Multi Kolom) ---
def catat_log(nama_file, nama_sheet, rincian_per_kolom):
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Format log agar rapi untuk banyak kolom
    # Contoh: [NIK: {Unik:90, Ganda:10} | KK: {Unik:100}]
    summary_text = ""
    for col, stats in rincian_per_kolom.items():
        stat_str = ", ".join([f"{k}:{v}" for k, v in stats.items()])
        summary_text += f"[{col}: {stat_str}] "

    pesan = f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | DETAIL: {summary_text}\n"
    
    with open("activity_log.txt", "a") as f:
        f.write(pesan)

# --- 3. LOGIKA UTAMA ---
st.title("🛡️ Validasi Data Multi-Kolom")
st.info("Fitur Baru: Anda bisa memilih LEBIH DARI SATU kolom (misal: NIK dan KK) untuk diperiksa sekaligus.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- LANGKAH 1: BACA METADATA ---
        xls = pd.ExcelFile(uploaded_file)
        daftar_sheet = xls.sheet_names
        
        # --- LANGKAH 2: PILIH SHEET ---
        st.subheader("1. Konfigurasi File")
        col_sheet, col_dummy = st.columns([1, 1])
        
        with col_sheet:
            selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
        
        # --- LANGKAH 3: BACA DATA ---
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        df = df.astype(str) # Paksa string sejak awal
        
        # --- LANGKAH 4: PILIH BANYAK KOLOM (MULTI-SELECT) ---
        st.caption(f"Preview Data Sheet: **{selected_sheet}**")
        st.dataframe(df.head(), use_container_width=True)
        
        cols = df.columns.tolist()
        
        # PERUBAHAN UTAMA DISINI: multiselect
        target_cols = st.multiselect(
            "Pilih Kolom-kolom yang akan dicek (Bisa NIK, KK, No. HP, dll):", 
            cols,
            placeholder="Klik untuk memilih satu atau lebih kolom..."
        )
        
        if st.button("🚀 Proses Cek Data") and target_cols:
            with st.spinner('Sedang memproses multi-kolom...'):
                df_result = df.copy()
                log_data_all = {} # Penampung data log
                
                # --- LOOPING UNTUK SETIAP KOLOM YANG DIPILIH ---
                for col_name in target_cols:
                    
                    # 1. Bersihkan Data Spesifik Kolom Ini
                    df_result[col_name] = df_result[col_name].replace('nan', '')
                    
                    # 2. Hitung Running Count (Khusus kolom ini)
                    # Kita pakai nama variabel temporary yang unik agar tidak bentrok antar kolom
                    temp_count_col = f"__temp_count_{col_name}"
                    df_result[temp_count_col] = df_result.groupby(col_name).cumcount() + 1
                    
                    # 3. Definisikan Logika (Inner Function)
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
                    
                    # 4. Terapkan Logika
                    # Nama kolom hasil: "STATUS_NIK", "STATUS_KK", dst
                    result_col_name = f"STATUS_{col_name}"
                    df_result[result_col_name] = df_result.apply(
                        lambda row: cek_validitas(row, col_name, temp_count_col), 
                        axis=1
                    )
                    
                    # 5. Bersihkan kolom bantuan
                    df_result.drop(columns=[temp_count_col], inplace=True)
                    
                    # 6. Siapkan Data untuk Log
                    counts = df_result[result_col_name].value_counts()
                    # Grouping log biar rapi
                    col_log = {}
                    ganda_sum = 0
                    for k, v in counts.items():
                        if str(k).startswith("GANDA"):
                            ganda_sum += v
                        else:
                            col_log[k] = v
                    if ganda_sum > 0: col_log['GANDA'] = ganda_sum
                    
                    log_data_all[col_name] = col_log

                # --- SIMPAN LOG GLOBAL ---
                catat_log(uploaded_file.name, selected_sheet, log_data_all)

                # --- TAMPILKAN HASIL ---
                st.success("✅ Pemeriksaan Multi-Kolom Selesai!")
                
                # Tampilkan tabel hasil
                # Urutkan kolom agar STATUS muncul di sebelah kolom aslinya (Opsional, tapi bagus untuk UX)
                # Untuk sederhananya, kita taruh semua kolom STATUS di paling kanan
                st.dataframe(df_result, use_container_width=True)
                
                # --- EXPORT EXCEL (TEXT FORMAT) ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    sheet_export = f"Cek_{selected_sheet}"[:30]
                    df_result.to_excel(writer, index=False, sheet_name=sheet_export)
                    
                    wb = writer.book
                    ws = writer.sheets[sheet_export]
                    txt_fmt = wb.add_format({'num_format': '@'})
                    
                    # Format Text untuk semua kolom
                    for idx, col in enumerate(df_result.columns):
                        ws.set_column(idx, idx, 25, txt_fmt)
                        
                        # (Opsional) Highlight Header Kolom STATUS dengan warna beda
                        if str(col).startswith("STATUS_"):
                            # Logic tambahan jika ingin styling header, tapi default sudah oke
                            pass

                buffer.seek(0)
                st.download_button(
                    label="📥 Download Hasil Lengkap",
                    data=buffer,
                    file_name=f"MultiCheck_{selected_sheet}_{uploaded_file.name}",
                    mime="application/vnd.ms-excel"
                )
        
        elif not target_cols and uploaded_file:
            st.warning("⚠️ Silakan pilih minimal 1 kolom untuk diperiksa.")
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Admin Log
with st.sidebar:
    st.header("⚙️ Admin Log")
    if st.checkbox("Lihat Log"):
        try:
            with open("activity_log.txt", "r") as f:
                st.text(f.read())
        except: st.text("Log kosong.")
    if st.button("Hapus Log"):
        try: open("activity_log.txt", "w").close(); st.rerun()
        except: pass