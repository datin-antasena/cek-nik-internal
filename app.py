import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Cek Validitas Internal Antasena", layout="wide")

# --- 2. STYLE & FOOTER (CSS) ---
# Kita menyuntikkan CSS agar footer menempel rapi di bawah
st.markdown("""
<style>
.footer {
    position: fixed;
    left: 0;
    bottom: 0;
    width: 100%;
    background-color: #f8f9fa; /* Warna abu-abu muda */
    color: #6c757d;
    text-align: center;
    padding: 10px;
    font-size: 13px;
    border-top: 1px solid #dee2e6;
    z-index: 1000;
}
.content {
    margin-bottom: 60px; /* Memberi jarak agar konten tidak tertutup footer */
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
st.info("Fitur: Multi-Kolom, Multi-Sheet, & Auto-Format Text. Data diproses aman di RAM.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- LANGKAH 1: BACA STRUKTUR FILE ---
        xls = pd.ExcelFile(uploaded_file)
        daftar_sheet = xls.sheet_names
        
        # --- LANGKAH 2: PILIH SHEET ---
        st.subheader("1. Konfigurasi File")
        col_sheet, col_dummy = st.columns([1, 1])
        
        with col_sheet:
            selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
        
        # --- LANGKAH 3: BACA DATA ---
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        df = df.astype(str) # Paksa string
        
        # --- LANGKAH 4: PILIH BANYAK KOLOM ---
        st.caption(f"Preview Data Sheet: **{selected_sheet}**")
        st.dataframe(df.head(), use_container_width=True)
        
        cols = df.columns.tolist()
        
        target_cols = st.multiselect(
            "Pilih Kolom-kolom yang akan dicek (Contoh: NIK, No. KK):", 
            cols,
            placeholder="Klik untuk memilih kolom..."
        )
        
        if st.button("🚀 Proses Cek Data") and target_cols:
            with st.spinner('Sedang memproses...'):
                df_result = df.copy()
                log_data_all = {}
                
                # --- LOOPING TIAP KOLOM ---
                for col_name in target_cols:
                    
                    # Bersihkan
                    df_result[col_name] = df_result[col_name].replace('nan', '')
                    
                    # Helper count
                    temp_count_col = f"__temp_count_{col_name}"
                    df_result[temp_count_col] = df_result.groupby(col_name).cumcount() + 1
                    
                    # Logic
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
                    
                    # Apply
                    result_col_name = f"STATUS_{col_name}"
                    df_result[result_col_name] = df_result.apply(
                        lambda row: cek_validitas(row, col_name, temp_count_col), 
                        axis=1
                    )
                    
                    df_result.drop(columns=[temp_count_col], inplace=True)
                    
                    # Stats untuk Log
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

                # Catat Log
                catat_log(uploaded_file.name, selected_sheet, log_data_all)

                # --- HASIL & DOWNLOAD ---
                st.success("✅ Pemeriksaan Selesai!")
                st.dataframe(df_result, use_container_width=True)
                
                # Export to Buffer
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
            st.warning("⚠️ Pilih minimal 1 kolom dulu.")
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

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
        
# Spacer kosong di bawah agar konten tidak tertutup footer saat discroll pol mentok
st.write("<br><br><br>", unsafe_allow_html=True)