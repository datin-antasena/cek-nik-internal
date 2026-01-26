import streamlit as st
import pandas as pd
import io
import plotly.express as px
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Dashboard Validasi Data Internal Sentra Antasena", layout="wide")

# --- 2. STYLE & FOOTER ---
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
    margin-bottom: 80px;
}
/* Style untuk Metrics */
[data-testid="stMetricValue"] {
    font-size: 2rem;
    font-weight: bold;
    color: #0d6efd;
}
/* Style Checkbox Auto-Clean */
.stCheckbox {
    background-color: #e2e3e5;
    padding: 10px;
    border-radius: 5px;
    border: 1px solid #ced4da;
}
</style>
<div class="footer">
    Dikembangkan oleh <strong>POKJA DATA DAN INFORMASI</strong> untuk digunakan internal <strong>SENTRA ANTASENA</strong>
</div>
""", unsafe_allow_html=True)

# --- 3. FUNGSI LOGGING ---
def catat_log(nama_file, nama_sheet, rincian_per_kolom):
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary_text = ""
    for col, stats in rincian_per_kolom.items():
        simple_stats = {}
        ganda_total = 0
        for k, v in stats.items():
            if str(k).startswith("GANDA"):
                ganda_total += v
            else:
                simple_stats[k] = v
        if ganda_total > 0:
            simple_stats['GANDA (TOTAL)'] = ganda_total
            
        stat_str = ", ".join([f"{k}:{v}" for k, v in simple_stats.items()])
        summary_text += f"[{col}: {stat_str}] "
        
    pesan = f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | DETAIL: {summary_text}\n"
    with open("activity_log.txt", "a") as f:
        f.write(pesan)

# --- 4. APLIKASI UTAMA ---
st.title("📊 Dashboard Validasi Data - Internal Antasena")
st.info("Fitur: Atur Posisi Header, Multi-Kolom, Multi-Sheet, Auto Cleansing, Visualisasi, & Auto-Format Text.")

uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # --- BACA STRUKTUR FILE ---
        xls = pd.ExcelFile(uploaded_file)
        daftar_sheet = xls.sheet_names
        
        st.subheader("1. Konfigurasi File")
        col_sheet, col_header_row = st.columns([2, 1])
        
        with col_sheet:
            selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
        
        # Preview Data Mentah
        df_preview_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, nrows=20)
        df_preview_raw = df_preview_raw.fillna('') 
        
        with st.expander("🔍 Klik untuk melihat Preview Data Mentah (Cek posisi Header)", expanded=False):
            st.caption("Baris ke berapa Header tabel Anda?")
            df_preview_raw.index += 1 
            st.dataframe(df_preview_raw, use_container_width=True)

        with col_header_row:
            header_row_input = st.number_input("Header Table ada di baris ke:", min_value=1, value=1)

        # Baca Data Full
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row_input - 1)
        df.dropna(how='all', inplace=True)
        # Konversi awal ke string agar aman
        df = df.astype(str)
        
        st.divider()
        st.subheader("2. Pilih Kolom Data")
        cols = df.columns.tolist()
        
        if len(cols) == 0:
            st.error("⚠️ Header tidak ditemukan.")
        else:
            col_left, col_right = st.columns([3, 1])
            with col_left:
                target_cols = st.multiselect(
                    "Pilih Kolom yang akan dicek:", 
                    cols,
                    placeholder="Pilih kolom NIK, KK, dll..."
                )
            with col_right:
                # Opsi untuk mematikan Auto-Clean jika tidak diinginkan (Default: Nyala)
                use_auto_clean = st.checkbox("Aktifkan Auto-Cleaning", value=True, help="Otomatis menghapus spasi, titik, strip, dan huruf.")

            if st.button("🚀 Proses & Analisa Data") and target_cols:
                with st.spinner('Sedang membersihkan dan memproses data...'):
                    df_result = df.copy()
                    log_data_all = {}
                    
                    # --- LOOPING PROSES ---
                    for col_name in target_cols:
                        # 1. Pastikan Nan jadi string kosong
                        df_result[col_name] = df_result[col_name].replace('nan', '')
                        
                        # =========================================
                        # FITUR BARU: AUTO CLEANING
                        # =========================================
                        if use_auto_clean:
                            # Regex r'\D' artinya: Cari semua karakter yang BUKAN DIGIT (0-9)
                            # Lalu ganti dengan string kosong ''
                            # Ini akan menghapus: Spasi, -, ., /, Huruf, Simbol, dll.
                            df_result[col_name] = df_result[col_name].str.replace(r'\D', '', regex=True)
                        else:
                            # Jika tidak auto clean, cuma trim spasi depan belakang standar
                            df_result[col_name] = df_result[col_name].str.strip()
                        
                        # 2. Hitung Duplikasi (Running Count)
                        temp_count_col = f"__temp_count_{col_name}"
                        df_result[temp_count_col] = df_result.groupby(col_name).cumcount() + 1
                        
                        # 3. Logika Validasi (Logic sudah lebih simple karena data sudah bersih)
                        def cek_validitas(row, c_name, c_temp):
                            val = row[c_name]
                            count = row[c_temp]
                            
                            # Karena sudah di-clean, kita tidak perlu .replace('.0') lagi secara agresif,
                            # tapi tetap berjaga-jaga jika cleaning dimatikan user.
                            val = val.replace('.0', '').strip()
                            
                            if len(val) == 0: return "KOSONG" # Handle sel kosong
                            elif len(val) != 16: return "TIDAK 16 DIGIT"
                            elif not val.isdigit(): return "BUKAN ANGKA" # Backup check
                            elif val.endswith("00"): return "TERKONVERSI (00)"
                            elif count == 1: return "UNIK"
                            else: return f"GANDA {count}"

                        result_col_name = f"STATUS_{col_name}"
                        df_result[result_col_name] = df_result.apply(
                            lambda row: cek_validitas(row, col_name, temp_count_col), axis=1
                        )
                        df_result.drop(columns=[temp_count_col], inplace=True)
                        
                        counts = df_result[result_col_name].value_counts().to_dict()
                        log_data_all[col_name] = counts

                    catat_log(uploaded_file.name, selected_sheet, log_data_all)

                    # DASHBOARD
                    st.divider()
                    st.subheader("📊 Hasil Analisa Visual")
                    tabs = st.tabs([f"Analisa: {c}" for c in target_cols])
                    
                    for i, col_name in enumerate(target_cols):
                        with tabs[i]:
                            status_col = f"STATUS_{col_name}"
                            viz_series = df_result[status_col].apply(
                                lambda x: "GANDA" if str(x).startswith("GANDA") else x
                            )
                            data_counts = viz_series.value_counts().reset_index()
                            data_counts.columns = ['Status', 'Jumlah']
                            
                            total_data = len(df_result)
                            total_unik = len(df_result[df_result[status_col] == 'UNIK'])
                            total_masalah = total_data - total_unik
                            
                            m1, m2, m3 = st.columns(3)
                            m1.metric("Total Data", total_data)
                            m2.metric("Data Valid (UNIK)", total_unik)
                            m3.metric("Data Perlu Perbaikan", total_masalah, delta_color="inverse")
                            
                            st.markdown("---")
                            col_grafik1, col_grafik2 = st.columns(2)
                            color_map = {
                                "UNIK": "#28a745",
                                "GANDA": "#dc3545",
                                "BUKAN ANGKA": "#ffc107",
                                "TIDAK 16 DIGIT": "#fd7e14",
                                "TERKONVENSI (00)": "#17a2b8",
                                "KOSONG": "#6c757d"
                            }
                            with col_grafik1:
                                fig_pie = px.pie(data_counts, values='Jumlah', names='Status', title=f'Persentase: {col_name}', color='Status', color_discrete_map=color_map, hole=0.4)
                                st.plotly_chart(fig_pie, use_container_width=True)
                            with col_grafik2:
                                fig_bar = px.bar(data_counts, x='Status', y='Jumlah', title=f'Jumlah Error: {col_name}', color='Status', color_discrete_map=color_map, text='Jumlah')
                                fig_bar.update_traces(textposition='outside')
                                st.plotly_chart(fig_bar, use_container_width=True)

                    # TABEL & DOWNLOAD
                    st.divider()
                    st.subheader("📋 Tabel Data (Sudah Dibersihkan)")
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
                        label="📥 Download Hasil Bersih & Cek (Excel)",
                        data=buffer,
                        file_name=f"Cleaned_{selected_sheet}_{uploaded_file.name}",
                        mime="application/vnd.ms-excel"
                    )
            
            elif not target_cols and uploaded_file:
                st.warning("⚠️ Silakan pilih minimal 1 kolom dulu.")
                
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Sidebar Admin (Tetap ada untuk log, tanpa login)
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