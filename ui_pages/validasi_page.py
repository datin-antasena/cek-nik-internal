from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import plotly.express as px
import streamlit as st

from config import COLOR_MAP, COLOR_MAP_KATEGORI, STYLES
from services.export_helpers import bersihkan_nama_file, buat_excel_buffer
from services.file_loading import baca_data_penuh, baca_preview_mentah, siapkan_dataframe
from services.logging_utils import catat_log
from services.reference_data import ambil_data_salur_gspread
from services.validation_logic import proses_kolom, proses_kolom_usia


def render_charts(df_result, col_name):
    status_col = f"STATUS_{col_name}"
    viz_series = df_result[status_col].apply(lambda x: "GANDA" if str(x).startswith("GANDA") else x)
    data_counts = viz_series.value_counts().reset_index()
    data_counts.columns = ["Status", "Jumlah"]

    total = len(df_result)
    total_unik = (df_result[status_col] == "UNIK").sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Data", total)
    m2.metric("Data Valid (UNIK)", total_unik)
    m3.metric("Data Perlu Perbaikan", total - total_unik, delta_color="inverse")

    st.markdown("---")
    col_pie, col_bar = st.columns(2)

    with col_pie:
        fig_pie = px.pie(
            data_counts,
            values="Jumlah",
            names="Status",
            title=f"Persentase: {col_name}",
            color="Status",
            color_discrete_map=COLOR_MAP,
            hole=0.4,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        fig_bar = px.bar(
            data_counts,
            x="Status",
            y="Jumlah",
            title=f"Jumlah Error: {col_name}",
            color="Status",
            color_discrete_map=COLOR_MAP,
            text="Jumlah",
        )
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(fig_bar, use_container_width=True)


def render_charts_kategori_umur(df_result, kategori_col, col_tgl_lahir, usia_col, parsed_col, catatan_col):
    data_counts = df_result[kategori_col].value_counts().reset_index()
    data_counts.columns = ["Kategori", "Jumlah"]

    total = len(df_result)
    jml_anak = (df_result[kategori_col] == "ANAK").sum()
    jml_dewasa = (df_result[kategori_col] == "DEWASA").sum()
    jml_lansia = (df_result[kategori_col] == "LANSIA").sum()
    jml_invalid = (df_result[kategori_col] == "TIDAK VALID").sum()

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total Data", total)
    m2.metric("Anak", jml_anak)
    m3.metric("Dewasa", jml_dewasa)
    m4.metric("Lansia", jml_lansia)
    m5.metric("Tgl Tidak Valid", jml_invalid)

    st.markdown("---")
    col_pie, col_bar = st.columns(2)

    with col_pie:
        fig_pie = px.pie(
            data_counts,
            values="Jumlah",
            names="Kategori",
            title=f"Distribusi Kategori Umur: {col_tgl_lahir}",
            color="Kategori",
            color_discrete_map=COLOR_MAP_KATEGORI,
            hole=0.4,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        fig_bar = px.bar(
            data_counts,
            x="Kategori",
            y="Jumlah",
            title=f"Jumlah per Kategori Umur: {col_tgl_lahir}",
            color="Kategori",
            color_discrete_map=COLOR_MAP_KATEGORI,
            text="Jumlah",
        )
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(fig_bar, use_container_width=True)

    df_valid_usia = df_result[df_result[usia_col].notna()].copy()
    if not df_valid_usia.empty:
        fig_hist = px.histogram(
            df_valid_usia,
            x=usia_col,
            color=kategori_col,
            nbins=30,
            title="Distribusi Usia",
            labels={usia_col: "Usia (Tahun)", "count": "Jumlah"},
            color_discrete_map=COLOR_MAP_KATEGORI,
        )
        fig_hist.add_vline(x=18, line_dash="dash", line_color="blue", annotation_text="18 th (Dewasa)", annotation_position="top right")
        fig_hist.add_vline(x=60, line_dash="dash", line_color="red", annotation_text="60 th (Lansia)", annotation_position="top right")
        st.plotly_chart(fig_hist, use_container_width=True)

    df_gagal = df_result[df_result[parsed_col] == "TIDAK DIKENALI"]
    if not df_gagal.empty:
        with st.expander(f"{len(df_gagal)} baris tanggal tidak berhasil di-parse - klik untuk lihat detail", expanded=False):
            st.caption("Baris berikut tidak dapat dikenali formatnya. Silakan perbaiki secara manual.")
            st.dataframe(df_gagal[[col_tgl_lahir, parsed_col]].reset_index(drop=True), use_container_width=True)

    df_ambigu = df_result[df_result[catatan_col] == "Ambigu (dd/mm atau mm/dd?)"]
    if not df_ambigu.empty:
        with st.expander(f"{len(df_ambigu)} baris berformat ambigu (angka pertama <= 12) - klik untuk verifikasi", expanded=False):
            st.caption(
                "Baris-baris ini sudah diproses sesuai pilihan interpretasi Anda, namun sebaiknya diverifikasi manual karena "
                "format dd/mm dan mm/dd tidak bisa dibedakan secara otomatis."
            )
            st.dataframe(
                df_ambigu[[col_tgl_lahir, parsed_col, catatan_col, usia_col, kategori_col]].reset_index(drop=True),
                use_container_width=True,
            )


def render_filtered_table(df_result, target_cols):
    df_display = df_result.copy()
    filter_cols = st.columns(len(target_cols))

    for idx, col_name in enumerate(target_cols):
        status_col = f"STATUS_{col_name}"
        with filter_cols[idx]:
            list_status = df_result[status_col].unique().tolist()
            pilihan = st.multiselect(f"Filter {status_col}:", options=list_status, default=list_status)
            df_display = df_display[df_display[status_col].isin(pilihan)]

    st.caption(f"Menampilkan {len(df_display)} dari total {len(df_result)} baris data.")
    st.dataframe(df_display, use_container_width=True)
    return df_display


def render_filtered_table_usia(df_result, kategori_cols):
    df_display = df_result.copy()
    filter_cols = st.columns(len(kategori_cols))

    for idx, kategori_col in enumerate(kategori_cols):
        with filter_cols[idx]:
            list_kat = df_result[kategori_col].unique().tolist()
            pilihan = st.multiselect(f"Filter {kategori_col}:", options=list_kat, default=list_kat)
            df_display = df_display[df_display[kategori_col].isin(pilihan)]

    st.caption(f"Menampilkan {len(df_display)} dari total {len(df_result)} baris data.")
    st.dataframe(df_display, use_container_width=True)
    return df_display


def render_sidebar(waktu_tarik):
    with st.sidebar:
        st.info(f"Data Salur Terakhir Ditarik:\n\n{waktu_tarik}")
        st.caption("Sistem mengunci memori selama 1 jam untuk mencegah blokir server Google.")
        st.divider()

        st.header("Admin Panel")
        if st.checkbox("Lihat Log Aktivitas"):
            try:
                with open("activity_log.txt", "r") as f:
                    st.text(f.read())
            except Exception:
                st.text("Log kosong.")

        if st.button("Hapus Log"):
            try:
                with open("activity_log.txt", "w"):
                    pass
                st.rerun()
            except Exception:
                pass


def render_validasi_page():
    st.markdown(STYLES, unsafe_allow_html=True)
    st.title("Dashboard Validasi Data - Internal Antasena")
    st.info("Fitur: Atur Posisi Header, Multi-Kolom, Multi-Sheet, Auto Cleansing, Visualisasi, Auto-Format Text & Kategori Umur.")

    set_salur_2026, waktu_tarik = ambil_data_salur_gspread()
    render_sidebar(waktu_tarik)

    if "is_processed" not in st.session_state:
        st.session_state.is_processed = False

    uploaded_file = st.file_uploader("Upload file Excel/CSV", type=["xlsx", "xlsm", "xls", "csv"])
    if uploaded_file is None:
        st.write("<br><br><br>", unsafe_allow_html=True)
        return

    try:
        is_csv = uploaded_file.name.endswith(".csv")
        daftar_sheet = ["Sheet1"] if is_csv else pd.ExcelFile(uploaded_file, engine="openpyxl").sheet_names

        st.subheader("1. Konfigurasi File")
        col_sheet, col_header_row = st.columns([2, 1])
        with col_sheet:
            if not is_csv:
                selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
            else:
                st.info("File CSV terdeteksi (Hanya 1 Sheet).")
                selected_sheet = daftar_sheet[0]

        df_preview_raw = baca_preview_mentah(uploaded_file, selected_sheet, is_csv).fillna("")
        with st.expander("Klik untuk melihat Preview Data Mentah (Cek posisi Header)", expanded=False):
            st.caption("Baris ke berapa Header tabel Anda?")
            df_preview_raw.index += 1
            st.dataframe(df_preview_raw, use_container_width=True)

        with col_header_row:
            header_row_input = st.number_input("Header Table ada di baris ke:", min_value=1, value=1)
            hapus_baris_penomoran = st.checkbox(
                "Abaikan baris nomor kolom (1, 2, 3...) Membuang 1 baris di bawah header bila terdapat urutan angka kolom",
                value=False,
                help="Otomatis membuang 1 baris tepat di bawah header jika isinya hanya urutan angka kolom.",
            )

        df = siapkan_dataframe(baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input), hapus_baris_penomoran)

        st.divider()
        st.subheader("2. Pilih Kolom Data")
        cols = df.columns.tolist()
        if not cols:
            st.error("Header tidak ditemukan.")
            return

        col_left, col_right = st.columns([3, 1])
        with col_left:
            target_cols = st.multiselect("Pilih Kolom yang akan dicek (NIK/NKK):", cols, placeholder="Pilih kolom NIK, KK, dll...")
        with col_right:
            use_auto_clean = st.checkbox("Aktifkan Auto-Cleaning", value=False, help="Otomatis menghapus spasi, titik, strip, dan huruf.")

        st.divider()
        st.subheader("3. Cek Kategori Umur (Opsional)")
        st.caption(
            "Fitur ini **dapat dipilih bila terdapat kolom tanggal lahir** dari PM tersebut. Jika file tidak memiliki "
            "isian tanggal lahir, abaikan bagian ini - proses validasi NIK/NKK tetap berjalan seperti biasa."
        )
        aktifkan_cek_umur = st.checkbox(
            "Aktifkan Pengecekan Kategori Umur (Anak / Dewasa / Lansia)",
            value=False,
            help="Hanya centang jika file Anda memiliki kolom tanggal lahir.",
        )

        cols_tgl_lahir_dipilih = []
        tgl_pengecekan = datetime.now(ZoneInfo("Asia/Jakarta")).replace(tzinfo=None)
        dayfirst = True
        if aktifkan_cek_umur:
            col_umur_left, col_umur_right = st.columns([3, 1])
            with col_umur_left:
                cols_tgl_lahir_dipilih = st.multiselect(
                    "Pilih Kolom Tanggal Lahir:",
                    cols,
                    placeholder="Pilih kolom yang berisi tanggal lahir...",
                    help="Format tanggal yang didukung: dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd, serial Excel, dll.",
                )
            with col_umur_right:
                tgl_pengecekan_input = st.date_input(
                    "Tanggal Pengecekan:",
                    value=datetime.now(ZoneInfo("Asia/Jakarta")).date(),
                    format="DD/MM/YYYY",
                    help="Usia akan dihitung berdasarkan tanggal ini.",
                )
                tgl_pengecekan = datetime.combine(tgl_pengecekan_input, datetime.min.time())

            if not cols_tgl_lahir_dipilih:
                st.info("Pilih kolom tanggal lahir di atas untuk mengaktifkan analisa kategori umur.")
            else:
                st.warning(
                    "**Perhatian Format Ambigu**\n\nTanggal seperti `03/05/1990` bisa berarti **3 Mei** (dd/mm) atau **5 Maret** (mm/dd). Pilih interpretasi default di bawah untuk kasus seperti ini."
                )
                interpretasi_ambigu = st.radio(
                    "Jika format tanggal ambigu, interpretasikan angka pertama sebagai:",
                    options=["Hari (dd/mm/yyyy) - default Indonesia", "Bulan (mm/dd/yyyy) - gaya Amerika"],
                    index=0,
                    horizontal=True,
                )
                dayfirst = interpretasi_ambigu.startswith("Hari")
                st.info(
                    f"**Aturan Kategorisasi:**\n- **ANAK**: Usia < 18 tahun\n- **DEWASA**: 18 <= Usia < 60 tahun\n"
                    f"- **LANSIA**: Usia >= 60 tahun\n\nTanggal pengecekan: **{tgl_pengecekan_input.strftime('%d/%m/%Y')}** | "
                    f"Interpretasi ambigu: **{'dd/mm' if dayfirst else 'mm/dd'}**"
                )

        if st.button("Proses & Analisa Data"):
            if not target_cols:
                st.warning("Silakan pilih minimal 1 kolom NIK/NKK untuk diproses.")
            else:
                with st.spinner("Memproses data..."):
                    df_result = df.copy()
                    log_data_all = {}
                    for col_name in target_cols:
                        df_result = proses_kolom(df_result, col_name, use_auto_clean, set_salur_2026)
                        log_data_all[col_name] = df_result[f"STATUS_{col_name}"].value_counts().to_dict()

                    hasil_usia = {}
                    if aktifkan_cek_umur and cols_tgl_lahir_dipilih:
                        for col_tgl in cols_tgl_lahir_dipilih:
                            df_result, usia_col, kategori_col, parsed_col, catatan_col = proses_kolom_usia(
                                df_result,
                                col_tgl,
                                tgl_pengecekan,
                                dayfirst=dayfirst,
                            )
                            hasil_usia[col_tgl] = (usia_col, kategori_col, parsed_col, catatan_col)

                    catat_log(uploaded_file.name, selected_sheet, log_data_all)
                    st.session_state.df_result = df_result
                    st.session_state.target_cols_saved = target_cols
                    st.session_state.hasil_usia = hasil_usia
                    st.session_state.is_processed = True

        if st.session_state.get("is_processed") and st.session_state.get("target_cols_saved") == target_cols:
            df_result = st.session_state.df_result
            hasil_usia = st.session_state.get("hasil_usia", {})

            if target_cols:
                st.divider()
                st.subheader("Hasil Analisa Visual NIK/NKK")
                tabs = st.tabs([f"Analisa: {c}" for c in target_cols])
                for i, col_name in enumerate(target_cols):
                    with tabs[i]:
                        render_charts(df_result, col_name)
                st.divider()
                st.subheader("Tabel Data NIK/NKK")
                render_filtered_table(df_result, target_cols)

            if hasil_usia:
                st.divider()
                st.subheader("Hasil Analisa Kategori Umur")
                tabs_umur = st.tabs([f"Umur: {col}" for col in hasil_usia.keys()])
                for i, (col_tgl, (usia_col, kategori_col, parsed_col, catatan_col)) in enumerate(hasil_usia.items()):
                    with tabs_umur[i]:
                        render_charts_kategori_umur(df_result, kategori_col, col_tgl, usia_col, parsed_col, catatan_col)
                st.divider()
                st.subheader("Tabel Data Kategori Umur")
                render_filtered_table_usia(df_result, [v[1] for v in hasil_usia.values()])

            st.divider()
            buffer = buat_excel_buffer(df_result, selected_sheet)
            clean_name = bersihkan_nama_file(uploaded_file.name)
            st.download_button(
                label="Download Hasil Seluruhnya (Excel)",
                data=buffer,
                file_name=f"Result_{clean_name}.xlsx",
                mime="application/vnd.ms-excel",
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

    st.write("<br><br><br>", unsafe_allow_html=True)
