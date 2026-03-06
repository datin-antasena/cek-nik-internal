import io
from datetime import datetime
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials


# ─── CONFIG ───────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Dashboard Validasi NIK/NKK Antasena", layout="wide")

STYLES = """
<style>
.footer {
    position: fixed;
    left: 0; bottom: 0;
    width: 100%;
    background-color: #f8f9fa;
    color: #6c757d;
    text-align: center;
    padding: 10px;
    font-size: 13px;
    border-top: 1px solid #dee2e6;
    z-index: 1000;
}
.stApp { margin-bottom: 80px; }

[data-testid="stMetricValue"] {
    font-size: 2rem;
    font-weight: bold;
    color: #0d6efd;
}
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
    Dikembangkan oleh <strong>POKJA DATA DAN INFORMASI</strong>
    untuk digunakan internal <strong>Antasena</strong>
</div>
"""

COLOR_MAP = {
    "UNIK":               "#28a745",
    "GANDA":              "#dc3545",
    "BUKAN ANGKA":        "#ffc107",
    "TIDAK 16 DIGIT":     "#fd7e14",
    "TERKONVERSI (000)":  "#17a2b8",
    "KOSONG":             "#6c757d",
    "SUDAH SALUR 2026":   "#6f42c1",
}


# ─── DATA FETCHING ─────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600)
def ambil_data_salur_gspread():
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_key(st.secrets["SPREADSHEET_ID"]).worksheet("BNBA")

        kolom_nik = sheet.col_values(4)
        set_nik_salur = {str(nik).strip() for nik in kolom_nik[1:] if nik}
        waktu_update = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%d %b %Y, %H:%M:%S WIB")

        return set_nik_salur, waktu_update

    except Exception as e:
        return set(), f"Gagal mengambil data: {e}"


# ─── LOGGING ──────────────────────────────────────────────────────────────────

def catat_log(nama_file, nama_sheet, rincian_per_kolom):
    waktu = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S")
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
            simple_stats["GANDA (TOTAL)"] = ganda_total

        stat_str = ", ".join(f"{k}:{v}" for k, v in simple_stats.items())
        summary_text += f"[{col}: {stat_str}] "

    pesan = (
        f"[{waktu}] FILE: {nama_file} | SHEET: {nama_sheet} | DETAIL: {summary_text}\n"
    )
    with open("activity_log.txt", "a") as f:
        f.write(pesan)


# ─── VALIDATION LOGIC ─────────────────────────────────────────────────────────

def cek_validitas(row, col_name, temp_col, referensi_salur):
    val = str(row[col_name]).replace(".0", "").strip()
    count = row[temp_col]

    if not val:                         return "KOSONG"
    if len(val) != 16:                  return "TIDAK 16 DIGIT"
    if not val.isdigit():               return "BUKAN ANGKA"
    if val.endswith("000"):             return "TERKONVERSI (000)"
    if val in referensi_salur:          return "SUDAH SALUR 2026"
    if count == 1:                      return "UNIK"
    return f"GANDA {count}"


def proses_kolom(df_result, col_name, use_auto_clean, referensi_salur):
    df_result[col_name] = df_result[col_name].replace("nan", "")

    if use_auto_clean:
        df_result[col_name] = (
            df_result[col_name]
            .str.replace(r"\.0$", "", regex=True)
            .str.replace(r"\D", "", regex=True)
        )
    else:
        df_result[col_name] = df_result[col_name].str.strip()

    temp_col = f"__temp_count_{col_name}"
    df_result[temp_col] = df_result.groupby(col_name).cumcount() + 1

    ref = referensi_salur if "NIK" in col_name.upper() else set()
    status_col = f"STATUS_{col_name}"
    df_result[status_col] = df_result.apply(
        lambda row: cek_validitas(row, col_name, temp_col, ref), axis=1
    )
    df_result.drop(columns=[temp_col], inplace=True)

    return df_result


# ─── UI HELPERS ───────────────────────────────────────────────────────────────

def render_charts(df_result, col_name):
    status_col = f"STATUS_{col_name}"
    viz_series = df_result[status_col].apply(
        lambda x: "GANDA" if str(x).startswith("GANDA") else x
    )
    data_counts = viz_series.value_counts().reset_index()
    data_counts.columns = ["Status", "Jumlah"]

    total      = len(df_result)
    total_unik = (df_result[status_col] == "UNIK").sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("Total Data",          total)
    m2.metric("Data Valid (UNIK)",   total_unik)
    m3.metric("Data Perlu Perbaikan", total - total_unik, delta_color="inverse")

    st.markdown("---")
    col_pie, col_bar = st.columns(2)

    with col_pie:
        fig_pie = px.pie(
            data_counts, values="Jumlah", names="Status",
            title=f"Persentase: {col_name}",
            color="Status", color_discrete_map=COLOR_MAP, hole=0.4,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_bar:
        fig_bar = px.bar(
            data_counts, x="Status", y="Jumlah",
            title=f"Jumlah Error: {col_name}",
            color="Status", color_discrete_map=COLOR_MAP, text="Jumlah",
        )
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(fig_bar, use_container_width=True)


def render_filtered_table(df_result, target_cols):
    df_display = df_result.copy()
    filter_cols = st.columns(len(target_cols))

    for idx, col_name in enumerate(target_cols):
        status_col = f"STATUS_{col_name}"
        with filter_cols[idx]:
            list_status = df_result[status_col].unique().tolist()
            pilihan = st.multiselect(
                f"Filter {status_col}:", options=list_status, default=list_status
            )
            df_display = df_display[df_display[status_col].isin(pilihan)]

    st.caption(f"Menampilkan {len(df_display)} dari total {len(df_result)} baris data.")
    st.dataframe(df_display, use_container_width=True)
    return df_display


def buat_excel_buffer(df_result, selected_sheet):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        sheet_name = f"Cek_{selected_sheet}"[:30]
        df_result.to_excel(writer, index=False, sheet_name=sheet_name)

        wb = writer.book
        ws = writer.sheets[sheet_name]
        txt_fmt = wb.add_format({"num_format": "@"})
        for idx in range(len(df_result.columns)):
            ws.set_column(idx, idx, 25, txt_fmt)

    buffer.seek(0)
    return buffer


def bersihkan_nama_file(nama):
    for ext in (".xlsx", ".xlsm", ".xls", ".csv"):
        if nama.endswith(ext):
            return nama[: -len(ext)]
    return nama


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────

def render_sidebar(waktu_tarik):
    with st.sidebar:
        st.info(f"⏱️ **Data Salur Terakhir Ditarik:**\n\n{waktu_tarik}")
        st.caption("Sistem mengunci memori selama 1 jam untuk mencegah blokir server Google.")
        st.divider()

        st.header("⚙️ Admin Panel")
        if st.checkbox("Lihat Log Aktivitas"):
            try:
                with open("activity_log.txt", "r") as f:
                    st.text(f.read())
            except Exception:
                st.text("Log kosong.")

        if st.button("Hapus Log"):
            try:
                open("activity_log.txt", "w").close()
                st.rerun()
            except Exception:
                pass


# ─── FILE READING ─────────────────────────────────────────────────────────────

def baca_preview_mentah(uploaded_file, selected_sheet, is_csv):
    if is_csv:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(uploaded_file, header=None, nrows=10)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=None, nrows=10, sep=";")
    else:
        return pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None, nrows=10, engine="openpyxl")


def baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input):
    header_idx = header_row_input - 1
    if is_csv:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(uploaded_file, header=header_idx)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=header_idx, sep=";")
    else:
        return pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_idx, engine="openpyxl")


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    st.markdown(STYLES, unsafe_allow_html=True)
    st.title("📊 Dashboard Validasi Data NIK/NKK - Internal Antasena")
    st.info("Fitur: Atur Posisi Header, Multi-Kolom, Multi-Sheet, Auto Cleansing, Visualisasi, & Auto-Format Text.")

    set_salur_2026, waktu_tarik = ambil_data_salur_gspread()
    render_sidebar(waktu_tarik)

    if "is_processed" not in st.session_state:
        st.session_state.is_processed = False

    # ── Upload ──
    uploaded_file = st.file_uploader("Upload file Excel/CSV", type=["xlsx", "xlsm", "xls", "csv"])
    if uploaded_file is None:
        st.write("<br><br><br>", unsafe_allow_html=True)
        return

    try:
        # ── Deteksi tipe file & daftar sheet ──
        is_csv = uploaded_file.name.endswith(".csv")
        if is_csv:
            daftar_sheet = ["Sheet1"]
        else:
            xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
            daftar_sheet = xls.sheet_names

        # ── Konfigurasi File ──
        st.subheader("1. Konfigurasi File")
        col_sheet, col_header_row = st.columns([2, 1])

        with col_sheet:
            if not is_csv:
                selected_sheet = st.selectbox("Pilih Sheet:", daftar_sheet)
            else:
                st.info("File CSV terdeteksi (Hanya 1 Sheet).")
                selected_sheet = daftar_sheet[0]

        # Preview mentah
        df_preview_raw = baca_preview_mentah(uploaded_file, selected_sheet, is_csv)
        df_preview_raw = df_preview_raw.fillna("")

        with st.expander("🔍 Klik untuk melihat Preview Data Mentah (Cek posisi Header)", expanded=False):
            st.caption("Baris ke berapa Header tabel Anda?")
            df_preview_raw.index += 1
            st.dataframe(df_preview_raw, use_container_width=True)

        with col_header_row:
            header_row_input = st.number_input("Header Table ada di baris ke:", min_value=1, value=1)
            hapus_baris_penomoran = st.checkbox(
                "Abaikan baris nomor kolom (1, 2, 3...) "
                "Membuang 1 baris dibawah header bila terdapat urutan angka kolom",
                value=False,
                help="Otomatis membuang 1 baris tepat di bawah header jika isinya hanya urutan angka kolom.",
            )

        # ── Baca data penuh ──
        df = baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input)
        df.dropna(how="all", inplace=True)

        if hapus_baris_penomoran and not df.empty:
            df = df.iloc[1:].reset_index(drop=True)

        df = df.astype(str)
        for col in df.columns:
            df[col] = df[col].replace("nan", "").str.replace(r"\.0$", "", regex=True)

        # ── Pilih Kolom ──
        st.divider()
        st.subheader("2. Pilih Kolom Data")
        cols = df.columns.tolist()

        if not cols:
            st.error("⚠️ Header tidak ditemukan.")
            return

        col_left, col_right = st.columns([3, 1])
        with col_left:
            target_cols = st.multiselect(
                "Pilih Kolom yang akan dicek:",
                cols,
                placeholder="Pilih kolom NIK, KK, dll...",
            )
        with col_right:
            use_auto_clean = st.checkbox(
                "Aktifkan Auto-Cleaning",
                value=False,
                help="Otomatis menghapus spasi, titik, strip, dan huruf.",
            )

        # ── Tombol Proses ──
        if st.button("🚀 Proses & Analisa Data") and target_cols:
            with st.spinner("Memproses data..."):
                df_result = df.copy()
                log_data_all = {}

                for col_name in target_cols:
                    df_result = proses_kolom(df_result, col_name, use_auto_clean, set_salur_2026)
                    log_data_all[col_name] = df_result[f"STATUS_{col_name}"].value_counts().to_dict()

                catat_log(uploaded_file.name, selected_sheet, log_data_all)
                st.session_state.df_result = df_result
                st.session_state.target_cols_saved = target_cols
                st.session_state.is_processed = True

        elif not target_cols and uploaded_file:
            st.warning("⚠️ Silakan pilih minimal 1 kolom dulu.")

        # ── Tampilan Hasil ──
        if (
            st.session_state.get("is_processed")
            and st.session_state.get("target_cols_saved") == target_cols
        ):
            df_result = st.session_state.df_result

            st.divider()
            st.subheader("📊 Hasil Analisa Visual")

            tabs = st.tabs([f"Analisa: {c}" for c in target_cols])
            for i, col_name in enumerate(target_cols):
                with tabs[i]:
                    render_charts(df_result, col_name)

            st.divider()
            st.subheader("📋 Tabel Data")
            render_filtered_table(df_result, target_cols)

            # ── Download ──
            buffer = buat_excel_buffer(df_result, selected_sheet)
            clean_name = bersihkan_nama_file(uploaded_file.name)
            st.download_button(
                label="📥 Download Hasil Seluruhnya (Excel)",
                data=buffer,
                file_name=f"Result_{clean_name}.xlsx",
                mime="application/vnd.ms-excel",
            )

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

    st.write("<br><br><br>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
