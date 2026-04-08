import io
import time
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from services.export_helpers import _enforce_text_format_in_memory
from services.file_loading import baca_data_penuh, baca_preview_mentah, siapkan_dataframe
from services.validation_logic import apply_cleaning_to_df, fuzzy_group_values, get_help_text


def render_cluster_controls(items, freq_map, user_winner_picks, suffix=""):
    for winner, members in items:
        with st.container():
            st.markdown(f"**Cluster ({len(members)} nilai):**")
            st.caption(", ".join([f"`{m}`" for m in members]))
            freq_sorted = sorted(members, key=lambda x: freq_map.get(x, 0), reverse=True)
            selected = st.selectbox("Pilih nilai yang dipilih:", options=freq_sorted, index=0, key=f"winner_{winner}{suffix}")
            user_winner_picks[winner] = selected
            with st.expander(" atau ketik manual...", expanded=False):
                manual_input = st.text_input("Ketik winner manual:", value="", key=f"manual_{winner}{suffix}")
                if manual_input.strip():
                    user_winner_picks[winner] = manual_input.strip()
            st.markdown("---")


def render_split_page():
    st.title("Split Workbook - Internal Antasena")
    st.caption("Pecah file Excel berdasarkan kolom tertentu menjadi beberapa file.")

    if "split_state" not in st.session_state:
        st.session_state.split_state = {
            "columns_loaded": False,
            "df_preview": None,
            "all_columns": [],
            "processing": False,
            "cancel_requested": False,
            "progress": 0,
            "files_created": [],
            "start_time": None,
        }

    uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx", "xlsm", "xls"], help=get_help_text("upload"))
    if not uploaded_file:
        st.write("<br><br>", unsafe_allow_html=True)
        return

    try:
        is_csv = uploaded_file.name.endswith(".csv")
        st.subheader("1. Konfigurasi File")
        col_file, col_header_row = st.columns([3, 1])

        with col_file:
            if not is_csv:
                daftar_sheet = pd.ExcelFile(uploaded_file, engine="openpyxl").sheet_names
                selected_sheet = st.selectbox("Sheet:", daftar_sheet)
            else:
                selected_sheet = "Sheet1"
                st.info("File CSV terdeteksi (Hanya 1 Sheet).")

        df_preview_raw = baca_preview_mentah(uploaded_file, selected_sheet, is_csv).fillna("")
        with st.expander("Klik untuk melihat Preview Data Mentah (Cek posisi Header)", expanded=False):
            st.caption("Baris ke berapa Header tabel Anda?")
            df_preview_raw.index += 1
            st.dataframe(df_preview_raw, use_container_width=True)

        with col_header_row:
            header_row_input = st.number_input("Header Table ada di baris ke:", min_value=1, value=1, help=get_help_text("header_row"))
            hapus_baris_penomoran = st.checkbox(
                "Abaikan baris nomor kolom (1, 2, 3...) Membuang 1 baris di bawah header bila terdapat urutan angka kolom",
                value=False,
                help="Otomatis membuang 1 baris tepat di bawah header jika isinya hanya urutan angka kolom.",
            )

        df_full = siapkan_dataframe(baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input), hapus_baris_penomoran)
        st.divider()
        st.subheader("2. Pilih Kolom Split")
        cols = df_full.columns.tolist()
        if not cols:
            st.error("Header tidak ditemukan.")
            return

        target_split_col = st.selectbox("Kolom Split:", cols, help=get_help_text("kolom_split"))
        if target_split_col and target_split_col in df_full.columns:
            st.divider()
            st.subheader("Preview Statistik")
            unique_vals = [v for v in df_full[target_split_col].unique() if str(v).strip() not in ("", "nan", "None")]
            freq_map = df_full[target_split_col].value_counts().to_dict()

            col_stat1, col_stat2, col_stat3 = st.columns(3)
            col_stat1.metric("Total Baris Data", len(df_full))
            col_stat2.metric("Unique Nilai Split", len(unique_vals))
            col_stat3.metric("Estimasi File Output", len(unique_vals))
            if len(unique_vals) > 100:
                st.warning(f"Perhatian: Akan ada {len(unique_vals)} file output. Proses mungkin memerlukan waktu lama.")
            st.info(get_help_text("preview_stats"))

            enable_auto_clean = st.checkbox(
                "Aktifkan Auto Cleaning (Fuzzy Match)",
                value=False,
                help="Gabungkan nilai yang serupa (misal: 'kab. boyolali' dan 'kabupaten boyolali') menjadi satu",
            )
            clusters = {}
            final_clusters = {}
            cleaned_unique_count = len(unique_vals)
            if enable_auto_clean:
                clusters = fuzzy_group_values(unique_vals, freq_map)
                cleaned_unique_count = len(clusters)

            if enable_auto_clean and clusters:
                st.divider()
                st.subheader(f"Preview Auto Cleaning: {target_split_col}")
                cluster_list = sorted(list(clusters.items()), key=lambda x: len(x[1]), reverse=True)
                multi_member_clusters = [(winner, members) for winner, members in cluster_list if len(members) > 1]
                col_stat_clean1, col_stat_clean2, col_stat_clean3 = st.columns(3)
                col_stat_clean1.metric("Klaster Ditemukan", len(clusters))
                col_stat_clean2.metric("Estimasi Penggabungan", len(unique_vals) - cleaned_unique_count)
                col_stat_clean3.metric("Estimasi File Output", cleaned_unique_count)
                st.caption("Suggestion = nilai yang paling sering muncul di data")

                show_all = st.checkbox("Tampilkan semua klaster", value=False)
                preview_limit = 15
                user_winner_picks = {}
                preview_clusters = multi_member_clusters[:preview_limit] if not show_all else multi_member_clusters
                remaining_clusters = multi_member_clusters[preview_limit:] if not show_all else []
                render_cluster_controls(preview_clusters, freq_map, user_winner_picks)
                if remaining_clusters:
                    with st.expander(f"Lihat {len(remaining_clusters)} klaster lainnya"):
                        render_cluster_controls(remaining_clusters, freq_map, user_winner_picks, "_remaining")

                for winner, members in clusters.items():
                    picked_winner = user_winner_picks.get(winner, winner)
                    final_clusters.setdefault(picked_winner, [])
                    for member in members:
                        if member != picked_winner:
                            final_clusters[picked_winner].append(member)
                cleaned_unique_count = len(final_clusters)

            st.divider()
            st.subheader("Pengaturan Format")
            enable_text_format = st.checkbox("Aktifkan format teks untuk kolom sensitif", value=False, help=get_help_text("text_format"))
            checked_columns = (
                st.multiselect(
                    "Pilih kolom yang perlu diformat teks (agar tidak terkonversi):",
                    options=cols,
                    default=[],
                    help=get_help_text("select_columns"),
                )
                if enable_text_format
                else []
            )

            st.divider()
            col_process, col_cancel = st.columns([1, 1])
            with col_process:
                btn_proses = st.button("JALANKAN PROSES", use_container_width=True, disabled=st.session_state.split_state["processing"])
            with col_cancel:
                if st.session_state.split_state["processing"]:
                    if st.button("BATALKAN", use_container_width=True):
                        st.session_state.split_state["cancel_requested"] = True
                        st.rerun()

            if btn_proses and target_split_col:
                st.session_state.split_state.update(
                    {"processing": True, "cancel_requested": False, "progress": 0, "files_created": [], "start_time": time.time()}
                )
                progress_bar = st.progress(0)
                progress_text = st.empty()
                status_text = st.empty()

                try:
                    df_split_data = apply_cleaning_to_df(df_full, target_split_col, final_clusters) if enable_auto_clean and final_clusters else df_full
                    df_split_data = df_split_data.fillna("")
                    unique_vals = [v for v in df_split_data[target_split_col].unique() if str(v).strip() not in ("", "nan", "None")]
                    total_files = len(unique_vals)
                    zip_buffer = io.BytesIO()

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, val in enumerate(unique_vals):
                            if st.session_state.split_state["cancel_requested"]:
                                status_text.warning("Proses dibatalkan. Semua file yang sudah dibuat akan dihapus.")
                                break

                            df_subset = df_split_data[df_split_data[target_split_col] == val].reset_index(drop=True)
                            safe_name = str(val).strip()
                            for char in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
                                safe_name = safe_name.replace(char, "-")
                            filename = f"{safe_name}.xlsx"

                            temp_buffer = io.BytesIO()
                            with pd.ExcelWriter(temp_buffer, engine="openpyxl") as writer:
                                df_subset.to_excel(writer, index=False, sheet_name=selected_sheet)
                            temp_buffer.seek(0)

                            if checked_columns:
                                excel_bytes = _enforce_text_format_in_memory(temp_buffer.getvalue(), selected_sheet, set(checked_columns))
                                zf.writestr(filename, excel_bytes)
                            else:
                                zf.writestr(filename, temp_buffer.getvalue())

                            st.session_state.split_state["files_created"].append(filename)
                            progress = int((i + 1) / total_files * 100)
                            st.session_state.split_state["progress"] = progress
                            progress_bar.progress(progress)
                            progress_text.text(f"{i + 1}/{total_files} files ({progress}%)")

                    if not st.session_state.split_state["cancel_requested"]:
                        elapsed = time.time() - st.session_state.split_state["start_time"]
                        zip_buffer.seek(0)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.divider()
                        st.success(
                            f"Selesai! {len(st.session_state.split_state['files_created'])} file berhasil dibuat ({elapsed:.1f} detik)"
                        )
                        st.download_button(
                            label="Download Hasil Split (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"split_result_{timestamp}.zip",
                            mime="application/zip",
                            use_container_width=True,
                        )
                    else:
                        zip_buffer.close()
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")
                finally:
                    st.session_state.split_state["processing"] = False

            if st.session_state.split_state["processing"]:
                st.progress(st.session_state.split_state["progress"])
                st.text(f"{st.session_state.split_state['progress']}%")
                if st.session_state.split_state["files_created"]:
                    st.text(f"Sedang memproses: {len(st.session_state.split_state['files_created'])} files sudah dibuat...")
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

    st.write("<br><br>", unsafe_allow_html=True)

