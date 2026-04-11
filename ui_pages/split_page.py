import io
import time
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from services.export_helpers import (
    _enforce_text_format_for_sheets_in_memory,
    _enforce_text_format_in_memory,
    sanitize_excel_sheet_name,
)
from services.file_loading import baca_data_penuh, baca_preview_mentah, siapkan_dataframe, tampilkan_nomor_baris_excel
from services.split_logic import build_output_path, build_sheet_label, build_split_summary, iter_split_groups
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


def _render_auto_clean_controls(df_full, split_columns, output_count_label):
    enable_auto_clean = st.checkbox(
        "Aktifkan Auto Cleaning (Fuzzy Match)",
        value=False,
        help="Gabungkan nilai yang serupa pada salah satu kolom split sebelum proses dijalankan.",
    )
    if not enable_auto_clean:
        return False, None, {}

    clean_target_col = st.selectbox(
        "Kolom split yang akan dibersihkan:",
        split_columns,
        help="Untuk menjaga proses tetap aman, auto cleaning diterapkan ke satu kolom split per proses.",
    )
    unique_vals = [v for v in df_full[clean_target_col].unique() if str(v).strip() not in ("", "nan", "None")]
    freq_map = df_full[clean_target_col].value_counts().to_dict()
    clusters = fuzzy_group_values(unique_vals, freq_map)
    cleaned_unique_count = len(clusters) if clusters else len(unique_vals)
    final_clusters = {}

    if clusters:
        st.divider()
        st.subheader(f"Preview Auto Cleaning: {clean_target_col}")
        cluster_list = sorted(list(clusters.items()), key=lambda x: len(x[1]), reverse=True)
        multi_member_clusters = [(winner, members) for winner, members in cluster_list if len(members) > 1]
        col_stat_clean1, col_stat_clean2, col_stat_clean3 = st.columns(3)
        col_stat_clean1.metric("Klaster Ditemukan", len(clusters))
        col_stat_clean2.metric("Estimasi Penggabungan", len(unique_vals) - cleaned_unique_count)
        col_stat_clean3.metric(output_count_label, cleaned_unique_count)
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

    return enable_auto_clean, clean_target_col, final_clusters


def _build_split_info_rows(
    uploaded_file_name,
    selected_sheet,
    header_row_input,
    split_columns,
    output_mode,
    total_rows,
    total_outputs,
    checked_columns,
    clean_target_col,
    output_key=None,
    output_rows=None,
):
    rows = [
        {"Bagian": "File Input", "Detail": uploaded_file_name},
        {"Bagian": "Sheet Input", "Detail": selected_sheet},
        {"Bagian": "Header Row", "Detail": str(header_row_input)},
        {"Bagian": "Kolom Split", "Detail": " > ".join(split_columns)},
        {"Bagian": "Mode Output", "Detail": output_mode},
        {"Bagian": "Total Baris Input", "Detail": str(total_rows)},
        {"Bagian": "Total Output", "Detail": str(total_outputs)},
        {"Bagian": "Kolom Format Teks", "Detail": ", ".join(checked_columns) or "-"},
        {"Bagian": "Auto Cleaning Kolom", "Detail": clean_target_col or "-"},
        {"Bagian": "Waktu Proses", "Detail": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ]
    if output_key is not None:
        rows.append({"Bagian": "Output Key", "Detail": " > ".join(output_key)})
    if output_rows is not None:
        rows.append({"Bagian": "Baris Output Ini", "Detail": str(output_rows)})
    return rows


def _write_info_sheet(writer, info_rows):
    pd.DataFrame(info_rows).to_excel(writer, index=False, sheet_name="INFO_PROSES")


def render_split_page():
    st.title("Split Workbook - Internal Antasena")
    st.caption("Pecah file Excel berdasarkan satu atau beberapa level kolom menjadi banyak file/sheet.")

    if st.button("RESET PROSES SPLIT", use_container_width=True):
        st.session_state.split_state = {
            "processing": False,
            "cancel_requested": False,
            "progress": 0,
            "files_created": [],
            "start_time": None,
        }
        st.rerun()

    if "split_state" not in st.session_state:
        st.session_state.split_state = {
            "processing": False,
            "cancel_requested": False,
            "progress": 0,
            "files_created": [],
            "start_time": None,
        }

    uploaded_file = st.file_uploader("Upload File Excel", type=["xlsx", "xlsm", "xls", "csv"], help=get_help_text("upload"))
    if not uploaded_file:
        st.write("<br><br>", unsafe_allow_html=True)
        return

    try:
        is_csv = uploaded_file.name.endswith(".csv")
        st.subheader("1. Konfigurasi File")
        col_file, col_header_row = st.columns([3, 1])

        with col_file:
            if not is_csv:
                daftar_sheet = pd.ExcelFile(uploaded_file).sheet_names
                selected_sheet = st.selectbox("Sheet:", daftar_sheet)
            else:
                selected_sheet = "Sheet1"
                st.info("File CSV terdeteksi (Hanya 1 Sheet).")

        df_preview_raw = baca_preview_mentah(uploaded_file, selected_sheet, is_csv).fillna("")
        with st.expander("Klik untuk melihat Preview Data Mentah (Cek posisi Header)", expanded=False):
            st.caption("Gunakan angka di kolom 'Nomor Baris Excel' sebagai input Header Table.")
            st.dataframe(tampilkan_nomor_baris_excel(df_preview_raw), use_container_width=True, hide_index=True)

        with col_header_row:
            header_row_input = st.number_input("Header Table ada di baris ke:", min_value=1, value=1, help=get_help_text("header_row"))
            hapus_baris_penomoran = st.checkbox(
                "Abaikan baris nomor kolom (1, 2, 3...) Membuang 1 baris di bawah header bila terdapat urutan angka kolom",
                value=False,
                help="Otomatis membuang 1 baris tepat di bawah header jika isinya hanya urutan angka kolom.",
            )

        df_full = siapkan_dataframe(baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input), hapus_baris_penomoran)
        st.divider()
        st.subheader("2. Pilih Kolom Split Bertingkat")
        cols = df_full.columns.tolist()
        if not cols:
            st.error("Header tidak ditemukan.")
            return

        level_count = st.number_input(
            "Jumlah level filter:",
            min_value=1,
            max_value=len(cols),
            value=1,
            help="Tambah jumlah level jika ingin split bertingkat, misalnya Provinsi > Kabupaten > Kecamatan.",
        )
        split_columns = []
        level_cols = st.columns(min(int(level_count), 4))
        for idx in range(int(level_count)):
            with level_cols[idx % len(level_cols)]:
                selected_col = st.selectbox(
                    f"Level {idx + 1}",
                    cols,
                    key=f"split_level_{idx}",
                    help=get_help_text("kolom_split_bertingkat") if idx == 0 else None,
                )
                split_columns.append(selected_col)

        if len(set(split_columns)) != len(split_columns):
            st.warning("Setiap level split harus memakai kolom yang berbeda.")
            st.write("<br><br>", unsafe_allow_html=True)
            return

        if not split_columns:
            st.info("Pilih minimal 1 kolom split untuk melanjutkan.")
            st.write("<br><br>", unsafe_allow_html=True)
            return

        st.divider()
        st.subheader("Preview Statistik")
        output_mode = st.radio(
            "Tipe Output:",
            options=["Split ke banyak workbook", "Split ke banyak sheet dalam satu workbook"],
            index=0,
            horizontal=True,
        )
        is_multi_sheet_mode = output_mode == "Split ke banyak sheet dalam satu workbook"
        output_count_label = "Estimasi Jumlah Sheet" if is_multi_sheet_mode else "Estimasi File Output"
        split_summary, empty_split_cells = build_split_summary(df_full, split_columns)
        output_count = len(split_summary)

        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        col_stat1.metric("Total Baris Data", len(df_full))
        col_stat2.metric("Level Split", len(split_columns))
        col_stat3.metric(output_count_label, output_count)
        col_stat4.metric("Sel Split Kosong", empty_split_cells)
        if output_count > 100:
            target_label = "sheet" if is_multi_sheet_mode else "file output"
            st.warning(f"Perhatian: Akan ada {output_count} {target_label}. Proses mungkin memerlukan waktu lama.")
        st.info(get_help_text("preview_stats_bertingkat"))

        with st.expander("Preview daftar output", expanded=False):
            st.dataframe(split_summary.head(100), use_container_width=True)
            if len(split_summary) > 100:
                st.caption(f"Menampilkan 100 dari {len(split_summary)} kombinasi split.")

        enable_auto_clean, clean_target_col, final_clusters = _render_auto_clean_controls(df_full, split_columns, output_count_label)

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

        if btn_proses and split_columns:
            st.session_state.split_state.update(
                {"processing": True, "cancel_requested": False, "progress": 0, "files_created": [], "start_time": time.time()}
            )
            progress_bar = st.progress(0)
            progress_text = st.empty()
            status_text = st.empty()

            try:
                if enable_auto_clean and clean_target_col and final_clusters:
                    df_split_data = apply_cleaning_to_df(df_full, clean_target_col, final_clusters)
                else:
                    df_split_data = df_full

                df_split_data = df_split_data.fillna("")
                split_groups = iter_split_groups(df_split_data, split_columns)
                total_outputs = len(split_groups)
                base_info_rows = _build_split_info_rows(
                    uploaded_file.name,
                    selected_sheet,
                    header_row_input,
                    split_columns,
                    output_mode,
                    len(df_split_data),
                    total_outputs,
                    checked_columns,
                    clean_target_col if enable_auto_clean and final_clusters else None,
                )

                if not total_outputs:
                    st.warning("Tidak ada data yang bisa diproses.")
                    return

                if is_multi_sheet_mode:
                    workbook_buffer = io.BytesIO()
                    sheet_columns_map = {}
                    used_sheet_names = ["INFO_PROSES"]

                    with pd.ExcelWriter(workbook_buffer, engine="openpyxl") as writer:
                        for i, group in enumerate(split_groups):
                            if st.session_state.split_state["cancel_requested"]:
                                status_text.warning("Proses dibatalkan. Semua sheet yang sudah dibuat akan dihapus.")
                                break

                            sheet_name = sanitize_excel_sheet_name(build_sheet_label(group.key), used_sheet_names)
                            used_sheet_names.append(sheet_name)
                            group.dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
                            sheet_columns_map[sheet_name] = set(checked_columns)
                            st.session_state.split_state["files_created"].append(sheet_name)

                            progress = int((i + 1) / total_outputs * 100)
                            st.session_state.split_state["progress"] = progress
                            progress_bar.progress(progress)
                            progress_text.text(f"{i + 1}/{total_outputs} sheet ({progress}%)")
                        _write_info_sheet(writer, base_info_rows)

                    if not st.session_state.split_state["cancel_requested"]:
                        workbook_buffer.seek(0)
                        workbook_bytes = workbook_buffer.getvalue()
                        if checked_columns:
                            workbook_bytes = _enforce_text_format_for_sheets_in_memory(workbook_bytes, sheet_columns_map)
                else:
                    zip_buffer = io.BytesIO()

                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for i, group in enumerate(split_groups):
                            if st.session_state.split_state["cancel_requested"]:
                                status_text.warning("Proses dibatalkan. Semua file yang sudah dibuat akan dihapus.")
                                break

                            filename = build_output_path(group.key)
                            temp_buffer = io.BytesIO()
                            data_sheet_name = selected_sheet[:31] or "DATA"
                            if data_sheet_name.upper() == "INFO_PROSES":
                                data_sheet_name = "DATA"
                            with pd.ExcelWriter(temp_buffer, engine="openpyxl") as writer:
                                group.dataframe.to_excel(writer, index=False, sheet_name=data_sheet_name)
                                _write_info_sheet(
                                    writer,
                                    _build_split_info_rows(
                                        uploaded_file.name,
                                        selected_sheet,
                                        header_row_input,
                                        split_columns,
                                        output_mode,
                                        len(df_split_data),
                                        total_outputs,
                                        checked_columns,
                                        clean_target_col if enable_auto_clean and final_clusters else None,
                                        output_key=group.key,
                                        output_rows=len(group.dataframe),
                                    ),
                                )
                            temp_buffer.seek(0)

                            if checked_columns:
                                excel_bytes = _enforce_text_format_in_memory(temp_buffer.getvalue(), data_sheet_name, set(checked_columns))
                                zf.writestr(filename, excel_bytes)
                            else:
                                zf.writestr(filename, temp_buffer.getvalue())

                            st.session_state.split_state["files_created"].append(filename)
                            progress = int((i + 1) / total_outputs * 100)
                            st.session_state.split_state["progress"] = progress
                            progress_bar.progress(progress)
                            progress_text.text(f"{i + 1}/{total_outputs} files ({progress}%)")

                if not st.session_state.split_state["cancel_requested"]:
                    elapsed = time.time() - st.session_state.split_state["start_time"]
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    st.divider()
                    if is_multi_sheet_mode:
                        st.success(f"Selesai! {len(st.session_state.split_state['files_created'])} sheet berhasil dibuat ({elapsed:.1f} detik)")
                        st.download_button(
                            label="Download Hasil Split (Excel)",
                            data=workbook_bytes,
                            file_name=f"split_result_{timestamp}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    else:
                        zip_buffer.seek(0)
                        st.success(f"Selesai! {len(st.session_state.split_state['files_created'])} file berhasil dibuat ({elapsed:.1f} detik)")
                        st.download_button(
                            label="Download Hasil Split (ZIP)",
                            data=zip_buffer.getvalue(),
                            file_name=f"split_result_{timestamp}.zip",
                            mime="application/zip",
                            use_container_width=True,
                        )
                else:
                    if not is_multi_sheet_mode:
                        zip_buffer.close()
            except Exception as e:
                st.error(f"Terjadi kesalahan: {e}")
            finally:
                st.session_state.split_state["processing"] = False

        if st.session_state.split_state["processing"]:
            st.progress(st.session_state.split_state["progress"])
            st.text(f"{st.session_state.split_state['progress']}%")
            if st.session_state.split_state["files_created"]:
                st.text(f"Sedang memproses: {len(st.session_state.split_state['files_created'])} output sudah dibuat...")
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

    st.write("<br><br>", unsafe_allow_html=True)
