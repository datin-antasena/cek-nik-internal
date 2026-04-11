from datetime import datetime

import pandas as pd
import streamlit as st

from services.export_helpers import buat_merge_error_report_buffer, buat_merge_excel_buffer
from services.merge_logic import (
    add_source_metadata_to_master,
    build_merge_error_frames,
    build_info_process_rows,
    default_column_mapping,
    get_sheet_names,
    map_source_to_master,
    mark_duplicates,
    read_workbook_sheet,
    validate_required_columns,
)
from services.validation_logic import get_help_text


def _file_label(index, uploaded_file):
    return f"{index + 1}. {uploaded_file.name}"


def _source_label(file_label, sheet_name):
    return f"{file_label} | {sheet_name}"


def _load_sheet_catalog(uploaded_files):
    catalog = {}
    for idx, uploaded_file in enumerate(uploaded_files):
        label = _file_label(idx, uploaded_file)
        try:
            catalog[idx] = {"label": label, "sheets": get_sheet_names(uploaded_file)}
        except Exception as exc:
            st.error(f"Gagal membaca sheet dari {uploaded_file.name}: {exc}")
            catalog[idx] = {"label": label, "sheets": []}
    return catalog


def _render_mapping_controls(master_columns, source_columns, source_label):
    default_mapping = default_column_mapping(master_columns, source_columns)
    source_options = [""] + source_columns
    mapping = {}

    st.caption("Kolom kosong berarti kolom master tersebut akan diisi kosong dari source ini.")
    for master_col in master_columns:
        default_source = default_mapping.get(master_col, "")
        default_index = source_options.index(default_source) if default_source in source_options else 0
        mapping[master_col] = st.selectbox(
            f"{master_col}",
            options=source_options,
            index=default_index,
            key=f"mapping_{source_label}_{master_col}",
            format_func=lambda value: "(kosong)" if value == "" else value,
        )
    return mapping


def _build_merge_summary(
    master_rows,
    source_row_counts,
    merged_df,
    include_master_rows,
    duplicate_key_columns,
    validation_summary,
):
    duplicate_count = 0
    if duplicate_key_columns and "MERGE_DUPLICATE_STATUS" in merged_df.columns:
        duplicate_count = int((merged_df["MERGE_DUPLICATE_STATUS"] == "DUPLIKAT").sum())

    required_empty_count = 0
    if validation_summary is not None and not validation_summary.empty:
        numeric_empty_counts = pd.to_numeric(validation_summary["Baris Kosong"], errors="coerce").fillna(0)
        required_empty_count = int(numeric_empty_counts.sum())

    return {
        "master_rows": int(master_rows),
        "source_sheet_count": len(source_row_counts),
        "source_rows": int(sum(source_row_counts.values())),
        "merged_rows": int(len(merged_df)),
        "included_master_rows": bool(include_master_rows),
        "duplicate_count": duplicate_count,
        "required_empty_count": required_empty_count,
        "source_row_counts": source_row_counts,
    }


def _render_merge_summary(summary):
    st.subheader("Rekapitulasi Proses")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Baris Master", summary["master_rows"] if summary["included_master_rows"] else 0)
    col2.metric("Source Sheet", summary["source_sheet_count"])
    col3.metric("Baris Source", summary["source_rows"])
    col4.metric("Total Baris Hasil", summary["merged_rows"])

    col5, col6 = st.columns(2)
    col5.metric("Baris Duplikat", summary["duplicate_count"])
    col6.metric("Kosong di Kolom Wajib", summary["required_empty_count"])

    source_rows = [{"Sheet Sumber": label, "Jumlah Baris": count} for label, count in summary["source_row_counts"].items()]
    if source_rows:
        with st.expander("Detail baris per source sheet", expanded=False):
            st.dataframe(pd.DataFrame(source_rows), use_container_width=True, hide_index=True)


def render_merge_page():
    st.title("Merge Workbook - Internal Antasena")
    st.caption("Gabungkan beberapa workbook/sheet ke satu sheet berdasarkan struktur kolom tabel master.")

    if "merge_result" not in st.session_state:
        st.session_state.merge_result = None

    if st.button("RESET PROSES MERGE", use_container_width=True):
        st.session_state.merge_result = None
        st.session_state.merge_upload_signature = None
        st.rerun()

    master_file = st.file_uploader(
        "Upload workbook master",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=False,
        help=get_help_text("merge_upload"),
        key="merge_master_upload",
    )
    if not master_file:
        st.session_state.merge_result = None
        st.session_state.merge_upload_signature = None
        st.write("<br><br>", unsafe_allow_html=True)
        return

    st.subheader("1. Pilih Tabel Master")
    master_catalog = _load_sheet_catalog([master_file])
    if not master_catalog.get(0, {}).get("sheets"):
        st.error("Workbook master tidak bisa dibaca.")
        return

    col_master_file, col_master_sheet, col_master_header = st.columns([2, 2, 1])
    with col_master_file:
        master_file_idx = 0
        st.text_input("Workbook master:", value=master_catalog[master_file_idx]["label"], disabled=True)
    with col_master_sheet:
        master_sheet = st.selectbox("Sheet master:", options=master_catalog[master_file_idx]["sheets"])
    with col_master_header:
        master_header_row = st.number_input("Header master di baris:", min_value=1, value=1)

    hapus_baris_penomoran_master = st.checkbox(
        "Master: abaikan baris nomor kolom (1, 2, 3...)",
        value=False,
        help="Membuang 1 baris di bawah header master bila terdapat urutan angka kolom.",
    )

    try:
        master_df = read_workbook_sheet(master_file, master_sheet, master_header_row, hapus_baris_penomoran_master)
    except Exception as exc:
        st.error(f"Gagal membaca tabel master: {exc}")
        return

    master_columns = master_df.columns.tolist()
    if not master_columns:
        st.error("Kolom master tidak ditemukan.")
        return

    with st.expander("Preview tabel master", expanded=False):
        st.dataframe(master_df.head(50), use_container_width=True)

    st.subheader("2. Pilih Sheet Sumber")
    source_uploaded_files = st.file_uploader(
        "Upload workbook sumber tambahan (opsional)",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        help="Kosongkan jika sumber data ada di sheet lain pada workbook master. Upload di sini jika sumber data ada di workbook lain.",
        key="merge_source_upload",
    )
    uploaded_files = [master_file] + list(source_uploaded_files or [])
    upload_signature = tuple((uploaded_file.name, uploaded_file.size) for uploaded_file in uploaded_files)
    if st.session_state.get("merge_upload_signature") != upload_signature:
        st.session_state.merge_result = None
        st.session_state.merge_upload_signature = upload_signature

    sheet_catalog = _load_sheet_catalog(uploaded_files)
    available_files = [idx for idx, info in sheet_catalog.items() if info["sheets"]]
    source_options = []
    source_lookup = {}
    for idx in available_files:
        for sheet_name in sheet_catalog[idx]["sheets"]:
            if idx == master_file_idx and sheet_name == master_sheet:
                continue
            label = _source_label(sheet_catalog[idx]["label"], sheet_name)
            source_options.append(label)
            source_lookup[label] = (idx, sheet_name)

    st.caption("Daftar sumber di bawah mencakup sheet lain dari workbook master dan semua sheet dari workbook sumber tambahan.")
    if not source_options:
        st.info("Belum ada sheet sumber. Tambahkan sheet lain pada workbook master atau upload workbook sumber tambahan di atas.")
        st.write("<br><br>", unsafe_allow_html=True)
        return

    selected_sources = st.multiselect(
        "Sheet yang akan diimport ke tabel master:",
        options=source_options,
        placeholder="Pilih satu atau beberapa sheet sumber...",
    )
    if not selected_sources:
        st.info("Pilih minimal 1 sheet sumber untuk melanjutkan merge.")
        st.write("<br><br>", unsafe_allow_html=True)
        return

    col_source_header, col_include_master, col_source_meta = st.columns([1, 2, 2])
    with col_source_header:
        source_header_row = st.number_input("Header source di baris:", min_value=1, value=1)
    with col_include_master:
        include_master_rows = st.checkbox("Sertakan baris dari tabel master", value=True)
    with col_source_meta:
        include_source_metadata = st.checkbox("Tambahkan SOURCE_FILE/SHEET/ROW", value=True)

    hapus_baris_penomoran_source = st.checkbox(
        "Source: abaikan baris nomor kolom (1, 2, 3...)",
        value=False,
        help="Membuang 1 baris di bawah header source bila terdapat urutan angka kolom.",
    )

    st.subheader("3. Mapping Kolom")
    source_dataframes = {}
    source_mappings = {}
    for label in selected_sources:
        file_idx, sheet_name = source_lookup[label]
        uploaded_file = uploaded_files[file_idx]
        try:
            source_df = read_workbook_sheet(uploaded_file, sheet_name, source_header_row, hapus_baris_penomoran_source)
            source_dataframes[label] = source_df
        except Exception as exc:
            st.error(f"Gagal membaca {label}: {exc}")
            continue

        with st.expander(f"Mapping: {label}", expanded=False):
            st.caption(f"Baris source: {len(source_df)} | Kolom source: {len(source_df.columns)}")
            source_mappings[label] = _render_mapping_controls(master_columns, source_df.columns.tolist(), label)

    if not source_dataframes:
        st.error("Tidak ada source sheet yang berhasil dibaca.")
        return

    st.subheader("4. Validasi & Duplikat")
    col_required, col_duplicate = st.columns(2)
    with col_required:
        required_columns = st.multiselect(
            "Kolom wajib terisi:",
            options=master_columns,
            placeholder="Contoh: NIK, NAMA, PROVINSI...",
        )
    with col_duplicate:
        duplicate_key_columns = st.multiselect(
            "Kolom kunci deteksi duplikat:",
            options=master_columns,
            placeholder="Contoh: NIK atau NIK + NAMA...",
        )

    if st.button("PROSES MERGE", use_container_width=True):
        try:
            progress_bar = st.progress(0)
            progress_text = st.empty()
            merged_parts = []
            master_label = _source_label(sheet_catalog[master_file_idx]["label"], master_sheet)
            total_steps = len(source_dataframes) + 4
            current_step = 0

            def update_progress(message):
                nonlocal current_step
                current_step += 1
                progress = min(int(current_step / total_steps * 100), 100)
                progress_bar.progress(progress)
                progress_text.text(f"{message} ({progress}%)")

            update_progress("Menyiapkan tabel master")
            if include_master_rows:
                if include_source_metadata:
                    merged_parts.append(add_source_metadata_to_master(master_df, uploaded_files[master_file_idx].name, master_sheet))
                else:
                    merged_parts.append(master_df.copy())

            source_row_counts = {}
            for label, source_df in source_dataframes.items():
                update_progress(f"Mapping source: {label}")
                file_idx, sheet_name = source_lookup[label]
                source_row_counts[label] = len(source_df)
                mapped_df = map_source_to_master(
                    source_df,
                    master_columns,
                    source_mappings.get(label, {}),
                    uploaded_files[file_idx].name,
                    sheet_name,
                    include_source_metadata,
                )
                merged_parts.append(mapped_df)

            update_progress("Menggabungkan seluruh data")
            merged_df = pd.concat(merged_parts, ignore_index=True) if merged_parts else pd.DataFrame(columns=master_columns)
            if duplicate_key_columns:
                update_progress("Menandai data duplikat")
                merged_df = mark_duplicates(merged_df, duplicate_key_columns)
            else:
                update_progress("Melewati deteksi duplikat")

            update_progress("Mengecek kolom wajib")
            validation_summary = validate_required_columns(merged_df, required_columns)
            info_rows = build_info_process_rows(
                master_label=master_label,
                source_labels=selected_sources,
                row_count=len(merged_df),
                mappings=source_mappings,
                validation_summary=validation_summary,
                duplicate_key_columns=duplicate_key_columns,
            )
            merge_summary = _build_merge_summary(
                master_rows=len(master_df),
                source_row_counts=source_row_counts,
                merged_df=merged_df,
                include_master_rows=include_master_rows,
                duplicate_key_columns=duplicate_key_columns,
                validation_summary=validation_summary,
            )
            info_rows.extend(
                [
                    {"Bagian": "Header Row Master", "Detail": str(master_header_row)},
                    {"Bagian": "Header Row Source", "Detail": str(source_header_row)},
                    {"Bagian": "Sertakan Baris Master", "Detail": "Ya" if include_master_rows else "Tidak"},
                    {"Bagian": "Tambahkan Metadata Source", "Detail": "Ya" if include_source_metadata else "Tidak"},
                    {"Bagian": "Kolom Wajib", "Detail": ", ".join(required_columns) or "-"},
                    {"Bagian": "Rekap Baris Master", "Detail": str(merge_summary["master_rows"] if include_master_rows else 0)},
                    {"Bagian": "Rekap Source Sheet", "Detail": str(merge_summary["source_sheet_count"])},
                    {"Bagian": "Rekap Baris Source", "Detail": str(merge_summary["source_rows"])},
                    {"Bagian": "Rekap Baris Hasil", "Detail": str(merge_summary["merged_rows"])},
                    {"Bagian": "Rekap Baris Duplikat", "Detail": str(merge_summary["duplicate_count"])},
                    {"Bagian": "Rekap Kosong Kolom Wajib", "Detail": str(merge_summary["required_empty_count"])},
                ]
            )
            for source_label, row_count in source_row_counts.items():
                info_rows.append({"Bagian": f"Rekap Source Rows {source_label}", "Detail": str(row_count)})

            error_frames = build_merge_error_frames(merged_df, required_columns, duplicate_key_columns)
            buffer = buat_merge_excel_buffer(merged_df, info_rows)
            error_report_buffer = buat_merge_error_report_buffer(error_frames)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            progress_bar.progress(100)
            progress_text.text("Merge selesai (100%)")

            st.session_state.merge_result = {
                "df": merged_df,
                "validation_summary": validation_summary,
                "summary": merge_summary,
                "error_frames": error_frames,
                "buffer": buffer.getvalue(),
                "error_report_buffer": error_report_buffer.getvalue(),
                "file_name": f"merge_result_{timestamp}.xlsx",
                "error_report_file_name": f"merge_error_report_{timestamp}.xlsx",
            }
        except Exception as exc:
            st.error(f"Terjadi kesalahan saat merge: {exc}")

    result = st.session_state.merge_result
    if result:
        st.divider()
        if result.get("summary"):
            _render_merge_summary(result["summary"])
        st.divider()
        st.subheader("Pemeriksaan Hasil")
        error_frames = result.get("error_frames", {})
        duplicate_df = error_frames.get("duplicate_df", pd.DataFrame())
        empty_key_df = error_frames.get("empty_key_df", pd.DataFrame())
        required_empty_df = error_frames.get("required_empty_df", pd.DataFrame())
        error_summary_df = error_frames.get("error_summary_df", pd.DataFrame())

        tab_hasil, tab_duplikat, tab_kunci_kosong, tab_wajib_kosong, tab_rekap = st.tabs(
            ["Preview Hasil", "Duplikat", "Kunci Kosong", "Kolom Wajib Kosong", "Rekap Error"]
        )
        with tab_hasil:
            st.caption(f"Total baris hasil: {len(result['df'])}")
            st.dataframe(result["df"].head(50), use_container_width=True)
        with tab_duplikat:
            if duplicate_df.empty:
                st.success("Tidak ada baris duplikat.")
            else:
                st.caption(f"Menampilkan {len(duplicate_df)} baris duplikat.")
                st.dataframe(duplicate_df, use_container_width=True)
        with tab_kunci_kosong:
            if empty_key_df.empty:
                st.success("Tidak ada baris dengan kunci duplikat kosong.")
            else:
                st.caption(f"Menampilkan {len(empty_key_df)} baris dengan kunci duplikat kosong.")
                st.dataframe(empty_key_df, use_container_width=True)
        with tab_wajib_kosong:
            if required_empty_df.empty:
                st.success("Tidak ada baris dengan kolom wajib kosong.")
            else:
                st.caption(f"Menampilkan {len(required_empty_df)} baris dengan kolom wajib kosong.")
                st.dataframe(required_empty_df, use_container_width=True)
        with tab_rekap:
            st.dataframe(error_summary_df, use_container_width=True, hide_index=True)

        validation_summary = result.get("validation_summary")
        if validation_summary is not None and not validation_summary.empty:
            with st.expander("Ringkasan validasi kolom wajib", expanded=True):
                st.dataframe(validation_summary, use_container_width=True)

        has_error_rows = any(not frame.empty for frame in (duplicate_df, empty_key_df, required_empty_df))
        if has_error_rows:
            st.download_button(
                label="Download Error Report (Excel)",
                data=result["error_report_buffer"],
                file_name=result["error_report_file_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.download_button(
            label="Download Hasil Merge (Excel)",
            data=result["buffer"],
            file_name=result["file_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.write("<br><br>", unsafe_allow_html=True)
