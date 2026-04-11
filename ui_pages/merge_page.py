from datetime import datetime

import pandas as pd
import streamlit as st

from services.export_helpers import buat_merge_excel_buffer
from services.merge_logic import (
    add_source_metadata_to_master,
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


def render_merge_page():
    st.title("Merge Workbook - Internal Antasena")
    st.caption("Gabungkan beberapa workbook/sheet ke satu sheet berdasarkan struktur kolom tabel master.")

    if "merge_result" not in st.session_state:
        st.session_state.merge_result = None

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
            merged_parts = []
            master_label = _source_label(sheet_catalog[master_file_idx]["label"], master_sheet)
            if include_master_rows:
                if include_source_metadata:
                    merged_parts.append(add_source_metadata_to_master(master_df, uploaded_files[master_file_idx].name, master_sheet))
                else:
                    merged_parts.append(master_df.copy())

            for label, source_df in source_dataframes.items():
                file_idx, sheet_name = source_lookup[label]
                mapped_df = map_source_to_master(
                    source_df,
                    master_columns,
                    source_mappings.get(label, {}),
                    uploaded_files[file_idx].name,
                    sheet_name,
                    include_source_metadata,
                )
                merged_parts.append(mapped_df)

            merged_df = pd.concat(merged_parts, ignore_index=True) if merged_parts else pd.DataFrame(columns=master_columns)
            if duplicate_key_columns:
                merged_df = mark_duplicates(merged_df, duplicate_key_columns)

            validation_summary = validate_required_columns(merged_df, required_columns)
            info_rows = build_info_process_rows(
                master_label=master_label,
                source_labels=selected_sources,
                row_count=len(merged_df),
                mappings=source_mappings,
                validation_summary=validation_summary,
                duplicate_key_columns=duplicate_key_columns,
            )
            buffer = buat_merge_excel_buffer(merged_df, info_rows)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            st.session_state.merge_result = {
                "df": merged_df,
                "validation_summary": validation_summary,
                "buffer": buffer.getvalue(),
                "file_name": f"merge_result_{timestamp}.xlsx",
            }
        except Exception as exc:
            st.error(f"Terjadi kesalahan saat merge: {exc}")

    result = st.session_state.merge_result
    if result:
        st.divider()
        st.subheader("Preview Hasil Merge")
        st.caption(f"Total baris hasil: {len(result['df'])}")
        st.dataframe(result["df"].head(50), use_container_width=True)

        validation_summary = result.get("validation_summary")
        if validation_summary is not None and not validation_summary.empty:
            with st.expander("Ringkasan validasi kolom wajib", expanded=True):
                st.dataframe(validation_summary, use_container_width=True)

        st.download_button(
            label="Download Hasil Merge (Excel)",
            data=result["buffer"],
            file_name=result["file_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.write("<br><br>", unsafe_allow_html=True)
