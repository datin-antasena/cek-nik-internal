from datetime import datetime
from difflib import SequenceMatcher
import re
from zoneinfo import ZoneInfo

import pandas as pd

from services.file_loading import siapkan_dataframe


SOURCE_FILE_COL = "SOURCE_FILE"
SOURCE_SHEET_COL = "SOURCE_SHEET"
SOURCE_ROW_COL = "SOURCE_ROW"
DUPLICATE_STATUS_COL = "MERGE_DUPLICATE_STATUS"

COLUMN_ALIASES = {
    "nik": {"nik", "nikpm", "nikpenerimamanfaat", "noktp", "nomorktp"},
    "nama": {"nama", "namalengkap", "namapm", "namapenerimamanfaat", "namapenerima"},
    "nokk": {"nokk", "kk", "nomorkk", "nokartukeluarga", "kartukeluarga"},
    "provinsi": {"provinsi", "prov", "propinsi", "province"},
    "kabupaten": {"kabupaten", "kabkota", "kab", "kota", "kabupatenkota"},
    "kecamatan": {"kecamatan", "kec"},
    "kelurahan": {"kelurahan", "desa", "kel", "ds"},
    "alamat": {"alamat", "address"},
    "tanggal_lahir": {"tanggallahir", "tgllahir", "dob", "lahir"},
    "no_hp": {"nohp", "nomorhp", "hp", "telepon", "telp", "nomortelepon"},
}


def is_csv_file(file_name: str) -> bool:
    return file_name.lower().endswith(".csv")


def get_sheet_names(uploaded_file) -> list[str]:
    if is_csv_file(uploaded_file.name):
        return ["Sheet1"]

    uploaded_file.seek(0)
    excel_file = pd.ExcelFile(uploaded_file)
    return excel_file.sheet_names


def read_workbook_sheet(uploaded_file, sheet_name: str, header_row: int, hapus_baris_penomoran: bool) -> pd.DataFrame:
    uploaded_file.seek(0)
    header_idx = header_row - 1

    if is_csv_file(uploaded_file.name):
        try:
            df = pd.read_csv(uploaded_file, header=header_idx, dtype=str)
        except Exception:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=header_idx, sep=";", dtype=str)
    else:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_idx, dtype=str)

    return siapkan_dataframe(df, hapus_baris_penomoran)


def normalize_column_name(column_name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(column_name).lower())


def _alias_key(column_name: str) -> str:
    normalized = normalize_column_name(column_name)
    for canonical, aliases in COLUMN_ALIASES.items():
        if normalized == canonical or normalized in aliases:
            return canonical
    return normalized


def _best_fuzzy_match(master_col: str, source_columns: list[str], threshold: float = 0.86) -> str:
    master_norm = normalize_column_name(master_col)
    best_source = ""
    best_score = 0.0

    for source_col in source_columns:
        score = SequenceMatcher(None, master_norm, normalize_column_name(source_col)).ratio()
        if score > best_score:
            best_score = score
            best_source = source_col

    return best_source if best_score >= threshold else ""


def default_column_mapping(master_columns: list[str], source_columns: list[str]) -> dict[str, str]:
    source_by_lower = {str(col).strip().lower(): col for col in source_columns}
    source_by_normalized = {normalize_column_name(col): col for col in source_columns}
    source_by_alias = {}
    for source_col in source_columns:
        source_by_alias.setdefault(_alias_key(source_col), source_col)

    mapping = {}
    for master_col in master_columns:
        exact_match = source_by_lower.get(str(master_col).strip().lower(), "")
        normalized_match = source_by_normalized.get(normalize_column_name(master_col), "")
        alias_match = source_by_alias.get(_alias_key(master_col), "")
        fuzzy_match = _best_fuzzy_match(master_col, source_columns)
        mapping[master_col] = exact_match or normalized_match or alias_match or fuzzy_match
    return mapping


def map_source_to_master(
    source_df: pd.DataFrame,
    master_columns: list[str],
    column_mapping: dict[str, str],
    source_file: str,
    source_sheet: str,
    include_source_metadata: bool,
) -> pd.DataFrame:
    mapped_df = pd.DataFrame(index=source_df.index)

    for master_col in master_columns:
        source_col = column_mapping.get(master_col, "")
        mapped_df[master_col] = source_df[source_col] if source_col in source_df.columns else ""

    if include_source_metadata:
        mapped_df[SOURCE_FILE_COL] = source_file
        mapped_df[SOURCE_SHEET_COL] = source_sheet
        mapped_df[SOURCE_ROW_COL] = source_df.index + 2

    return mapped_df.reset_index(drop=True)


def add_source_metadata_to_master(master_df: pd.DataFrame, source_file: str, source_sheet: str) -> pd.DataFrame:
    df = master_df.copy()
    df[SOURCE_FILE_COL] = source_file
    df[SOURCE_SHEET_COL] = source_sheet
    df[SOURCE_ROW_COL] = df.index + 2
    return df


def validate_required_columns(df: pd.DataFrame, required_columns: list[str]) -> pd.DataFrame:
    rows = []
    for col in required_columns:
        if col not in df.columns:
            rows.append({"Kolom": col, "Baris Kosong": "Kolom tidak ditemukan"})
            continue

        empty_count = df[col].astype(str).str.strip().isin(["", "nan", "None"]).sum()
        rows.append({"Kolom": col, "Baris Kosong": int(empty_count)})
    return pd.DataFrame(rows)


def mark_duplicates(df: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    if not key_columns:
        return df

    result = df.copy()
    existing_keys = [col for col in key_columns if col in result.columns]
    if not existing_keys:
        result[DUPLICATE_STATUS_COL] = "KOLOM KUNCI TIDAK DITEMUKAN"
        return result

    key_frame = result[existing_keys].astype(str).apply(lambda col: col.str.strip())
    has_empty_key = key_frame.eq("").any(axis=1)
    duplicate_mask = key_frame.duplicated(keep=False) & ~has_empty_key
    result[DUPLICATE_STATUS_COL] = "UNIK"
    result.loc[duplicate_mask, DUPLICATE_STATUS_COL] = "DUPLIKAT"
    result.loc[has_empty_key, DUPLICATE_STATUS_COL] = "KUNCI KOSONG"
    return result


def build_merge_error_frames(
    df: pd.DataFrame,
    required_columns: list[str],
    duplicate_key_columns: list[str],
) -> dict[str, pd.DataFrame]:
    duplicate_df = pd.DataFrame()
    empty_key_df = pd.DataFrame()
    required_empty_df = pd.DataFrame()

    if duplicate_key_columns and DUPLICATE_STATUS_COL in df.columns:
        duplicate_df = df[df[DUPLICATE_STATUS_COL] == "DUPLIKAT"].copy()
        empty_key_df = df[df[DUPLICATE_STATUS_COL] == "KUNCI KOSONG"].copy()

    if required_columns:
        masks = []
        for col in required_columns:
            if col in df.columns:
                masks.append(df[col].astype(str).str.strip().isin(["", "nan", "None"]))
        if masks:
            required_mask = masks[0]
            for mask in masks[1:]:
                required_mask = required_mask | mask
            required_empty_df = df[required_mask].copy()

    summary_rows = [
        {"Jenis Error": "DUPLIKAT", "Jumlah Baris": len(duplicate_df)},
        {"Jenis Error": "KUNCI DUPLIKAT KOSONG", "Jumlah Baris": len(empty_key_df)},
        {"Jenis Error": "KOLOM WAJIB KOSONG", "Jumlah Baris": len(required_empty_df)},
    ]
    error_summary_df = pd.DataFrame(summary_rows)
    return {
        "duplicate_df": duplicate_df,
        "empty_key_df": empty_key_df,
        "required_empty_df": required_empty_df,
        "error_summary_df": error_summary_df,
    }


def build_info_process_rows(
    master_label: str,
    source_labels: list[str],
    row_count: int,
    mappings: dict[str, dict[str, str]],
    validation_summary: pd.DataFrame,
    duplicate_key_columns: list[str],
) -> list[dict[str, str]]:
    now = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S WIB")
    rows = [
        {"Bagian": "Waktu Proses", "Detail": now},
        {"Bagian": "Master", "Detail": master_label},
        {"Bagian": "Jumlah Source Sheet", "Detail": str(len(source_labels))},
        {"Bagian": "Source Sheet", "Detail": "; ".join(source_labels) or "-"},
        {"Bagian": "Jumlah Baris Hasil", "Detail": str(row_count)},
        {"Bagian": "Kolom Deteksi Duplikat", "Detail": ", ".join(duplicate_key_columns) or "-"},
    ]

    for source_label, mapping in mappings.items():
        mapped_pairs = [f"{master} <- {source or '(kosong)'}" for master, source in mapping.items()]
        rows.append({"Bagian": f"Mapping {source_label}", "Detail": "; ".join(mapped_pairs)})

    if validation_summary is not None and not validation_summary.empty:
        for _, row in validation_summary.iterrows():
            rows.append({"Bagian": f"Validasi {row['Kolom']}", "Detail": f"Baris kosong: {row['Baris Kosong']}"})

    return rows
