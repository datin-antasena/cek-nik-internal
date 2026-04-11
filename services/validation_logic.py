import re as _re
from datetime import datetime, timedelta as _timedelta
from difflib import get_close_matches as _get_close_matches

from dateutil import parser as dateutil_parser
import pandas as pd

from config import SPLIT_HELP_TEXT, TEXT_COLUMNS_KEYWORDS

_BULAN_ID = {
    "januari": "January", "februari": "February", "maret": "March",
    "april": "April", "mei": "May", "juni": "June",
    "juli": "July", "agustus": "August", "september": "September",
    "oktober": "October", "november": "November", "desember": "December",
}

_FORMATS_PASTI = [
    "%Y/%m/%d", "%Y-%m-%d",
    "%d %b %Y", "%d %B %Y",
    "%d-%b-%Y", "%d-%B-%Y",
    "%d/%b/%Y", "%d/%B/%Y",
    "%d %b %y", "%d %B %y",
]

_FORMATS_AMBIGU_DAYFIRST = [
    "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d %m %Y", "%d %m %y",
    "%d|%m|%Y", "%d|%m|%y",
]
_FORMATS_AMBIGU_MONTHFIRST = [
    "%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y", "%m %d %Y", "%m %d %y",
    "%m|%d|%Y", "%m|%d|%y",
]

ADMIN_TYPE_TOKENS = {
    "kabupaten": "kabupaten",
    "kota": "kota",
    "kecamatan": "kecamatan",
    "kelurahan": "kelurahan",
    "desa": "desa",
    "kab": "kabupaten",
    "kec": "kecamatan",
    "kel": "kelurahan",
    "ds": "desa",
}


def cek_validitas(row, col_name, temp_col, referensi_salur):
    val = str(row[col_name]).replace(".0", "").strip()
    count = row[temp_col]

    if not val:
        return "KOSONG"
    if len(val) != 16:
        return "TIDAK 16 DIGIT"
    if not val.isdigit():
        return "BUKAN ANGKA"
    if val.endswith("000"):
        return "TERKONVERSI (000)"
    if val in referensi_salur:
        return "SUDAH SALUR 2026"
    if count == 1:
        return "UNIK"
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


def _ganti_bulan_id(tgl_str: str) -> str:
    kata_list = _re.findall(r"[a-zA-Z]{3,}", tgl_str)
    hasil = tgl_str
    for kata in kata_list:
        kata_lower = kata.lower()
        if kata_lower in _BULAN_ID:
            hasil = _re.sub(kata, _BULAN_ID[kata_lower], hasil, flags=_re.IGNORECASE)
        else:
            cutoff = 0.75 if len(kata_lower) <= 5 else 0.6
            cocok = _get_close_matches(kata_lower, _BULAN_ID.keys(), n=1, cutoff=cutoff)
            if cocok:
                hasil = _re.sub(kata, _BULAN_ID[cocok[0]], hasil, flags=_re.IGNORECASE)
    return hasil


def _angka_bagian(tgl_str: str):
    return _re.findall(r"\d+", tgl_str)


def _parse_tanggal(tgl_str: str, dayfirst: bool = True) -> tuple[datetime | None, bool]:
    tgl_str = tgl_str.strip()
    if not tgl_str or tgl_str.lower() in ("nan", "none", "-", ""):
        return None, False

    tgl_str = _ganti_bulan_id(tgl_str)

    if tgl_str.isdigit() and 10000 < int(tgl_str) < 60000:
        try:
            return datetime(1899, 12, 30) + _timedelta(days=int(tgl_str)), False
        except Exception:
            pass

    for fmt in _FORMATS_PASTI:
        try:
            return datetime.strptime(tgl_str, fmt), False
        except ValueError:
            continue

    bagian = _angka_bagian(tgl_str)
    if len(bagian) >= 2:
        angka_pertama = int(bagian[0])

        if angka_pertama > 12:
            for fmt in _FORMATS_AMBIGU_DAYFIRST:
                try:
                    return datetime.strptime(tgl_str, fmt), False
                except ValueError:
                    continue
        else:
            formats_utama = _FORMATS_AMBIGU_DAYFIRST if dayfirst else _FORMATS_AMBIGU_MONTHFIRST
            formats_alt = _FORMATS_AMBIGU_MONTHFIRST if dayfirst else _FORMATS_AMBIGU_DAYFIRST

            for fmt in formats_utama:
                try:
                    return datetime.strptime(tgl_str, fmt), True
                except ValueError:
                    continue

            for fmt in formats_alt:
                try:
                    return datetime.strptime(tgl_str, fmt), True
                except ValueError:
                    continue

    try:
        return dateutil_parser.parse(tgl_str, dayfirst=dayfirst), False
    except Exception:
        pass

    return None, False


def tentukan_kategori_umur(usia):
    if usia is None:
        return "TIDAK VALID"
    if usia < 18:
        return "ANAK"
    if usia >= 60:
        return "LANSIA"
    return "DEWASA"


def proses_kolom_usia(df_result, col_tgl_lahir, tgl_pengecekan, dayfirst: bool = True):
    usia_col = f"USIA_{col_tgl_lahir}"
    kategori_col = f"KATEGORI_UMUR_{col_tgl_lahir}"
    parsed_col = f"TGL_PARSED_{col_tgl_lahir}"
    catatan_col = f"CATATAN_PARSE_{col_tgl_lahir}"

    def _parse_row(x):
        tgl, is_ambigu = _parse_tanggal(str(x), dayfirst=dayfirst)
        if tgl is None:
            return None, None, "TIDAK DIKENALI", "Format tidak dikenali"
        if tgl > tgl_pengecekan:
            return None, None, "TIDAK DIKENALI", "Tanggal di masa depan"
        usia = (
            tgl_pengecekan.year - tgl.year
            - ((tgl_pengecekan.month, tgl_pengecekan.day) < (tgl.month, tgl.day))
        )
        if usia > 130:
            return None, None, "TIDAK DIKENALI", "Usia > 130 tahun"
        catatan = "Ambigu (dd/mm atau mm/dd?)" if is_ambigu else "OK"
        return usia, tentukan_kategori_umur(usia), tgl.strftime("%d/%m/%Y"), catatan

    hasil = df_result[col_tgl_lahir].apply(_parse_row)
    df_result[usia_col] = hasil.apply(lambda x: x[0])
    df_result[kategori_col] = hasil.apply(lambda x: x[1] if x[1] else "TIDAK VALID")
    df_result[parsed_col] = hasil.apply(lambda x: x[2])
    df_result[catatan_col] = hasil.apply(lambda x: x[3])
    return df_result, usia_col, kategori_col, parsed_col, catatan_col


def get_help_text(section: str) -> str:
    return SPLIT_HELP_TEXT.get(section, "")


def auto_detect_text_columns(columns: list) -> set:
    detected = set()
    for col in columns:
        col_upper = str(col).upper()
        for keyword in TEXT_COLUMNS_KEYWORDS:
            if keyword.upper() in col_upper:
                detected.add(col)
                break
    return detected


def _normalize_col_name(col) -> str:
    return _re.sub(r"[^A-Z0-9]", "", str(col).upper())


def auto_detect_identity_columns(columns: list) -> list:
    detected = []
    for col in columns:
        normalized = _normalize_col_name(col)
        if normalized in ("NIK", "NONIK", "NIKPM", "NIKPENERIMAMANFAAT", "KK", "NOKK", "NKK", "NOKARTUKELUARGA"):
            detected.append(col)
        elif "NIK" in normalized or "NOKK" in normalized or normalized.endswith("KK"):
            detected.append(col)
    return detected


def auto_detect_birthdate_columns(columns: list) -> list:
    detected = []
    for col in columns:
        normalized = _normalize_col_name(col)
        if any(keyword in normalized for keyword in ("TGLLAHIR", "TANGGALLAHIR", "TTL", "DOB", "LAHIR")):
            detected.append(col)
    return detected


def build_validation_error_frames(df_result: pd.DataFrame, target_cols: list, usia_result: dict | None = None) -> dict[str, pd.DataFrame]:
    duplicate_frames = []
    salur_frames = []
    empty_frames = []
    invalid_frames = []

    for col_name in target_cols:
        status_col = f"STATUS_{col_name}"
        if status_col not in df_result.columns:
            continue

        df_status = df_result.copy()
        if "ERROR_KOLOM_DICEK" in df_status.columns:
            df_status["ERROR_KOLOM_DICEK"] = col_name
        else:
            df_status.insert(0, "ERROR_KOLOM_DICEK", col_name)
        duplicate_mask = df_status[status_col].astype(str).str.startswith("GANDA")
        salur_mask = df_status[status_col].eq("SUDAH SALUR 2026")
        empty_mask = df_status[status_col].eq("KOSONG")
        invalid_mask = ~df_status[status_col].isin(["UNIK", "KOSONG", "SUDAH SALUR 2026"]) & ~duplicate_mask

        if duplicate_mask.any():
            duplicate_frames.append(df_status.loc[duplicate_mask])
        if salur_mask.any():
            salur_frames.append(df_status.loc[salur_mask])
        if empty_mask.any():
            empty_frames.append(df_status.loc[empty_mask])
        if invalid_mask.any():
            invalid_frames.append(df_status.loc[invalid_mask])

    usia_invalid_frames = []
    for col_tgl, (_, kategori_col, parsed_col, catatan_col) in (usia_result or {}).items():
        if kategori_col not in df_result.columns:
            continue
        mask = df_result[kategori_col].eq("TIDAK VALID")
        cols = [col for col in [col_tgl, parsed_col, catatan_col, kategori_col] if col in df_result.columns]
        if mask.any() and cols:
            usia_invalid_frames.append(df_result.loc[mask, cols])

    def _concat(frames):
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    duplicate_df = _concat(duplicate_frames)
    salur_df = _concat(salur_frames)
    empty_df = _concat(empty_frames)
    invalid_df = _concat(invalid_frames)
    usia_invalid_df = _concat(usia_invalid_frames)
    summary_df = pd.DataFrame(
        [
            {"Jenis Error": "DUPLIKAT", "Jumlah Baris": len(duplicate_df)},
            {"Jenis Error": "SUDAH SALUR 2026", "Jumlah Baris": len(salur_df)},
            {"Jenis Error": "KOSONG", "Jumlah Baris": len(empty_df)},
            {"Jenis Error": "TIDAK VALID NIK/NKK", "Jumlah Baris": len(invalid_df)},
            {"Jenis Error": "TANGGAL LAHIR TIDAK VALID", "Jumlah Baris": len(usia_invalid_df)},
        ]
    )
    return {
        "duplicate_df": duplicate_df,
        "salur_df": salur_df,
        "empty_df": empty_df,
        "invalid_df": invalid_df,
        "usia_invalid_df": usia_invalid_df,
        "summary_df": summary_df,
    }


def _tokenize(val: str) -> set:
    val = str(val).lower().strip()
    val = val.replace(".", " ").replace(",", " ").replace("-", " ")
    val = val.replace("_", " ").replace("/", " ").replace("\\", " ")
    tokens = set(val.split())
    return {t for t in tokens if len(t) > 1}


def _get_admin_type(tokens: set) -> str:
    for token in tokens:
        if token in ADMIN_TYPE_TOKENS:
            return ADMIN_TYPE_TOKENS[token]
    return None


def _is_similar(val1: str, val2: str) -> bool:
    tokens1 = _tokenize(val1)
    tokens2 = _tokenize(val2)

    type1 = _get_admin_type(tokens1)
    type2 = _get_admin_type(tokens2)
    if type1 and type2 and type1 != type2:
        return False

    non_type_tokens1 = tokens1 - set(ADMIN_TYPE_TOKENS.keys())
    non_type_tokens2 = tokens2 - set(ADMIN_TYPE_TOKENS.keys())
    return bool(non_type_tokens1 & non_type_tokens2)


def fuzzy_group_values(unique_values: list, frequency_map: dict) -> dict:
    if not unique_values:
        return {}

    values_to_check = list(unique_values)
    clusters = {}
    used = set()

    for i, val in enumerate(values_to_check):
        if val in used:
            continue

        cluster = [val]
        used.add(val)

        for j, other in enumerate(values_to_check):
            if other in used or i == j:
                continue
            if _is_similar(val, other):
                cluster.append(other)
                used.add(other)

        winner = max(cluster, key=lambda x: frequency_map.get(x, 0))
        clusters[winner] = cluster

    return clusters


def apply_cleaning_to_df(df: pd.DataFrame, col: str, clusters: dict) -> pd.DataFrame:
    df = df.copy()
    value_mapping = {}
    for winner, members in clusters.items():
        for member in members:
            value_mapping[member] = winner
    df[col] = df[col].replace(value_mapping)
    return df
