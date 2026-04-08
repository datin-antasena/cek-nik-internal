import pandas as pd


def baca_preview_mentah(uploaded_file, selected_sheet, is_csv):
    if is_csv:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(uploaded_file, header=None, nrows=10)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=None, nrows=10, sep=";")
    return pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet,
        header=None,
        nrows=10,
        engine="openpyxl",
    )


def baca_data_penuh(uploaded_file, selected_sheet, is_csv, header_row_input):
    header_idx = header_row_input - 1
    if is_csv:
        uploaded_file.seek(0)
        try:
            return pd.read_csv(uploaded_file, header=header_idx)
        except Exception:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=header_idx, sep=";")
    return pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet,
        header=header_idx,
        engine="openpyxl",
    )


def siapkan_dataframe(df, hapus_baris_penomoran):
    df = df.copy()
    df.dropna(how="all", inplace=True)

    if hapus_baris_penomoran and not df.empty:
        df = df.iloc[1:].reset_index(drop=True)

    df = df.astype(str)
    for col in df.columns:
        df[col] = df[col].replace("nan", "").str.replace(r"\.0$", "", regex=True)

    return df
