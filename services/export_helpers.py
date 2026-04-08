import io

import pandas as pd
from openpyxl import load_workbook


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


def sanitize_excel_sheet_name(name, existing_names=None):
    invalid_chars = set('[]:*?/\\')
    cleaned = "".join("-" if ch in invalid_chars else ch for ch in str(name).strip())
    cleaned = cleaned.strip("'")
    cleaned = cleaned or "Sheet"
    cleaned = cleaned[:31]

    existing = set(existing_names or [])
    if cleaned not in existing:
        return cleaned

    base = cleaned[:28] if len(cleaned) > 28 else cleaned
    counter = 2
    candidate = f"{base}_{counter}"
    while candidate in existing:
        counter += 1
        suffix = f"_{counter}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
    return candidate


def _enforce_text_format_in_memory(excel_bytes: bytes, sheet_name: str, selected_columns: set) -> bytes:
    return _enforce_text_format_for_sheets_in_memory(
        excel_bytes,
        {sheet_name: selected_columns},
    )


def _enforce_text_format_for_sheets_in_memory(excel_bytes: bytes, sheet_columns_map: dict[str, set]) -> bytes:
    wb = load_workbook(io.BytesIO(excel_bytes))

    for sheet_name, selected_columns in sheet_columns_map.items():
        if sheet_name not in wb.sheetnames or not selected_columns:
            continue

        ws = wb[sheet_name]
        col_indices = {}
        for col_cell in ws[1]:
            if col_cell.value in selected_columns:
                col_indices[col_cell.column] = col_cell.value

        if not col_indices:
            continue

        for col_idx in col_indices:
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                for cell in row:
                    if cell.value is not None:
                        cell.value = str(cell.value).strip()
                        cell.number_format = "@"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

