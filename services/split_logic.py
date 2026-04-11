import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd


EMPTY_SPLIT_LABEL = "Kosong"
INVALID_PATH_CHARS = set('/\\:*?"<>|')


@dataclass
class SplitGroup:
    key: tuple[str, ...]
    dataframe: pd.DataFrame


def normalize_split_value(value) -> str:
    text = str(value).strip()
    if text.lower() in ("", "nan", "none", "<na>"):
        return EMPTY_SPLIT_LABEL
    return text


def sanitize_path_part(value) -> str:
    cleaned = "".join("-" if ch in INVALID_PATH_CHARS else ch for ch in normalize_split_value(value))
    cleaned = re.sub(r"\s+", " ", cleaned).strip().strip(".")
    return cleaned or EMPTY_SPLIT_LABEL


def build_output_path(key: Iterable[str]) -> str:
    parts = [sanitize_path_part(part) for part in key]
    if not parts:
        return f"{EMPTY_SPLIT_LABEL}.xlsx"
    parts[-1] = f"{parts[-1]}.xlsx"
    return "/".join(parts)


def build_sheet_label(key: Iterable[str]) -> str:
    return " - ".join(sanitize_path_part(part) for part in key) or EMPTY_SPLIT_LABEL


def prepare_split_dataframe(df: pd.DataFrame, split_columns: list[str]) -> pd.DataFrame:
    prepared = df.copy()
    for col in split_columns:
        prepared[col] = prepared[col].apply(normalize_split_value)
    return prepared


def iter_split_groups(df: pd.DataFrame, split_columns: list[str]) -> list[SplitGroup]:
    if not split_columns:
        return []

    prepared = prepare_split_dataframe(df, split_columns)
    groups = []
    groupby_key = split_columns[0] if len(split_columns) == 1 else split_columns

    for raw_key, df_subset in prepared.groupby(groupby_key, dropna=False, sort=True):
        key = raw_key if isinstance(raw_key, tuple) else (raw_key,)
        normalized_key = tuple(normalize_split_value(part) for part in key)
        groups.append(SplitGroup(key=normalized_key, dataframe=df_subset.reset_index(drop=True)))

    return groups


def build_split_summary(df: pd.DataFrame, split_columns: list[str]) -> tuple[pd.DataFrame, int]:
    if not split_columns:
        return pd.DataFrame(), 0

    prepared = prepare_split_dataframe(df, split_columns)
    summary = (
        prepared.groupby(split_columns, dropna=False)
        .size()
        .reset_index(name="Jumlah Baris")
        .sort_values("Jumlah Baris", ascending=False)
        .reset_index(drop=True)
    )

    empty_rows = 0
    for col in split_columns:
        empty_rows += prepared[col].eq(EMPTY_SPLIT_LABEL).sum()

    return summary, empty_rows
