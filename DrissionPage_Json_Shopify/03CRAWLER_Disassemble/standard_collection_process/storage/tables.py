from __future__ import annotations

from pathlib import Path
import csv
import pandas as pd
from ..processing.cells import _clean_cell


def _detect_encoding(path):
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                f.read(2048)
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8-sig"


def _read_table(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path, encoding=_detect_encoding(path))


def _save_table(df, path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False, encoding="utf-8-sig")


def _read_csv_rows(path):
    encoding = _detect_encoding(path)
    with open(path, encoding=encoding, newline="") as f:
        return list(csv.DictReader(f)), encoding


def _read_shopify_rows(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path).fillna("")
        return [
            {key: _clean_cell(value) for key, value in row.items()} for row in df.to_dict("records")
        ]
    rows, _ = _read_csv_rows(path)
    return rows
