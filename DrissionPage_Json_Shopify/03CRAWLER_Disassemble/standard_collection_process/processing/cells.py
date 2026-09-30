from __future__ import annotations

import pandas as pd


def _clean_cell(value):
    if pd.isna(value):
        return ""
    return str(value).strip()
