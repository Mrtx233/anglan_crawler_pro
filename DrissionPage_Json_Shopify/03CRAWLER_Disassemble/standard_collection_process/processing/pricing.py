from __future__ import annotations

import pandas as pd
from ..resources.price_library import PRICE_LIBRARY


def match_prices(source_df):
    df = source_df.copy(deep=True).reset_index(drop=True)
    for col in ("Handle", "Variant Price"):
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")
    # 使用 object 类型，既保留原始空值/文本，也允许写入小数价格。
    df["Variant Price"] = df["Variant Price"].astype(object)
    if "Variant Compare At Price" not in df.columns:
        df["Variant Compare At Price"] = None
    df["Variant Compare At Price"] = df["Variant Compare At Price"].astype(object)
    df["匹配价格"] = pd.Series([None] * len(df), dtype=object)

    global_used_prices = set()
    price_library_sorted = sorted(PRICE_LIBRARY)

    for handle, group in df.groupby("Handle"):
        if pd.isna(handle):
            continue
        valid_rows = []
        for idx, row in group.iterrows():
            price = pd.to_numeric(row.get("Variant Price"), errors="coerce")
            if pd.notna(price):
                valid_rows.append((idx, float(price)))
        if not valid_rows:
            continue

        available_prices = [p for p in price_library_sorted if p not in global_used_prices]
        for idx, original_price in sorted(valid_rows, key=lambda x: x[1]):
            lower_bound = original_price * 0.75
            upper_bound = original_price * 1.25
            matched = next((p for p in available_prices if lower_bound <= p <= upper_bound), None)
            if matched is not None:
                compare_at_price = round(matched * 1.2, 2)

                global_used_prices.add(matched)
                available_prices.remove(matched)

                # 保留匹配结果，方便核对价格匹配情况
                df.at[idx, "匹配价格"] = matched

                # 匹配成功后，直接更新 Shopify 售价与划线原价
                df.at[idx, "Variant Price"] = matched
                df.at[idx, "Variant Compare At Price"] = compare_at_price

    unused_prices = [p for p in price_library_sorted if p not in global_used_prices]
    return df, unused_prices
