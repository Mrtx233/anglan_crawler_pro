from __future__ import annotations

from pathlib import Path
import csv
import pandas as pd
from ..resources.columns import WP_HEADERS
from ..resources.price_library import PRICE_LIBRARY
from ..storage.tables import _save_table


def export_shopify_frame(df, input_path, log=print):
    input_path = Path(input_path)
    xlsx_dir, csv_dir = input_path.parent / "xlsx", input_path.parent / "csv"
    xlsx_dir.mkdir(exist_ok=True)
    csv_dir.mkdir(exist_ok=True)
    output_path = xlsx_dir / f"{input_path.stem}_Shopify_原价.xlsx"
    csv_path = csv_dir / f"{input_path.stem}_Shopify_原价.csv"
    _save_table(df, output_path)
    _save_table(df, csv_path)
    log(f"已生成 Shopify 文件: {output_path}")
    log(f"已生成 CSV 文件: {csv_path}")
    return output_path, csv_path


def export_matched_prices(input_path, df, unused_prices, log=print):
    input_path = Path(input_path)
    source_stem = input_path.stem
    for suffix in ("_Shopify_待匹配", "_Shopify_原价_拆分", "_Shopify_原价", "_原价"):
        if source_stem.endswith(suffix):
            source_stem = source_stem[: -len(suffix)]
            break
    base_dir = input_path.parent.parent if input_path.parent.name == "xlsx" else input_path.parent
    xlsx_dir, csv_dir = base_dir / "xlsx", base_dir / "csv"
    xlsx_dir.mkdir(exist_ok=True)
    csv_dir.mkdir(exist_ok=True)
    output_path = xlsx_dir / f"{source_stem}_Shopify_价格匹配{input_path.suffix}"
    csv_output = csv_dir / f"{source_stem}_Shopify_价格匹配.csv"
    _save_table(df, output_path)
    _save_table(df, csv_output)
    log(f"已生成价格匹配文件: {output_path}")
    log(f"已生成价格匹配 CSV: {csv_output}")
    if unused_prices:
        unused = pd.DataFrame({"未匹配价格": unused_prices})
        _save_table(unused, xlsx_dir / f"{source_stem}_Shopify_未匹配到的价格{input_path.suffix}")
        _save_table(unused, csv_dir / f"{source_stem}_Shopify_未匹配到的价格.csv")
    log(
        f"已使用价格数: {len(PRICE_LIBRARY) - len(unused_prices)}，未使用价格数: {len(unused_prices)}"
    )
    return output_path


def export_wp_rows(rows, input_path, log=print):
    input_path = Path(input_path)
    base_stem = input_path.stem
    for suffix in ("_Shopify_价格匹配", "_Shopify_原价"):
        if base_stem.endswith(suffix):
            base_stem = base_stem[: -len(suffix)]
            break
    output_dir = input_path.parent / f"wp-{base_stem}"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"wp-{input_path.stem}.csv"
    with open(output_path, "w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=WP_HEADERS)
        writer.writeheader()
        writer.writerows(rows)
    log(f"已生成 WP 文件: {output_path}")
    return output_path
