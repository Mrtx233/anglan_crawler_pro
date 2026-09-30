from __future__ import annotations

from pathlib import Path
import csv
import pandas as pd
from ..resources.columns import WP_HEADERS
from ..resources.price_library import PRICE_LIBRARY
from ..storage.tables import _save_table

RAW_DIR_NAME = "原价"
MATCH_DIR_NAME = "价格匹配"


def _output_dirs(input_path):
    """按 Styles_递增价格匹配.py 的布局返回 (原价目录, 价格匹配目录)，并确保存在。"""
    parent = Path(input_path).parent
    raw_dir = parent / RAW_DIR_NAME
    match_dir = parent / MATCH_DIR_NAME
    raw_dir.mkdir(exist_ok=True)
    match_dir.mkdir(exist_ok=True)
    return raw_dir, match_dir


def export_raw_table(df, input_path, log=print):
    """原价修正后的表 → 原价/<表名>_原价.xlsx。"""
    raw_dir, _ = _output_dirs(input_path)
    output_path = raw_dir / f"{Path(input_path).stem}_原价.xlsx"
    _save_table(df, output_path)
    log(f"已保存原价表: {output_path}")
    return output_path


def export_matched_table(df, input_path, log=print):
    """价格匹配后的表 → 价格匹配/<表名>_价格匹配.xlsx。"""
    _, match_dir = _output_dirs(input_path)
    output_path = match_dir / f"{Path(input_path).stem}_价格匹配.xlsx"
    _save_table(df, output_path)
    log(f"已保存价格匹配表: {output_path}")
    return output_path


def export_shopify_csv(df, table_path, log=print):
    """由 <表名>_原价.xlsx / <表名>_价格匹配.xlsx 输出同目录的 <表名>_Shopify.csv。"""
    table_path = Path(table_path)
    output_path = table_path.with_name(f"{table_path.stem}_Shopify.csv")
    _save_table(df, output_path)
    log(f"已生成 Shopify 文件: {output_path}")
    return output_path


def export_unused_prices(unused_prices, input_path, log=print):
    """未匹配价格清单 → 价格匹配/<输入表名>_未匹配到的价格.csv，始终生成。"""
    _, match_dir = _output_dirs(input_path)
    output_path = match_dir / f"{Path(input_path).stem}_未匹配到的价格.csv"
    _save_table(pd.DataFrame({"未匹配价格": unused_prices}), output_path)
    log(
        f"已使用价格数: {len(PRICE_LIBRARY) - len(unused_prices)}，未使用价格数: {len(unused_prices)}"
    )
    log(f"已生成未匹配价格清单: {output_path}")
    return output_path


def export_wp_rows(rows, csv_path, log=print):
    """WP 文件与 Shopify CSV 同目录，文件名为 wp-<Shopify文件名>。"""
    csv_path = Path(csv_path)
    output_path = csv_path.parent / f"wp-{csv_path.name}"
    with open(output_path, "w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=WP_HEADERS)
        writer.writeheader()
        writer.writerows(rows)
    log(f"已生成 WP 文件: {output_path}")
    return output_path
