from __future__ import annotations

from pathlib import Path
import traceback
from ..pipeline.models import PipelineStage, TaskResult
from ..processing.cells import _clean_cell
from ..processing.pricing import fix_original_prices, match_prices
from ..processing.shopify import build_shopify_frame
from ..processing.woocommerce import build_wp_rows
from ..storage.exports import (
    export_matched_table,
    export_raw_table,
    export_shopify_csv,
    export_unused_prices,
    export_wp_rows,
)
from ..storage.tables import _read_table, _read_shopify_rows
from ..storage.workbooks import save_xlsx
from ..storage.merge import read_folder_products


def convert_styles_file(input_file, log=print):
    """合并表 → 原价修正 → 价格匹配 → Shopify CSV → WP。

    输出目录与文件名对齐 Styles_递增价格匹配.py：
    原价/<表名>_原价.xlsx 与 价格匹配/<表名>_价格匹配.xlsx 为两张表，
    Shopify 与 WP 只输出 CSV，未匹配价格清单始终生成。
    """
    input_path = Path(input_file)
    fixed, row_fixed, segment_fixed = fix_original_prices(_read_table(input_path))
    if row_fixed or segment_fixed:
        log(f"原价修正: price2 < price1 修正 {row_fixed} 行，styles1 内嵌价修正 {segment_fixed} 个变体")
    matched, unused_prices, matched_map = match_prices(fixed)

    raw_table = export_raw_table(fixed, input_path, log)
    match_table = export_matched_table(matched, input_path, log)
    raw_csv = export_shopify_csv(build_shopify_frame(fixed), raw_table, log)
    match_csv = export_shopify_csv(build_shopify_frame(matched, matched_map), match_table, log)
    export_unused_prices(unused_prices, input_path, log)

    for csv_path in (raw_csv, match_csv):
        try:
            rows = [
                {key: _clean_cell(value) for key, value in row.items()}
                for row in _read_shopify_rows(csv_path)
            ]
            export_wp_rows(build_wp_rows(rows, log), csv_path, log)
        except Exception as exc:
            log(f"Shopify 转 WP 出错: {exc}")
            log(traceback.format_exc())
    return raw_table


def run_conversion_task(context, config):
    context.log(f"第三阶段文件夹: {config.folder}")
    merge_path, rows = read_folder_products(config.folder, context.log)
    context.checkpoints.save(str(merge_path), save_xlsx, rows, merge_path)
    context.state.last_merge_path = str(merge_path)
    context.log(f"合并完成: {merge_path}（{len(rows)} 条）")
    output = convert_styles_file(merge_path, context.log)
    return TaskResult(PipelineStage.CONVERTING, output_path=str(output))
