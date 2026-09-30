from __future__ import annotations

from pathlib import Path
import traceback
from ..pipeline.models import PipelineStage, TaskResult
from ..processing.cells import _clean_cell
from ..processing.pricing import match_prices
from ..processing.shopify import build_shopify_frame
from ..processing.woocommerce import build_wp_rows
from ..storage.exports import export_matched_prices, export_shopify_frame, export_wp_rows
from ..storage.tables import _read_table, _read_shopify_rows
from ..storage.workbooks import save_xlsx
from ..storage.merge import read_folder_products


def convert_styles_file(input_file, log=print):
    input_path = Path(input_file)
    frame = build_shopify_frame(_read_table(input_path))
    output_path, csv_path = export_shopify_frame(frame, input_path, log)
    matched, unused_prices = match_prices(_read_table(output_path))
    price_match_path = export_matched_prices(output_path, matched, unused_prices, log)
    try:
        rows = [
            {key: _clean_cell(value) for key, value in row.items()}
            for row in _read_shopify_rows(csv_path)
        ]
        export_wp_rows(build_wp_rows(rows, log), csv_path, log)
        price_csv = csv_path.parent / price_match_path.with_suffix(".csv").name
        if price_csv.exists():
            rows = [
                {key: _clean_cell(value) for key, value in row.items()}
                for row in _read_shopify_rows(price_csv)
            ]
            export_wp_rows(build_wp_rows(rows, log), price_csv, log)
    except Exception as exc:
        log(f"Shopify 转 WP 出错: {exc}")
        log(traceback.format_exc())
    return output_path


def run_conversion_task(context, config):
    context.log(f"第三阶段文件夹: {config.folder}")
    merge_path, rows = read_folder_products(config.folder, context.log)
    context.checkpoints.save(str(merge_path), save_xlsx, rows, merge_path)
    context.state.last_merge_path = str(merge_path)
    context.log(f"合并完成: {merge_path}（{len(rows)} 条）")
    output = convert_styles_file(merge_path, context.log)
    return TaskResult(PipelineStage.CONVERTING, output_path=str(output))
