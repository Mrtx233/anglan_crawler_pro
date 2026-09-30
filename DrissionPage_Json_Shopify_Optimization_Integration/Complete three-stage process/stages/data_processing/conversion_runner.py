"""第四阶段：从已合并的工作簿生成 Shopify 和 WooCommerce 文件。"""
from pathlib import Path
from app.task_controller import checkpoint
from stages.data_processing import converter
from stages.data_processing.file_utils import load_product_rows


def convert_file(filename, log=print, stop_event=None):
    def check():
        if stop_event is not None:
            checkpoint(stop_event)
    check()
    if not str(filename).strip():
        raise ValueError('请选择合并表')
    source = Path(filename).expanduser().resolve()
    if not source.is_file() or source.suffix.lower() != '.xlsx':
        raise ValueError('请选择有效的 XLSX 合并表')
    rows = load_product_rows(source)
    if not rows:
        raise ValueError('合并表没有数据行')
    check()
    converter.styles_to_shopify(source, log, checkpoint=check)
    log(f'转换导出完成: {source.parent}')
    return {'output': str(source.parent), 'count': len(rows), 'status': '完成'}
