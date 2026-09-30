"""第四阶段：从已合并的工作簿生成 Shopify 和 WooCommerce 文件。"""
from pathlib import Path
from app.task_controller import checkpoint
from stages.data_processing import converter
from stages.data_processing.file_utils import load_product_rows


def convert_file(filename, log=print, stop_event=None, expand_count=3, expand_options=None):
    """转换单个合并表。

    expand_count 限定参与组合的选项个数（1..3）：超过这个数量的选项
    整列丢弃、不展开，用于压制维度爆炸。默认 3 即全部展开。界面上不再
    暴露这个开关——维度在第三阶段就该按勾选裁掉，这里的截断只作兜底。

    expand_options 是按选项名给出的白名单；非 None 时只有名字在其中的
    选项参与展开。None 表示不按名字限制。
    """
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
    if expand_count < 3:
        log(f'展开选项: 仅前 {expand_count} 个，其余不参与组合')
    if expand_options is not None:
        log(f'仅展开已勾选选项: {"、".join(expand_options) or "（无，退回第一维）"}')
    converter.styles_to_shopify(
        source, log, checkpoint=check,
        expand_count=expand_count, expand_options=expand_options,
    )
    log(f'转换导出完成: {source.parent}')
    return {'output': str(source.parent), 'count': len(rows), 'status': '完成'}
