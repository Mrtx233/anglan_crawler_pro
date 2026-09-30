from pathlib import Path
from ..config import EXCLUDE_NAME_KEYWORDS
from ..processing.domain import _domain_label, replace_source_domain
from ..processing.exclude import split_by_keywords
from ..resources.columns import CSV_FIELDS
from .workbooks import load_xlsx


def read_folder_products(folder, log=print):
    """读取所选目录直属的来源商品表，排除本程序输出及 Excel 锁文件。

    来源表名是来源站域名，文件夹名开头是目标站域名；读取时把 name / details 里
    出现的来源站主标签替换为目标站主标签。name 命中 EXCLUDE_NAME_KEYWORDS 的商品整行剔除。
    """
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError("请选择存在的商品数据文件夹")
    target_label = _domain_label(folder.name)
    sources = sorted(
        (
            path
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.lower() == ".xlsx"
            and not path.name.startswith("~$")
            and not path.stem.endswith("_合并")
            and "_shopify_" not in path.stem.lower()
        ),
        key=lambda path: path.name.lower(),
    )
    if not sources:
        raise ValueError("文件夹中没有可合并的来源商品 XLSX")
    if not target_label:
        log(f"提示: 文件夹名 {folder.name} 未识别出目标域名，跳过域名替换")
    rows = []
    excluded_total = 0
    for source in sources:
        headers, records = load_xlsx(source)
        missing = [name for name in CSV_FIELDS if name not in headers]
        if missing:
            raise ValueError(f"{source.name} 不是完整商品数据表，缺少列：{', '.join(missing)}")
        records = [
            row for row in records if any(row.get(name) not in (None, "") for name in CSV_FIELDS)
        ]
        records, excluded = split_by_keywords(records, EXCLUDE_NAME_KEYWORDS)
        excluded_total += len(excluded)
        source_label = _domain_label(source.stem)
        records, replaced = replace_source_domain(records, source_label, target_label)
        rows.extend(records)
        log(f"读取 {source.name}：{len(records)} 条")
        for name, hits in excluded:
            log(f"    排除: 命中 {'、'.join(hits)} → {name}")
        if replaced:
            log(f"域名替换: {source_label} → {target_label}，命中 {replaced} 处")
    if excluded_total:
        log(f"共排除 {excluded_total} 条商品（关键词：{'、'.join(EXCLUDE_NAME_KEYWORDS)}）")
    if not rows:
        raise ValueError("来源商品 XLSX 均无有效数据")
    return folder / f"{folder.name}_合并.xlsx", rows
