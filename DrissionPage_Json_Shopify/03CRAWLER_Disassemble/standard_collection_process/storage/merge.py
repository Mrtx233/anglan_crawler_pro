from pathlib import Path
from ..resources.columns import CSV_FIELDS
from .workbooks import load_xlsx


def read_folder_products(folder, log=print):
    """读取所选目录直属的来源商品表，排除本程序输出及 Excel 锁文件。"""
    folder = Path(folder)
    if not folder.is_dir():
        raise ValueError("请选择存在的商品数据文件夹")
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
    rows = []
    for source in sources:
        headers, records = load_xlsx(source)
        missing = [name for name in CSV_FIELDS if name not in headers]
        if missing:
            raise ValueError(f"{source.name} 不是完整商品数据表，缺少列：{', '.join(missing)}")
        records = [
            row for row in records if any(row.get(name) not in (None, "") for name in CSV_FIELDS)
        ]
        rows.extend(records)
        log(f"读取 {source.name}：{len(records)} 条")
    if not rows:
        raise ValueError("来源商品 XLSX 均无有效数据")
    return folder / f"{folder.name}_合并.xlsx", rows
