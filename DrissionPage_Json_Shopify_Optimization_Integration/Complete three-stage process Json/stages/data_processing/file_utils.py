"""商品中间表读取与原子保存。"""

import json
import os
import tempfile
from pathlib import Path

import openpyxl

from stages.data_processing import config

# 第二阶段的商品记录后缀（{域名}商品.json）
RECORD_FILE_SUFFIX = ".json"


def is_record_file(path):
    """判断是否为第二阶段的商品记录 JSON。

    与商品中间表的命名约定一致：以 商品.json 结尾，且不是临时文件、
    历史合并表或 recovery 文件。
    """
    path = Path(path)
    if not path.is_file() or path.suffix.lower() != RECORD_FILE_SUFFIX:
        return False
    if path.name.startswith(("~", ".")):
        return False
    if path.stem.endswith("_合并") or ".recovery-" in path.name:
        return False
    return path.stem.endswith("商品")


def list_record_files(folder):
    """按文件名排序返回目录下所有商品记录 JSON。"""
    source = Path(folder)
    if not source.is_dir():
        return []
    return sorted(
        (p for p in source.iterdir() if is_record_file(p)),
        key=lambda p: p.name.casefold(),
    )


def read_record_products(filepath):
    """读取单个商品记录 JSON，返回商品字典列表。

    兼容商品采集阶段的 {"domain": ..., "products": [...]} 信封，也接受
    裸数组。格式非法或缺少 products 时抛 ValueError，由调用方决定
    是跳过该文件还是整体中止。
    """
    path = Path(filepath)
    try:
        with open(path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
    except (ValueError, OSError) as exc:
        raise ValueError(f"读取商品记录失败: {path.name}（{exc}）") from exc
    if isinstance(loaded, dict):
        products = loaded.get("products")
    elif isinstance(loaded, list):
        products = loaded
    else:
        raise ValueError(f"商品记录格式错误（应为对象或数组）: {path.name}")
    if not isinstance(products, list):
        raise ValueError(f"商品记录缺少 products 数组: {path.name}")
    return [item for item in products if isinstance(item, dict)]


def summarize_records(filepath):
    """汇总单个商品记录 JSON：商品数与出现过的选项名。

    选项名做**去重**收集——同名选项跨商品合并，因此返回的是「这个文件里
    出现过哪些选项」，而不是各商品值数之和。只报名字，不统计取值，
    因为这里关心的是「这批数据有哪些维度可供展开」，不是规模。
    返回值形如：
        {"name": "x.com商品.json", "products": 65,
         "options": ["Color", "Size"]}
    选项名保持首次出现的顺序；排序稳定的目的是让多次扫描结果可比对。
    """
    products = read_record_products(filepath)
    option_names = {}
    for product in products:
        options = product.get("options")
        if not isinstance(options, list):
            continue
        for option in options:
            if not isinstance(option, dict):
                continue
            name = str(option.get("name") or "").strip()
            if name:
                option_names[name] = None
    return {
        "name": Path(filepath).name,
        "path": str(filepath),
        "products": len(products),
        "options": list(option_names),
    }


def scan_record_folder(folder):
    """扫描目录下全部商品记录 JSON，返回逐个文件的汇总列表。

    单个文件解析失败只记入 errors，不中断整次扫描——目录里混着别的
    JSON 或半截文件时，仍然能看到其余文件的结果。
    """
    summaries = []
    errors = []
    for path in list_record_files(folder):
        try:
            summaries.append(summarize_records(path))
        except ValueError as exc:
            errors.append(str(exc))
    return summaries, errors


def save_xlsx(rows, output_path):
    """原子写入 XLSX：先写同目录临时文件，成功后 os.replace 覆盖目标。

    签名与原实现保持一致，所有调用方无需改动。
    失败时清理临时文件并抛出异常；原文件不会被破坏。
    """
    output_path = Path(output_path)
    target_dir = output_path.parent if str(output_path.parent) else Path(".")
    target_dir.mkdir(parents=True, exist_ok=True)

    fd, tmp_path = tempfile.mkstemp(
        prefix=f".{output_path.stem}_tmp_",
        suffix=".xlsx",
        dir=str(target_dir),
    )
    os.close(fd)
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(config.CSV_FIELDS)
        for row in rows:
            ws.append([row.get(f, "") for f in config.CSV_FIELDS])
            for field in ('price1', 'price2'):
                ws.cell(ws.max_row, config.CSV_FIELDS.index(field) + 1).number_format = '0.00'
        wb.save(tmp_path)
        wb.close()
        os.replace(tmp_path, str(output_path))
    except Exception:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise


def load_product_rows(filepath, *, with_locations=False):
    """按列名读取中间表，保留 0/False，跳过全空行。"""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    try:
        rows = wb.active.iter_rows(values_only=True)
        headers = [str(v).strip() if v is not None else "" for v in next(rows, ())]
        required = [field for field in config.CSV_FIELDS if field != "link-href"]
        missing = [field for field in required if field not in headers]
        if missing:
            raise ValueError(f"{Path(filepath).name} 缺少字段: {', '.join(missing)}")
        if len([h for h in headers if h]) != len(set(h for h in headers if h)):
            raise ValueError(f"{Path(filepath).name} 存在重复列名")
        result = []
        for row_number, values in enumerate(rows, 2):
            if all(value is None or value == "" for value in values):
                continue
            row = dict(zip(headers, values))
            item = {field: "" if row.get(field) is None else row[field]
                    for field in config.CSV_FIELDS}
            result.append((row_number, item) if with_locations else item)
        return result
    finally:
        wb.close()
