"""读取文件夹中的商品中间表，执行字段规范与合并。"""

from app.task_controller import checkpoint
from dataclasses import dataclass
from pathlib import Path

from stages.data_processing import config
from stages.data_processing.file_utils import list_record_files, load_product_rows, save_xlsx
from stages.data_processing.record_import import convert_record_file
from stages.data_processing.field_rules.title import normalize_title, TitleValidationError
from stages.data_processing.field_rules.name import extract_brand, normalize_name, NameValidationError, UnsupportedNameCharacters
from stages.data_processing.field_rules.price1 import normalize_price1, PriceValidationError
from stages.data_processing.field_rules.price2 import normalize_price2
from stages.data_processing.field_rules.src_links import normalize_src_links, parse_src_links


@dataclass
class ProcessReport:
    input_files: int
    product_rows: int
    output_dir: Path
    merged_file: Path


def process_folder(folder, log=print, stop_event=None, keep_options=None):
    """第三阶段主入口：JSON 转中间表 → 逐字段规范 → 合并写表。

    keep_options 是页面上勾选参与转换的选项名集合；None 表示不限制。
    它只作用于 JSON → 中间表这一步：**未勾选的选项维度不会被编码进
    styles1**，因此第四阶段不会展开它，合并表与最终产物的行数都真正减少。
    已生成的 xlsx 是普通中间表，勾选不会再去改动它。

    策略是「先全量校验、后统一写入」：任何一行字段规范失败都只累积错误，
    待全部文件处理完后若发现错误，直接抛出且不写合并表、不执行转换，
    避免格式问题导致半成品结果。错误信息带真实 Excel 行号。
    全程在关键步骤前后调用 check() 响应停止信号。
    返回 ProcessReport（输入文件数、合并行数、输出目录、合并表路径）。
    """
    def check():
        """步骤边界的停止检查：未提供 stop_event 时为空操作。"""
        if stop_event is not None:
            checkpoint(stop_event)
    check()
    raw = str(folder).strip()
    if not raw:
        raise ValueError("请选择数据文件夹")
    source = Path(raw).expanduser().resolve()
    if not source.is_dir():
        raise ValueError(f"文件夹不存在: {source}")
    check()
    # 先把第二阶段的商品记录 JSON 转成同目录的 {域名}商品.xlsx，再走原有的
    # 中间表流程。每次运行都无条件重转：生成的 xlsx 能通过下面全部过滤条件，
    # 一旦被重新采集的 JSON 甩在后面就会被当成有效输入消费、把过期数据混进
    # 合并表，重转是消除该竞态的唯一可靠办法。整批编码失败时不留下半成品。
    records = list_record_files(source)
    if records:
        if keep_options is not None:
            log(f"参与转换的选项: {'、'.join(keep_options) or '（无，仅保留无选项商品）'}")
        for record_path in records:
            check()
            target = convert_record_file(record_path, log=log, keep_options=keep_options)
            log(f"已生成中间表: {target.name}")
    check()
    files = []
    for path in sorted(source.iterdir(), key=lambda p: p.name.casefold()):
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            continue
        if (path.name.startswith(("~", ".")) or path.stem.endswith("_合并")
                or ".recovery-" in path.name or "商品" not in path.stem):
            log(f"跳过非商品中间表: {path.name}")
            continue
        files.append(path)
    if not files:
        raise ValueError("文件夹中没有可处理的商品中间表（应为 {域名}商品.xlsx）")
    new_brand = extract_brand(source.name)
    log(f'目标品牌: {new_brand.word}（{new_brand.domain}）')

    # 全部输入验证通过后才写结果，避免格式错误造成部分合并。
    rows = []
    errors = []
    skipped_count = 0
    for index, path in enumerate(files, 1):
        old_brand = extract_brand(path.stem)
        log(f'{path.name}: 品牌 {old_brand.word} → {new_brand.word}')
        check()
        current = load_product_rows(path, with_locations=True)
        for row_number, item in current:
            check()
            # 先检查名称，整条跳过不支持文字的商品，不再校验其余字段。
            original_name = item['name']
            try:
                normalized_name = normalize_name(original_name, old_brand, new_brand)
            except UnsupportedNameCharacters as exc:
                skipped_count += 1
                log(f'警告：跳过商品 {path.name} 第 {row_number} 行 name={original_name!r}: {exc}')
                continue
            except NameValidationError as exc:
                message = f'{path.name} 第 {row_number} 行 name={original_name!r}: {exc}'
                errors.append(message)
                log(message)
                continue
            for field, normalize in (
                ('title', normalize_title),
                ('name', lambda value: normalized_name),
            ):
                original = item[field]
                try:
                    item[field] = normalize(original)
                except (TitleValidationError, NameValidationError) as exc:
                    message = f'{path.name} 第 {row_number} 行 {field}={original!r}: {exc}'
                    errors.append(message)
                    log(message)
                    continue
                if item[field] != original:
                    log(f"{path.name} 第 {row_number} 行 {field}: {original!r} → {item[field]!r}")
            price1_valid = False
            for field in ('price1', 'price2'):
                original = item[field]
                try:
                    if field == 'price1':
                        item[field] = normalize_price1(original)
                        price1_valid = True
                    elif price1_valid:
                        item[field] = normalize_price2(original, item['price1'])
                    else:
                        # price1 错误已阻止整批输出，仅检查 price2 本身是否合法。
                        normalize_price1(original)
                        continue
                except PriceValidationError as exc:
                    message = f'{path.name} 第 {row_number} 行 {field}={original!r}: {exc}'
                    errors.append(message)
                    log(message)
                    continue
                log(f"{path.name} 第 {row_number} 行 {field}: {original!r} → {item[field]:.2f}")
            original_images = item['src_links']
            item['src_links'] = normalize_src_links(original_images, item['styles1'])
            before_count = len(parse_src_links(original_images)[0])
            after_count = len(parse_src_links(item['src_links'])[0])
            log(f'{path.name} 第 {row_number} 行 src_links: {before_count} 张 → {after_count} 张')
            rows.append(item)
        log(f"[{index}/{len(files)}] {path.name}: {len(current)} 条")
    if skipped_count:
        log(f'警告：本次因名称含非英文文字或未支持字符跳过 {skipped_count} 个商品，保留 {len(rows)} 个商品')
    if errors:
        raise ValueError(f'字段校验失败，共 {len(errors)} 项；本次未写合并表或执行转换。\n' + '\n'.join(errors[:10]))
    if not rows:
        raise ValueError("没有可合并的商品（输入为空或商品已全部跳过）")

    output = source / config.RESULT_DIRNAME
    merged = output / f"{source.name}_合并.xlsx"
    check()
    save_xlsx(rows, merged)
    log(f"合并完成: {merged}（{len(rows)} 条）")
    log(f"规范与合并完成: {output}")
    return ProcessReport(len(files), len(rows), output, merged)
