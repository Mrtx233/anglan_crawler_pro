from __future__ import annotations

from pathlib import Path
import os
import sys
import traceback
from ..config import DELAY, PRODUCT_MAX_RETRIES, RETRY_BASE_DELAY, SAVE_EVERY
from ..crawler.browser import create_browser
from ..crawler.product_api import fetch_json
from ..processing.currency import fetch_exchange_rates
from ..processing.product import InvalidProductError, parse_product
from ..processing.images import parse_image_positions
from ..storage.paths import group_targets_by_json_url, map_to_output_path
from ..storage.workbooks import load_xlsx, read_targets_from_excel, save_xlsx


def run_product_task(context, config):
    folder = config.folder
    rates = fetch_exchange_rates(log=context.log)
    all_results = []
    saved_merge_count = 0
    merge_path = str(map_to_output_path(Path(folder)) / f"{Path(folder).name}_合并.xlsx")
    grand_success = 0
    grand_fail = 0
    file_count = len(config.files)
    for fi, file_info in enumerate(config.files):
        if context.stop_event.is_set():
            break
        context.state.active_file_index = fi
        context.emit("file_status", fi, "处理中", "running")
        fpath = file_info.path
        fname = os.path.basename(fpath)
        file_keep = parse_image_positions(file_info.keep)
        file_skip = parse_image_positions(file_info.skip)
        output_file = str(map_to_output_path(Path(fpath)))
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        context.log(f"\n{'=' * 50}")
        context.log(f"[{fi + 1}/{file_count}] {fname}")
        context.log(f"  输出: {output_file}")
        if file_skip:
            context.log(f"  跳过图片: {file_skip}")
        targets = read_targets_from_excel(fpath)
        json_url_groups, dup_count = group_targets_by_json_url(targets)
        total = len(json_url_groups)
        context.log(f"  共 {total} 个唯一商品（跳过 {dup_count} 个重复）")
        if total == 0:
            context.log("  无商品，跳过")
            context.emit("file_status", fi, "已跳过", "skipped")
            continue
        results = []
        if os.path.exists(output_file):
            _, results = load_xlsx(output_file)
        result_indexes = {
            str(row["link-href"]): index
            for index, row in enumerate(results)
            if row.get("link-href")
        }
        todo = {
            url: group
            for url, group in json_url_groups.items()
            if group["original_url"] not in result_indexes
        }
        context.log(f"  本次配置：跳过 {total - len(todo)} 个已完成商品，处理 {len(todo)} 个")
        if not todo:
            all_results.extend(results)
            context.emit("file_status", fi, "已完成", "success")
            continue
        context.emit("progress", 0, len(todo), f"[{fi + 1}/{file_count}] 正在启动浏览器 · {fname}")
        chrome = create_browser(log=context.log)
        success_count = 0
        saved_success_count = 0
        fail_count = 0
        try:
            tab = chrome.latest_tab
            for index, (json_url, group) in enumerate(todo.items(), 1):
                if context.stop_event.is_set():
                    break
                table_title = group["title"]
                original_url = group["original_url"]
                context.emit(
                    "progress",
                    index,
                    len(todo),
                    f"[{fi + 1}/{file_count}] {fname} · {table_title[:24]}",
                )
                context.log(f"  [{index}/{len(todo)}] {table_title}")
                retry_count = 0
                item = None
                while retry_count <= PRODUCT_MAX_RETRIES:
                    if context.stop_event.is_set():
                        break
                    status, data = fetch_json(tab, json_url, log=context.log)
                    if status == 200:
                        try:
                            item = parse_product(
                                data,
                                table_title,
                                original_url,
                                skip_positions=file_skip,
                                keep_positions=file_keep,
                                skip_options=config.skip_options,
                                rates=rates,
                            )
                        except InvalidProductError as exc:
                            context.log(f"    跳过当前商品: {exc}")
                        break
                    if status == 429:
                        retry_count += 1
                        if retry_count > PRODUCT_MAX_RETRIES:
                            context.log(f"    429 超过 {PRODUCT_MAX_RETRIES} 次，放弃")
                            break
                        delay = RETRY_BASE_DELAY * retry_count
                        context.log(f"    429 第{retry_count}次，{delay}秒后重试...")
                        context.stop_event.wait(delay)
                    else:
                        if status == 404:
                            context.log("    商品不存在 (status=404)，跳过当前商品")
                            break
                        retry_count += 1
                        if retry_count > PRODUCT_MAX_RETRIES:
                            context.log(
                                f"    失败 {PRODUCT_MAX_RETRIES} 次，放弃 (status={status})"
                            )
                            break
                        delay = 5 * retry_count
                        context.log(f"    失败 (status={status})，{delay}秒后重试...")
                        context.stop_event.wait(delay)
                    context.stop_event.wait(DELAY)
                if item:
                    if original_url in result_indexes:
                        results[result_indexes[original_url]] = item
                    else:
                        result_indexes[original_url] = len(results)
                        results.append(item)
                    success_count += 1
                    context.log(f"    成功: {item['name'][:40]}")
                elif not context.stop_event.is_set():
                    fail_count += 1
                    context.state.failed_product_count += 1
                if context.stop_event.is_set():
                    break
                if success_count - saved_success_count >= SAVE_EVERY:
                    context.checkpoints.save(output_file, save_xlsx, results, output_file)
                    saved_success_count = success_count
                    context.log(f"    [自动保存] {len(results)} 条")
                context.stop_event.wait(DELAY)
        finally:
            all_results.extend(results)
            if success_count != saved_success_count:
                context.checkpoints.pending[output_file] = (save_xlsx, (results, output_file))
            context.state.last_merge_path = merge_path if all_results else ""
            if (context.stop_event.is_set() or sys.exc_info()[0] is not None) and len(
                all_results
            ) != saved_merge_count:
                context.checkpoints.pending[merge_path] = (
                    save_xlsx,
                    (list(all_results), merge_path),
                )
            try:
                for path in (output_file, merge_path):
                    if path in context.checkpoints.pending:
                        writer, args = context.checkpoints.pending[path]
                        context.checkpoints.save(path, writer, *args)
                        if path == merge_path:
                            saved_merge_count = len(all_results)
                        context.log(f"  数据已保存: {path}")
            finally:
                context.log("  正在关闭浏览器...")
                try:
                    chrome.quit()
                except Exception:
                    context.log(traceback.format_exc())
        grand_success += success_count
        grand_fail += fail_count
        if context.stop_event.is_set():
            context.emit("file_status", fi, "已停止", "stopped")
            break
        if fail_count > 0 and success_count == 0:
            context.emit("file_status", fi, "采集失败", "failed")
        elif fail_count > 0:
            context.emit("file_status", fi, "部分失败", "partial")
        else:
            context.emit("file_status", fi, "已完成", "success")
    if context.stop_event.is_set() and context.state.active_file_index is not None:
        for pending_index in range(context.state.active_file_index + 1, len(config.files)):
            context.emit("file_status", pending_index, "未处理", "waiting")
    if all_results:
        context.merge_ready(merge_path)
        context.log(f"来源文件已保存，共 {len(all_results)} 条；请选文件夹合并并转换")
    final_text = f"总计：{grand_success} 成功，{grand_fail} 失败，{len(all_results)} 条数据"
    if context.stop_event.is_set():
        final_text = "任务已中断 · " + final_text
    context.emit("progress", grand_success + grand_fail, grand_success + grand_fail, final_text)
    context.log(f"\n{final_text}")
