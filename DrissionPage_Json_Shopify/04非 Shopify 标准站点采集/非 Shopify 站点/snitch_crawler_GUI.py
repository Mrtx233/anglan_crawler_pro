# -*- coding: utf-8 -*-
"""
snitch.com 单站采集工具（DrissionPage + XPath）
- 输入: 01INPUT_XLSX 下 snitch.com.xlsx（第1列=标题/分类, 第2列=商品详情页URL）
- 输出: 02OUTPUT_XLSX 对应目录下同名 xlsx，含原始 9 字段
- 仅采集: 不做价格匹配、不做 WooCommerce 转换
- 断点续采: 输出文件已存在的 link-href 自动跳过；每采集 1 条即保存
"""
import os
import json
import random
import re
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

import openpyxl

from DrissionPage import Chromium, ChromiumOptions

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # 使上级目录 non_shopify_utils 可导入

from non_shopify_utils import (
    clean_body_html,
    fetch_exchange_rates,
    format_image_url,
    format_price,
)

DEFAULT_INPUT = r"D:\C_code\anglan_crawler_pro\DrissionPage_Json_Shopify\01INPUT_XLSX\A0812\周晓东\男装0829\snitch.com.xlsx"

CSV_FIELDS = [
    "title", "name", "price1", "price2",
    "styles1", "styles2", "styles3",
    "src_links", "link-href", "details",
]

# 非商品图（logo/促销/标签）关键字
IMAGE_EXCLUDE_KEYWORDS = ("snitch_logo", "offers.gif", "chemical-free-tag", "favicon")

SIZE_XPATH = '//h1[normalize-space(text())="SIZES"]/following-sibling::div[1]//span/div/span'
PRICE_XPATH = '//h1[contains(@class,"flex-1")]/following-sibling::div//p'
DESC_XPATH = '//div[contains(@class,"Collapsible__contentInner")]'
LDJSON_XPATH = '//script[@type="application/ld+json"]'


def to_output_path(input_path):
    out = input_path.replace("01INPUT_XLSX", "02OUTPUT_XLSX")
    return out


def load_tasks(input_path):
    wb = openpyxl.load_workbook(input_path, read_only=True)
    ws = wb.active
    tasks = []
    for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        title = str(row[0]).strip() if row[0] else ""
        url = str(row[1]).strip() if row[1] else ""
        if url.startswith("http"):
            tasks.append((title, url))
    wb.close()
    return tasks


def load_done_links(output_path):
    if not os.path.exists(output_path):
        return set(), []
    wb = openpyxl.load_workbook(output_path, read_only=True)
    ws = wb.active
    rows = []
    done = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[8]:
            done.add(str(row[8]))
            rows.append({f: (row[i] or "") for i, f in enumerate(CSV_FIELDS)})
    wb.close()
    return done, rows


def save_xlsx(rows, output_path):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(CSV_FIELDS)
    for row in rows:
        ws.append([row.get(f, "") for f in CSV_FIELDS])
    wb.save(output_path)


def parse_ldjson(tab):
    """XPath 提取 ld+json，返回商品字典或 {}"""
    for el in tab.eles(f"xpath:{LDJSON_XPATH}"):
        try:
            # DrissionPage 的 .text 取不到 script 内容，须用 inner_html
            data = json.loads(el.inner_html)
        except (ValueError, TypeError):
            continue
        if data.get("@type") == "Product":
            return data
    return {}


def _extract_flight_array(html, key):
    """从 Next.js flight 数据（self.__next_f.push）中取 key 对应的 JSON 数组

    兼容转义（\\"key\\":）与未转义（"key":）两种形态，返回 list 或 None
    """
    for pat in (f"\\\"{key}\\\":", f'"{key}":'):
        i = html.find(pat)
        if i == -1:
            continue
        start = html.find("[", i + len(pat))
        if start == -1:
            continue
        depth = 0
        for j in range(start, len(html)):
            c = html[j]
            if c == "[":
                depth += 1
            elif c == "]":
                depth -= 1
                if depth == 0:
                    raw = html[start:j + 1].replace('\\"', '"')
                    try:
                        parsed = json.loads(raw)
                        return parsed if isinstance(parsed, list) else None
                    except ValueError:
                        return None
    return None


def build_color_image_map(html):
    """解析 flight 数据的 colors / color_variants_ids / color_variants

    颜色与缩略图不能按数组顺序直接对齐（两数组顺序可能不同），
    必须走 颜色 → 同位置 product_id → 按 id 查 preview_image 的链路。
    返回 (颜色->缩略图URL 映射, 商品id->缩略图URL 映射)
    """
    colors_arr = _extract_flight_array(html, "colors") or []
    ids_arr = _extract_flight_array(html, "color_variants_ids") or []
    variants = _extract_flight_array(html, "color_variants") or []

    def _as_int(v):
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    img_by_id = {}
    for v in variants:
        if not isinstance(v, dict):
            continue
        pid = _as_int(v.get("shopify_product_id"))
        preview = str(v.get("preview_image", "") or "").split("?")[0]
        if pid and preview:
            img = format_image_url(preview)
            if img:
                img_by_id[pid] = img

    color_map = {}
    for i, c in enumerate(colors_arr):
        if i >= len(ids_arr):
            break
        img = img_by_id.get(_as_int(ids_arr[i]), "")
        if img:
            color_map[str(c).strip()] = img
    return color_map, img_by_id


def extract_product(tab, rates):
    ld = parse_ldjson(tab)

    # 标题: ld+json -> DOM h1
    name = ld.get("name", "")
    if not name:
        el = tab.ele('xpath://h1[contains(@class,"flex-1")]', timeout=3)
        name = el.text.strip() if el else ""

    # 售价: 优先 DOM 价格元素（页面首位），回退 ld+json offers
    price_raw = ""
    currency = ld.get("offers", {}).get("priceCurrency", "INR")
    el = tab.ele(f"xpath:{PRICE_XPATH}", timeout=3)
    if el:
        m = re.search(r"[\d,]+(?:\.\d+)?", el.text or "")
        if m:
            price_raw = m.group().replace(",", "")
    if not price_raw:
        price_raw = str(ld.get("offers", {}).get("price", ""))
    # 带上小数位，避免 format_price 按"分"除以 100
    if price_raw and "." not in price_raw:
        price_raw = f"{price_raw}.00"
    price1 = format_price(price_raw, rates=rates, currency=currency) if price_raw else ""
    # snitch 页面无划线价/MRP，price2 与 price1 相同
    price2 = price1

    # 尺码: DOM 按钮（顺序正确） -> ld+json size 字符串
    sizes = []
    for el in tab.eles(f"xpath:{SIZE_XPATH}"):
        t = (el.text or "").strip()
        if t and t not in sizes:
            sizes.append(t)
    if not sizes and ld.get("size"):
        sizes = [s.strip() for s in str(ld["size"]).split(",") if s.strip()]

    # 颜色: ld+json（多色商品是逗号串，需拆开做 颜色×尺码 笛卡尔积）
    color_raw = str(ld.get("color", "") or "").strip()
    colors = [c.strip() for c in color_raw.split(",") if c.strip()] if color_raw else []

    # 图片: ld+json image 数组，过滤 logo/促销/标签图，加 _600x600 后缀
    images = []
    for url in ld.get("image", []) or []:
        low = url.lower()
        if any(k in low for k in IMAGE_EXCLUDE_KEYWORDS):
            continue
        formatted = format_image_url(url)
        if formatted and formatted not in images:
            images.append(formatted)
    src_links = "#".join(images)

    # styles1: 颜色×尺码 笛卡尔积，每个颜色用自己的缩略图（snitch 各色同价）
    # 兜底链: 该色 preview_image -> 当前商品自己的 preview_image -> images[0]
    color_img_map, img_by_id = build_color_image_map(tab.html)
    current_pid = None
    for part in tab.url.rstrip("/").split("/"):
        if part.isdigit() and len(part) >= 10:
            current_pid = int(part)
    own_preview = img_by_id.get(current_pid, "")
    segments = []
    for color in colors:
        color_img = color_img_map.get(color) or own_preview or (images[0] if images else "")
        for size in sizes:
            seg = f"{color}&Size&{size}${price1}${price2}"
            if color_img:
                seg = f"{seg}@{color_img}"
            segments.append(seg)
    styles1 = f"Color#{'#'.join(segments)}" if colors and segments else ""

    # 详情: DOM 描述容器 innerHTML -> ld+json description
    details = ""
    el = tab.ele(f"xpath:{DESC_XPATH}", timeout=3)
    if el:
        details = clean_body_html(el.inner_html)
    elif ld.get("description"):
        details = clean_body_html(f"<p>{ld['description']}</p>")

    return {
        "name": name,
        "price1": price1,
        "price2": price2,
        "styles1": styles1,
        "src_links": src_links,
        "details": details,
    }


class SnitchGUI:
    def __init__(self, root):
        self.root = root
        root.title("snitch.com 采集工具（仅采集，无价格匹配/WP转换）")
        root.geometry("760x520")

        top = tk.Frame(root)
        top.pack(fill="x", padx=8, pady=6)
        tk.Label(top, text="输入文件:").pack(side="left")
        self.path_var = tk.StringVar(value=DEFAULT_INPUT)
        tk.Entry(top, textvariable=self.path_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(top, text="浏览...", command=self.pick_file).pack(side="left")

        btns = tk.Frame(root)
        btns.pack(fill="x", padx=8, pady=4)
        self.start_btn = tk.Button(btns, text="开始采集", width=14, command=self.start)
        self.start_btn.pack(side="left")
        self.stop_btn = tk.Button(btns, text="停止", width=10, state="disabled", command=self.stop)
        self.stop_btn.pack(side="left", padx=8)
        self.progress_var = tk.StringVar(value="待开始")
        tk.Label(btns, textvariable=self.progress_var).pack(side="left", padx=12)

        self.log_box = scrolledtext.ScrolledText(root, height=24)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=6)

        self.stop_event = threading.Event()
        self.worker = None

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.path_var.set(path)

    def log(self, msg):
        self.log_box.after(0, lambda: (self.log_box.insert("end", msg + "\n"), self.log_box.see("end")))

    def start(self):
        input_path = self.path_var.get().strip()
        if not os.path.exists(input_path):
            tk.messagebox.showerror("错误", f"输入文件不存在:\n{input_path}")
            return
        self.stop_event.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.worker = threading.Thread(target=self.run, args=(input_path,), daemon=True)
        self.worker.start()

    def stop(self):
        self.stop_event.set()
        self.log("→ 已请求停止，等待当前商品完成...")

    def run(self, input_path):
        try:
            self._run(input_path)
        except Exception as e:
            self.log(f"[异常] {e}")
        finally:
            self.root.after(0, lambda: (self.start_btn.config(state="normal"),
                                        self.stop_btn.config(state="disabled")))

    def _run(self, input_path):
        output_path = to_output_path(input_path)
        tasks = load_tasks(input_path)
        done, rows = load_done_links(output_path)
        self.log(f"输入: {input_path}")
        self.log(f"输出: {output_path}")
        self.log(f"任务 {len(tasks)} 条，已完成 {len(done)} 条，本次待采 {len(tasks) - len([u for _, u in tasks if u in done])} 条")

        rates = fetch_exchange_rates()
        if not rates:
            self.log("[警告] 汇率获取失败，价格将不做美元换算")

        chrome = create_browser(self.log)
        tab = chrome.latest_tab

        collected = 0
        try:
            for i, (title, url) in enumerate(tasks, 1):
                if self.stop_event.is_set():
                    self.log("→ 已停止")
                    break
                if url in done:
                    continue
                self.progress_var.set(f"{i}/{len(tasks)}")
                record = {"title": title, "link-href": url}
                ok = False
                for attempt in (1, 2):
                    try:
                        tab.get(url)
                        tab.ele(f"xpath:{LDJSON_XPATH}", timeout=20)
                        data = extract_product(tab, rates)
                        record.update(data)
                        ok = True
                        break
                    except Exception as e:
                        self.log(f"  [重试 {attempt}/2] {url}\n    {e}")
                if not ok:
                    self.log(f"  [失败] {url}")
                rows.append({f: record.get(f, "") for f in CSV_FIELDS})
                save_xlsx(rows, output_path)
                done.add(url)
                collected += 1
                if ok:
                    img_count = len([u for u in record["src_links"].split("#") if u])
                    self.log(f"  [{i}/{len(tasks)}] {record['name']} | ₹价格→${record['price1']} | {img_count}图 | {len(record['styles1'].split('#')) - 1}个规格组合")
        finally:
            chrome.quit()
        self.log(f"完成: 本次采集 {collected} 条，输出共 {len(rows)} 条 → {output_path}")


def create_browser(log):
    co = ChromiumOptions()
    co.incognito(on_off=True)
    co.set_argument("--proxy-server=http://127.0.0.1:7897")
    port = random.randint(9222, 9322)
    co.set_local_port(port)
    for attempt in range(3):
        try:
            chrome = Chromium(co)
            log(f"浏览器启动成功，端口: {port}")
            return chrome
        except Exception as e:
            log(f"端口 {port} 启动失败 (尝试 {attempt + 1}/3): {e}")
            port = random.randint(9222, 9322)
            co.set_local_port(port)
    raise RuntimeError("无法启动浏览器，已重试多次")


if __name__ == "__main__":
    root = tk.Tk()
    SnitchGUI(root)
    root.mainloop()
