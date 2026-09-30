# -*- coding: utf-8 -*-
"""teeshoppen.com 商品采集工具（Shopify Hydrogen + DrissionPage）。

标准 ``/products/*.json`` 和 ``*.js`` 会返回 Hydrogen HTML。商品数据取自
服务端渲染的 JSON-LD 与 DOM；输出保持主管线的 10 字段契约。
"""

import json
import os
import random
import tempfile
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import openpyxl
from DrissionPage import Chromium, ChromiumOptions
from lxml import etree, html as lxml_html

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))  # 使上级目录 non_shopify_utils 可导入

from non_shopify_utils import (
    clean_body_html,
    fetch_exchange_rates,
    format_image_url,
    format_price,
)


MODULE_DIR = Path(__file__).resolve().parent
SHOPIFY_ROOT = MODULE_DIR.parent.parent
DEFAULT_INPUT = str(SHOPIFY_ROOT / "01INPUT_XLSX" / "teeshoppen.com.xlsx")
PROXY_SERVER = "http://127.0.0.1:7897"


CSV_FIELDS = [
    "title", "name", "price1", "price2",
    "styles1", "styles2", "styles3",
    "src_links", "link-href", "details",
]


class ProductExtractionError(ValueError):
    """页面不是可保存的完整 TeeShoppen 商品时抛出。"""


class CurrencyConversionError(RuntimeError):
    """商品币种无法可靠换算为 USD 时抛出，并终止整批采集。"""


def parse_product_html(rendered_html, rates):
    source = str(rendered_html or "")
    lowered = source.lower()
    if "just a moment" in lowered or "challenges.cloudflare.com" in lowered:
        raise ProductExtractionError("页面被 Cloudflare 人机验证拦截，请检查代理出口")

    try:
        document = lxml_html.fromstring(source)
    except (etree.ParserError, ValueError) as exc:
        raise ProductExtractionError(f"页面 HTML 无法解析: {exc}") from exc

    product = _ldjson_product(document)
    name = str(product.get("name") or "").strip()
    if not name:
        headings = document.xpath(
            '//*[@data-wv-type="mp--title"]//h1 | '
            '//*[@data-wv-type="main-product"]//h1'
        )
        name = _node_text(headings[0]) if headings else ""

    offers = product.get("offers") or []
    if isinstance(offers, dict):
        offers = [offers]
    first_offer = next((offer for offer in offers if isinstance(offer, dict)), {})
    price_raw = str(first_offer.get("price") or "").strip()
    currency = str(first_offer.get("priceCurrency") or "USD").strip() or "USD"
    currency = currency.upper()
    if currency != "USD":
        rate = rates.get(currency) if isinstance(rates, dict) else None
        try:
            valid_rate = float(rate) > 0
        except (TypeError, ValueError):
            valid_rate = False
        if not valid_rate:
            raise CurrencyConversionError(
                f"汇率表缺少有效的 {currency}→USD 汇率，停止采集"
            )
    if price_raw and "." not in price_raw:
        price_raw = f"{price_raw}.00"
    price1 = format_price(price_raw, rates=rates, currency=currency) if price_raw else ""
    price2 = price1

    sizes = []
    selector_xpath = (
        '//*[@data-wv-type="mp--variant-selector"]'
        '//*[contains(concat(" ", normalize-space(@class), " "), " grid ")]'
        '//button'
    )
    for button in document.xpath(selector_xpath):
        size = _node_text(button)
        classes = set(str(button.get("class") or "").split())
        if size and "diagonal" not in classes and size not in sizes:
            sizes.append(size)
    if not sizes:
        sizes = _offer_sizes(offers)

    images = []
    for image in document.xpath('//*[@data-wv-type="mp--media"][1]//img[@src]'):
        clean_url = _clean_shopify_image_url(image.get("src"))
        formatted = format_image_url(clean_url)
        if formatted and formatted not in images:
            images.append(formatted)

    detail_nodes = document.xpath('//*[@data-wv-type="mp--collapsible-details"]')
    details = ""
    if detail_nodes:
        details = clean_body_html(_inner_html(detail_nodes[0]))
    elif product.get("description"):
        details = clean_body_html(f"<p>{product['description']}</p>")

    missing = [
        field
        for field, value in (
            ("name", name),
            ("price1", price1),
            ("sizes", sizes),
            ("src_links", images),
            ("details", details),
        )
        if not value
    ]
    if missing:
        raise ProductExtractionError(f"页面关键字段缺失: {', '.join(missing)}")

    first_image = images[0]
    segments = [
        f"{size}${price1}${price2}@{first_image}"
        for size in sizes
    ]
    return {
        "name": name,
        "price1": price1,
        "price2": price2,
        "styles1": f"Size#{'#'.join(segments)}",
        "src_links": "#".join(images),
        "details": details,
    }


def build_record(title, url, product):
    values = {
        "title": str(title or ""),
        "link-href": str(url or ""),
        "styles2": "",
        "styles3": "",
    }
    values.update(product or {})
    return {field: values.get(field, "") for field in CSV_FIELDS}


def _node_text(node):
    return " ".join(" ".join(node.itertext()).split())


def _inner_html(element):
    return "".join(
        etree.tostring(child, encoding="unicode", method="html")
        for child in element
    )


def _ldjson_product(document):
    for script in document.xpath('//script[@type="application/ld+json"]'):
        text = (script.text_content() or "").strip()
        if not text:
            continue
        try:
            payload = json.loads(text)
        except (TypeError, ValueError):
            continue
        nodes = list(payload) if isinstance(payload, list) else [payload]
        while nodes:
            node = nodes.pop(0)
            if not isinstance(node, dict):
                continue
            node_type = node.get("@type")
            node_types = node_type if isinstance(node_type, list) else [node_type]
            if "Product" in node_types:
                return node
            graph = node.get("@graph")
            if isinstance(graph, list):
                nodes.extend(graph)
            elif isinstance(graph, dict):
                nodes.append(graph)
    return {}


def _offer_sizes(offers):
    sizes = []
    for offer in offers:
        if not isinstance(offer, dict):
            continue
        if str(offer.get("availability") or "").rsplit("/", 1)[-1] != "InStock":
            continue
        query = dict(parse_qsl(urlsplit(str(offer.get("url") or "")).query))
        size = str(query.get("Size") or query.get("size") or "").strip()
        if size and size not in sizes:
            sizes.append(size)
    return sizes


def _clean_shopify_image_url(url):
    """删除 Shopify 动态裁切参数，保留版本号后再添加文件名尺寸标记。"""
    parsed = urlsplit(str(url or ""))
    transform_keys = {"width", "height", "crop"}
    query = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if key.lower() not in transform_keys
    ]
    return urlunsplit((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        urlencode(query),
        parsed.fragment,
    ))


def is_record_complete(record):
    required = (
        "name", "price1", "price2", "styles1",
        "src_links", "link-href", "details",
    )
    return all(str(record.get(field, "") or "").strip() for field in required)


def to_output_path(input_path):
    path = Path(input_path)
    parts = list(path.parts)
    try:
        input_index = parts.index("01INPUT_XLSX")
    except ValueError as exc:
        raise ValueError("输入文件必须位于 01INPUT_XLSX 目录内，避免覆盖原文件") from exc
    parts[input_index] = "02OUTPUT_XLSX"
    output = Path(*parts)
    if os.path.normcase(os.path.abspath(output)) == os.path.normcase(os.path.abspath(path)):
        raise ValueError("输出路径不能与输入路径相同")
    return str(output)


def upsert_row(rows, record):
    url = str(record.get("link-href", "") or "").strip()
    for index, existing in enumerate(rows):
        if str(existing.get("link-href", "") or "").strip() == url:
            rows[index] = record
            return
    rows.append(record)


def wait_for_product(tab, rates, expected_url=None, timeout=30, poll_interval=0.25):
    if expected_url and _product_location(tab.url) != _product_location(expected_url):
        raise ProductExtractionError(
            f"页面地址不匹配: 期望 {expected_url}，实际 {tab.url}"
        )
    deadline = time.monotonic() + timeout
    last_error = None
    while True:
        try:
            return parse_product_html(tab.html, rates)
        except ProductExtractionError as exc:
            last_error = exc
            if "拦截" in str(exc):
                raise
        if time.monotonic() >= deadline:
            raise ProductExtractionError(f"商品数据等待超时: {last_error}")
        time.sleep(poll_interval)


def _product_location(url):
    parsed = urlsplit(str(url or ""))
    host = (parsed.hostname or "").lower()
    if host.startswith("www."):
        host = host[4:]
    path = "/".join(part for part in parsed.path.split("/") if part)
    return host, f"/{path}" if path else "/"


def load_tasks(input_path):
    workbook = openpyxl.load_workbook(input_path, read_only=True)
    try:
        worksheet = workbook.active
        tasks = []
        for title, url in worksheet.iter_rows(
            min_row=2, min_col=1, max_col=2, values_only=True
        ):
            clean_title = str(title).strip() if title else ""
            clean_url = str(url).strip() if url else ""
            if clean_url.startswith(("http://", "https://")):
                tasks.append((clean_title, clean_url))
        return tasks
    finally:
        workbook.close()


def load_done_links(output_path):
    path = Path(output_path)
    if not path.exists():
        return set(), []
    workbook = openpyxl.load_workbook(path, read_only=True)
    try:
        worksheet = workbook.active
        header = [
            str(cell.value or "").strip()
            for cell in worksheet[1][:len(CSV_FIELDS)]
        ]
        if header != CSV_FIELDS:
            raise ValueError("现有输出文件表头与 10 字段契约不一致")
        rows = []
        done = set()
        for values in worksheet.iter_rows(
            min_row=2, min_col=1, max_col=len(CSV_FIELDS), values_only=True
        ):
            row = {
                field: (values[index] if values[index] is not None else "")
                for index, field in enumerate(CSV_FIELDS)
            }
            url = str(row["link-href"] or "").strip()
            if not url:
                continue
            rows.append(row)
            if is_record_complete(row):
                done.add(url)
        return done, rows
    finally:
        workbook.close()


def save_xlsx(rows, output_path):
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(CSV_FIELDS)
    for row in rows:
        worksheet.append([row.get(field, "") for field in CSV_FIELDS])
    temp_handle = tempfile.NamedTemporaryFile(
        dir=path.parent,
        prefix=f".{path.stem}-",
        suffix=".tmp.xlsx",
        delete=False,
    )
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    try:
        workbook.save(temp_path)
        os.replace(temp_path, path)
    finally:
        workbook.close()
        if temp_path.exists():
            temp_path.unlink()


def create_browser(log):
    options = ChromiumOptions()
    options.incognito(on_off=True)
    options.set_argument(f"--proxy-server={PROXY_SERVER}")
    port = random.randint(9222, 9322)
    options.set_local_port(port)
    for attempt in range(1, 4):
        try:
            chrome = Chromium(options)
            log(f"浏览器启动成功，端口: {port}")
            return chrome
        except Exception as exc:
            log(f"端口 {port} 启动失败（{attempt}/3）: {exc}")
            port = random.randint(9222, 9322)
            options.set_local_port(port)
    raise RuntimeError("无法启动浏览器，已重试 3 次")


class TeeShoppenGUI:
    def __init__(self, root):
        self.root = root
        root.title("teeshoppen.com 采集工具（仅采集）")
        root.geometry("780x540")

        top = tk.Frame(root)
        top.pack(fill="x", padx=8, pady=6)
        tk.Label(top, text="输入文件:").pack(side="left")
        self.path_var = tk.StringVar(value=DEFAULT_INPUT)
        tk.Entry(top, textvariable=self.path_var).pack(
            side="left", fill="x", expand=True, padx=6
        )
        tk.Button(top, text="浏览...", command=self.pick_file).pack(side="left")

        buttons = tk.Frame(root)
        buttons.pack(fill="x", padx=8, pady=4)
        self.start_btn = tk.Button(
            buttons, text="开始采集", width=14, command=self.start
        )
        self.start_btn.pack(side="left")
        self.stop_btn = tk.Button(
            buttons, text="停止", width=10, state="disabled", command=self.stop
        )
        self.stop_btn.pack(side="left", padx=8)
        self.progress_var = tk.StringVar(value="待开始")
        tk.Label(buttons, textvariable=self.progress_var).pack(side="left", padx=12)

        self.log_box = scrolledtext.ScrolledText(root, height=25)
        self.log_box.pack(fill="both", expand=True, padx=8, pady=6)
        self.stop_event = threading.Event()
        self.worker = None

    def pick_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.path_var.set(path)

    def log(self, message):
        def append():
            self.log_box.insert("end", str(message) + "\n")
            self.log_box.see("end")
        self.root.after(0, append)

    def set_progress(self, message):
        self.root.after(0, self.progress_var.set, message)

    def start(self):
        input_path = self.path_var.get().strip()
        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入文件不存在:\n{input_path}")
            return
        try:
            to_output_path(input_path)
        except ValueError as exc:
            messagebox.showerror("错误", str(exc))
            return
        self.stop_event.clear()
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.worker = threading.Thread(
            target=self.run, args=(input_path,), daemon=True
        )
        self.worker.start()

    def stop(self):
        self.stop_event.set()
        self.log("→ 已请求停止，等待当前商品完成...")

    def run(self, input_path):
        try:
            self._run(input_path)
        except Exception as exc:
            self.log(f"[异常] {exc}")
        finally:
            self.root.after(0, self.start_btn.config, {"state": "normal"})
            self.root.after(0, self.stop_btn.config, {"state": "disabled"})

    def _run(self, input_path):
        output_path = to_output_path(input_path)
        tasks = load_tasks(input_path)
        done, rows = load_done_links(output_path)
        pending = sum(1 for _, url in tasks if url not in done)
        self.log(f"输入: {input_path}")
        self.log(f"输出: {output_path}")
        self.log(f"任务 {len(tasks)} 条，完整记录 {len(done)} 条，本次待采 {pending} 条")

        rates = fetch_exchange_rates()
        if not rates:
            raise RuntimeError("汇率获取失败，停止采集，避免写入非 USD 价格")

        chrome = create_browser(self.log)
        tab = chrome.latest_tab
        processed = 0
        succeeded = 0
        try:
            for index, (title, url) in enumerate(tasks, 1):
                if self.stop_event.is_set():
                    self.log("→ 已停止")
                    break
                if url in done:
                    continue

                self.set_progress(f"{index}/{len(tasks)}")
                record = build_record(title, url, {})
                error = None
                for attempt in (1, 2):
                    try:
                        tab.get(url)
                        product = wait_for_product(
                            tab, rates, expected_url=url, timeout=30
                        )
                        record = build_record(title, url, product)
                        error = None
                        break
                    except CurrencyConversionError:
                        raise
                    except Exception as exc:
                        error = exc
                        self.log(f"  [重试 {attempt}/2] {url}\n    {exc}")

                upsert_row(rows, record)
                save_xlsx(rows, output_path)
                processed += 1
                if error is None and is_record_complete(record):
                    done.add(url)
                    succeeded += 1
                    image_count = len(record["src_links"].split("#"))
                    variant_count = max(0, len(record["styles1"].split("#")) - 1)
                    self.log(
                        f"  [{index}/{len(tasks)}] {record['name']} | "
                        f"${record['price1']} | {image_count}图 | "
                        f"{variant_count}个规格组合"
                    )
                else:
                    self.log(f"  [失败占位，后续会重采] {url}: {error}")
        finally:
            chrome.quit()

        self.log(
            f"完成: 本次处理 {processed} 条，成功 {succeeded} 条，"
            f"输出共 {len(rows)} 条 → {output_path}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    TeeShoppenGUI(root)
    root.mainloop()
