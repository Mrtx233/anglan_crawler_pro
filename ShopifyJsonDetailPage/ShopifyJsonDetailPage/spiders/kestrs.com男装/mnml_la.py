import json
from pathlib import Path

import scrapy

from ShopifyJsonDetailPage.items import ShopifyJsonItem
from ShopifyJsonDetailPage.utils.json_variants import (
    build_first_image_srcs,
    build_variant_combo,
    clean_body_html,
    fetch_exchange_rates,
    find_fallback_variant,
    group_targets_by_json_url,
    _fmt_price,
    _fmt_price2,
    _rates_cache,
)
from ShopifyJsonDetailPage.utils.link_collector import collect_links
from ShopifyJsonDetailPage.utils.spider_helpers import (
    get_spider_category,
    load_cached_links,
    load_done_urls,
    load_targets,
)


PROJECT_ROOT = Path(__file__).resolve().parents[3]

# ── 模式选择：二选一 ──
INPUT_FILE = None  # PROJECT_ROOT / "01xlsx/xxx.xlsx"

# PAGES 模式：DrissionPage 自动采集 collection 链接
# mnml.la 男装
XPATH_LINK = "//div[@id='products']/div[@id='productGrid']/div/a/@href"
MAX_PAGES = 5
PAGES = [
    ("Tops", "https://mnml.la/collections/tops"),
    ("Outerwear", "https://mnml.la/collections/outerwear"),
    ("Loungewear", "https://mnml.la/collections/loungewear"),
    ("Denim", "https://mnml.la/collections/denim"),
    ("Bottoms", "https://mnml.la/collections/bottoms"),
    ("Shorts", "https://mnml.la/collections/shorts"),
    ("New Arrivals", "https://mnml.la/collections/new-arrivals"),
    ("Best Sellers", "https://mnml.la/collections/best-sellers"),
    ("Back In Stock", "https://mnml.la/collections/back-in-stock"),
]

TEST_MODE = False
TEST_COUNT = 10
SKIP_POSITIONS = None
SKIP_OPTIONS = None
FORCE_RECOLLECT = False  # True = 强制重新采集链接，忽略缓存

fetch_exchange_rates()
_category = get_spider_category(__file__)


class MnmlLaSpider(scrapy.Spider):
    name = "mnml_la"
    allowed_domains = ["mnml.la"]
    _category = _category

    def _build_requests(self):
        if PAGES:
            if _category:
                output_dir = Path(__file__).resolve().parent.parent.parent / "output_details" / _category / self.name
            else:
                output_dir = Path(__file__).resolve().parent.parent.parent / "output_details" / self.name
            links_cache = output_dir / f"{self.name}_详细链接.xlsx"

            if links_cache.is_file() and not FORCE_RECOLLECT:
                targets = load_cached_links(str(links_cache))
                print(f"[{self.name}] 使用缓存链接: {len(targets)} 条 ← {links_cache}")
            else:
                output_dir.mkdir(parents=True, exist_ok=True)
                targets = collect_links(
                    pages=PAGES,
                    xpath_link=XPATH_LINK,
                    spider_name=self.name,
                    output_dir=str(output_dir),
                    max_pages=MAX_PAGES,
                )
                print(f"[{self.name}] PAGES 模式采集完成: {len(targets)} 条链接")
        elif INPUT_FILE:
            targets = load_targets(INPUT_FILE)
            print(f"[{self.name}] 读取 {len(targets)} 条 URL ← {INPUT_FILE}")
        else:
            print(f"[{self.name}] 未配置 PAGES 或 INPUT_FILE")
            return

        if TEST_MODE:
            targets = targets[:TEST_COUNT]

        groups, dupes = group_targets_by_json_url(targets)

        done_urls = load_done_urls(self.name, category=_category)
        skipped = 0
        pending_groups = {}
        for json_url, info in groups.items():
            if info["original_url"] in done_urls:
                skipped += 1
            else:
                pending_groups[json_url] = info

        self.total_items = len(pending_groups)
        print(
            f"[{self.name}] {len(groups)} 个不重复商品 "
            f"(共 {len(targets)} 行，跳过 {dupes} 条重复"
            f"{f'，已采集 {skipped} 条' if skipped else ''})"
        )
        if skipped:
            print(f"[{self.name}] 本次待采集: {self.total_items} 条")

        for json_url, info in pending_groups.items():
            yield scrapy.Request(
                json_url,
                callback=self.parse_json,
                headers={"Accept": "application/json,text/plain,*/*"},
                dont_filter=True,
                meta={
                    "table_title": info["title"],
                    "original_url": info["original_url"],
                    "json_url": json_url,
                },
            )

    def start_requests(self):
        yield from self._build_requests()

    async def start(self):
        for req in self._build_requests():
            yield req

    def parse_json(self, response):
        data = json.loads(response.text)
        product = data.get("product", {})
        variant = find_fallback_variant(product)

        item = ShopifyJsonItem()
        item["title"] = response.meta["table_title"]
        item["name"] = product.get("title", "")
        item["price1"] = _fmt_price(
            variant.get("price", ""),
            rates=_rates_cache,
            currency=variant.get("price_currency", ""),
        )
        item["price2"] = _fmt_price2(variant, rates=_rates_cache)
        item["styles1"] = build_variant_combo(
            product, skip_positions=SKIP_POSITIONS,
            skip_options=SKIP_OPTIONS, rates=_rates_cache,
        )
        item["styles2"] = ""
        item["styles3"] = ""
        item["src_links"] = build_first_image_srcs(
            product, limit=5, skip_positions=SKIP_POSITIONS,
        )
        item["link-href"] = response.meta["original_url"]
        item["details"] = clean_body_html(product.get("body_html", ""))
        yield item
