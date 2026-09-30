from common import REFERENCE, payload
import unittest
from unittest.mock import Mock, patch

from standard_collection_process.config import MAX_STYLES_LENGTH
from standard_collection_process.crawler import browser, links
from standard_collection_process.processing import currency, html_cleaner, product, variants


def long_option_payload(count, value_size=200):
    """构造只含单个 Color 选项、且选项值很长的商品，用于触发 styles1 长度上限。"""
    data = payload()
    value = "V" * value_size
    data["product"]["options"] = [
        {"name": "Color", "position": 1, "values": [f"{value}{i:04d}" for i in range(count)]}
    ]
    return data


class ProcessingTests(unittest.TestCase):
    def test_prices_have_consistent_units(self):
        for value in (100, "100"):
            self.assertEqual(currency.format_price(value), "1")
        for value in (100.0, "100.00"):
            self.assertEqual(currency.format_price(value), "100")
        self.assertEqual(currency.format_price("12.50"), "12.5")
        self.assertEqual(currency.format_price(None), "")
        self.assertEqual(currency.format_price("0"), "0")

    def test_exchange_conversion(self):
        self.assertEqual(currency.format_price("100", {"EUR": 2}, "EUR"), "50")
        self.assertEqual(currency.format_price("100.00", {"EUR": 2}, "EUR"), "50")

    def test_product_matches_pre_split_reference(self):
        row = product.parse_product(payload(), "Category", "https://shop.test/products/fixture")
        expected = dict(REFERENCE["product_rows"][0], details="<p>Comfort cotton<br></p>")
        self.assertEqual(row, expected)

    def test_dynamic_skips_match_reference_without_mutating_input(self):
        data = payload()
        row = product.parse_product(
            data,
            "Category",
            "https://shop.test/products/fixture",
            skip_positions=[1],
            raw_skip_spec="1,-1",
        )
        expected = dict(REFERENCE["skipped_row"], details="<p>Comfort cotton<br></p>")
        self.assertEqual(row, expected)
        self.assertEqual(data, payload())

    def test_synthetic_combos_and_fallbacks_are_preserved(self):
        text = variants.build_variant_combo(payload()["product"])
        self.assertEqual(text, REFERENCE["product_rows"][0]["styles1"])
        self.assertIn("Green&Size&S$20$20", text)
        self.assertEqual(
            variants.build_variant_combo(payload()["product"], skip_options=["Size", "Color"]),
            "$10$12",
        )

    def test_invalid_json_is_rejected(self):
        invalid = [None, [], {"error": "404"}, {"product": {}}, {"product": []}]
        for field, value in (
            ("id", 0),
            ("handle", None),
            ("title", ""),
            ("variants", []),
            ("images", "bad"),
        ):
            data = payload()
            data["product"][field] = value
            invalid.append(data)
        for value in (None, -1, "NaN", "Infinity", "bad"):
            data = payload()
            data["product"]["variants"][0]["price"] = value
            invalid.append(data)
        for data in invalid:
            with self.subTest(data=data), self.assertRaises(product.InvalidProductError):
                product.parse_product(data, "Category", "https://shop.test/products/a")

    def test_too_long_styles1_is_rejected(self):
        # 300 个长值 → styles1 约 63000 字符，远超阈值
        data = long_option_payload(300)
        with self.assertRaises(product.InvalidProductError) as caught:
            product.parse_product(data, "Category", "https://shop.test/products/a")
        self.assertIn(str(MAX_STYLES_LENGTH), str(caught.exception))

    def test_styles1_at_limit_is_accepted(self):
        """边界：151 个值（31866 字符）通过，再多一个值（32077）即被拒绝。"""
        row = product.parse_product(
            long_option_payload(151), "Category", "https://shop.test/products/a"
        )
        self.assertLessEqual(len(row["styles1"]), MAX_STYLES_LENGTH)
        self.assertGreater(len(row["styles1"]), MAX_STYLES_LENGTH - 300)
        with self.assertRaises(product.InvalidProductError):
            product.parse_product(
                long_option_payload(152), "Category", "https://shop.test/products/a"
            )

    def test_handle_mode_has_no_dom_fallback(self):
        tab = Mock()
        tab.url = "https://shop.test/collections/all"
        tab.wait.doc_loaded.return_value = True
        tab.run_js.return_value = (
            '<script>{"handle":"a"},{"handle":"a"}</script><a href="/products/b">B</a>'
        )
        self.assertEqual(links.collect_page_links(tab), ["https://shop.test/products/a"])
        tab.wait.doc_loaded.assert_called_once_with(timeout=30, raise_err=False)
        tab.run_js.assert_called_once_with("return document.documentElement.outerHTML;")
        tab.run_js.return_value = '<a href="/products/b">B</a>'
        self.assertEqual(links.collect_page_links(tab), [])

    def test_xpath_is_exclusive(self):
        tab = Mock()
        tab.wait.doc_loaded.return_value = True
        with patch.object(links, "extract_links_by_xpath", return_value=[]) as xpath:
            self.assertEqual(links.collect_page_links(tab, xpath="//a/@href"), [])
        xpath.assert_called_once_with(tab, "//a/@href")
        tab.run_js.assert_not_called()

    def test_timeout_does_not_extract(self):
        tab = Mock()
        tab.wait.doc_loaded.return_value = False
        with self.assertRaises(TimeoutError):
            links.collect_page_links(tab)
        tab.run_js.assert_not_called()

    def test_fixed_proxy_and_image_blocking(self):
        with patch.object(browser, "ChromiumOptions") as options, patch.object(browser, "Chromium"):
            browser.create_browser(log=lambda text: None)
        options.return_value.no_imgs.assert_called_once_with(True)
        options.return_value.set_load_mode.assert_called_once_with("normal")
        options.return_value.set_argument.assert_called_once_with(
            "--proxy-server=http://127.0.0.1:7897"
        )

    def test_html_handlers_have_same_behavior(self):
        self.assertEqual(
            html_cleaner.clean_body_html('<p><a href="x">A</a><br/><img src="a.jpg"/></p>'),
            "<p>A<br></p>",
        )
