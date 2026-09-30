from common import payload
import unittest
from unittest.mock import patch
from standard_collection_process.processing import product, options, shopify, variants, html_cleaner


class BaselineRulesTests(unittest.TestCase):
    def test_option_names_and_escaped_segments_round_trip(self):
        for name, values, expected in [
            ("Product SIZE", ["Small", "Large"], "Size"),
            ("Choose", ["S", "M"], "Size"),
            ("Type", ["S", "M"], "Size"),
            ("TYPE", ["A", "B"], "Style"),
        ]:
            self.assertEqual(
                options.get_sorted_options(
                    {"options": [{"position": 1, "name": name, "values": values}]}
                )[0]["name"],
                expected,
            )
        value = r"Tie & Square#Set@Shop\Blue"
        name = "Finish & Material#@"
        opts = [{"name": "Color", "position": 1}, {"name": name, "position": 2}]
        segment = variants.build_segment(
            (value, value),
            opts,
            {"price": "19.99", "compare_at_price": "25.00"},
            {value: "https://a/img.jpg"},
        )
        group, segments = shopify._parse_style_group("Color#" + segment)
        parsed = shopify._parse_variant_segment(segments[0])
        self.assertEqual(parsed["option1_value"], value)
        self.assertEqual(parsed["option2_name"], name)
        self.assertEqual(parsed["option2_value"], value)
        self.assertEqual(parsed["image"], "https://a/img.jpg")

    def test_all_images_and_keep_priority(self):
        data = payload()
        data["product"]["images"] = [
            {"id": i, "position": i, "src": f"https://cdn.test/{i}.jpg"} for i in range(1, 9)
        ]
        row = product.parse_product(data, "C", "https://shop.test/products/a")
        self.assertEqual(len(row["src_links"].split("#")), 8)
        row = product.parse_product(data, "C", "url", keep_positions=[2, 8], skip_positions=[2])
        self.assertEqual(
            row["src_links"], "https://cdn.test/2_600x600.jpg#https://cdn.test/8_600x600.jpg"
        )
        self.assertEqual(
            product.parse_product(data, "C", "url", keep_positions=[])["src_links"], ""
        )

    def test_variant_images_appended_once(self):
        with patch.object(product, "build_first_image_srcs", return_value="a"), patch.object(
            product,
            "build_images_by_first_option",
            return_value={"Red": "b", "Blue": "b", "Green": "a"},
        ):
            self.assertEqual(product.parse_product(payload(), "C", "url")["src_links"], "a#b")

    def test_nested_empty_tags_removed(self):
        self.assertEqual(
            html_cleaner.clean_body_html(
                '<div><p><span> </span><img src="x"></p></div><p><a href="x">Text</a></p>'
            ),
            "<p>Text</p>",
        )
