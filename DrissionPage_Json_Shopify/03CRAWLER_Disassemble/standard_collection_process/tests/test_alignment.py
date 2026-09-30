from common import WorkspaceTest, payload
from pathlib import Path
from unittest.mock import patch
from standard_collection_process.pipeline import product_task, conversion_task
from standard_collection_process.pipeline.models import PipelineStage, ConversionTaskConfig
from standard_collection_process.processing.currency import format_price2
from standard_collection_process.processing.images import parse_image_positions
from standard_collection_process.processing.product import parse_product
from standard_collection_process.storage import workbooks


class AlignmentTests(WorkspaceTest):
    def test_zero_compare_price_falls_back_before_wp(self):
        for value in (0, "0", "0.00", " 0 "):
            self.assertEqual(format_price2({"price": "100.00", "compare_at_price": value}), "100")

    def test_shared_positions_and_negative_skip(self):
        self.assertEqual(parse_image_positions("1，－2, -3,bad"), [1, -2, -3])
        data = payload()
        row = parse_product(data, "C", "url", skip_positions=[-2])
        self.assertNotIn("2_600x600", row["src_links"])
        self.assertEqual(parse_image_positions("invalid"), [])
        self.assertIsNone(parse_image_positions(""))

    def test_conversion_reads_edited_source_files(self):
        url = "https://shop.test/products/a"
        with patch.object(product_task, "create_browser"), patch.object(
            product_task, "fetch_exchange_rates", return_value={}
        ), patch.object(
            product_task, "read_targets_from_excel", return_value=[("C", url)]
        ), patch.object(
            product_task, "fetch_json", return_value=(200, payload())
        ), patch.object(
            product_task, "DELAY", 0
        ):
            product_task.run_product_task(self.context, self.config)
        merge_path = Path(self.controller.state.last_merge_path)
        self.assertFalse(merge_path.exists())
        self.assertTrue((self.output / "shop.xlsx").exists())
        workbooks.save_xlsx([{"name": "external edit"}], self.output / "shop.xlsx")
        with patch.object(conversion_task, "convert_styles_file", return_value=merge_path):
            conversion_task.run_conversion_task(self.context, ConversionTaskConfig(self.output))
        self.assertEqual(workbooks.load_xlsx(merge_path)[1][0]["name"], "external edit")

    def test_failed_conversion_can_retry_folder(self):
        path = self.output / "category_合并.xlsx"
        workbooks.save_xlsx([{"name": "Retained"}], self.output / "shop.xlsx")
        with patch.object(
            conversion_task, "convert_styles_file", side_effect=RuntimeError("failed")
        ):
            self.controller.start(PipelineStage.CONVERTING, ConversionTaskConfig(self.output))
            self.wait_worker()
        self.assertTrue(path.exists())
        self.assertFalse(self.controller.shutdown_requested)
        with patch.object(conversion_task, "convert_styles_file", return_value=path):
            self.controller.start(PipelineStage.CONVERTING, ConversionTaskConfig(self.output))
            self.wait_worker()
        self.assertEqual(self.controller.state.stage, PipelineStage.COMPLETED)
        self.assertEqual(len(workbooks.load_xlsx(path)[1]), 1)

    def test_resume_uses_exact_original_url_even_without_name(self):
        url = "https://shop.test/products/a"
        for previous, count in ((url, 0), (url + ".json", 1)):
            workbooks.save_xlsx([{"link-href": previous}], self.output / "shop.xlsx")
            with patch.object(product_task, "create_browser"), patch.object(
                product_task, "fetch_exchange_rates", return_value={}
            ), patch.object(
                product_task, "read_targets_from_excel", return_value=[("C", url)]
            ), patch.object(
                product_task, "fetch_json", return_value=(404, None)
            ) as fetch, patch.object(
                product_task, "DELAY", 0
            ):
                product_task.run_product_task(self.context, self.config)
            self.assertEqual(fetch.call_count, count)
