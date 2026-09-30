from common import WorkspaceTest, payload
from contextlib import ExitStack
from pathlib import Path
from unittest.mock import Mock, patch

from standard_collection_process.config import MAX_STYLES_LENGTH
from standard_collection_process.pipeline import controller, link_task, product_task
from standard_collection_process.pipeline.models import (
    LinkTaskConfig,
    PipelineStage,
    ProductFile,
    ProductTaskConfig,
)
from standard_collection_process.processing import product
from standard_collection_process.storage import checkpoints, workbooks


class PipelineTests(WorkspaceTest):
    def product_patches(self, fetch, targets=None, browser=None):
        stack = ExitStack()
        self.addCleanup(stack.close)
        stack.enter_context(
            patch.object(product_task, "create_browser", return_value=browser or Mock())
        )
        stack.enter_context(patch.object(product_task, "fetch_exchange_rates", return_value={}))
        stack.enter_context(patch.object(product_task, "DELAY", 0))
        stack.enter_context(
            patch.object(
                product_task,
                "read_targets_from_excel",
                return_value=targets
                or [
                    ("Category", "https://shop.test/products/a"),
                    ("Category", "https://shop.test/products/b"),
                ],
            )
        )
        return stack.enter_context(patch.object(product_task, "fetch_json", side_effect=fetch))

    def test_exception_saves_current_file_and_merge_before_browser_closes(self):
        browser = Mock()

        def quit_browser():
            for name in ("shop.xlsx", "category_合并.xlsx"):
                self.assertEqual(len(workbooks.load_xlsx(self.output / name)[1]), 1)

        browser.quit.side_effect = quit_browser
        self.product_patches([(200, payload()), RuntimeError("network lost")], browser=browser)
        self.controller.start(PipelineStage.PROCESSING_PRODUCTS, self.config)
        self.wait_worker()
        browser.quit.assert_called_once()
        self.assertEqual(self.controller.state.stage, PipelineStage.FAILED)
        self.assertTrue(self.controller.shutdown_requested)
        self.assertTrue(self.controller.try_close())
        self.assertIn("failed", [event.kind for event in self.events()])

    def test_stop_during_response_keeps_completed_product(self):
        def fetch(*args, **kwargs):
            self.controller.request_shutdown()
            return 200, payload()

        self.product_patches(fetch)
        self.controller.start(PipelineStage.PROCESSING_PRODUCTS, self.config)
        self.wait_worker()
        self.assertEqual(self.controller.state.stage, PipelineStage.STOPPED)
        self.assertEqual(len(workbooks.load_xlsx(self.output / "shop.xlsx")[1]), 1)

    def test_browser_close_failure_does_not_lose_data(self):
        browser = Mock()
        browser.quit.side_effect = RuntimeError("already closed")
        self.product_patches([(200, payload()), (404, None)], browser=browser)
        product_task.run_product_task(self.context, self.config)
        self.assertEqual(len(workbooks.load_xlsx(self.output / "shop.xlsx")[1]), 1)

    def test_only_one_validation_and_one_save_per_revision(self):
        targets = [("Category", f"https://shop.test/products/{i}") for i in range(22)]
        self.product_patches([(200, payload()) for _ in range(20)] + [(404, None)] * 2, targets)
        with patch.object(
            product, "validate_product_json", wraps=product.validate_product_json
        ) as validate, patch.object(product_task, "save_xlsx", wraps=workbooks.save_xlsx) as save:
            product_task.run_product_task(self.context, self.config)
        self.assertEqual(validate.call_count, 20)
        self.assertEqual(
            [Path(call.args[1]).name for call in save.call_args_list],
            ["shop.xlsx"],
        )

    def test_invalid_response_is_not_completed(self):
        self.product_patches(
            [(200, {"error": "not a product"})], [("Category", "https://shop.test/products/a")]
        )
        self.controller.start(PipelineStage.PROCESSING_PRODUCTS, self.config)
        self.wait_worker()
        self.assertFalse((self.output / "shop.xlsx").exists())
        self.assertEqual(self.controller.state.stage, PipelineStage.FAILED)

    def test_too_long_styles1_product_is_skipped(self):
        data = payload()
        long_value = "V" * 200
        data["product"]["options"] = [
            {"name": "Color", "position": 1, "values": [f"{long_value}{i:04d}" for i in range(300)]}
        ]
        self.product_patches(
            [(200, data)],
            [("Category", "https://shop.test/products/a")],
        )
        product_task.run_product_task(self.context, self.config)
        self.assertFalse((self.output / "shop.xlsx").exists())
        self.assertFalse((self.output / "category_合并.xlsx").exists())
        self.assertEqual(self.controller.state.failed_product_count, 1)
        logs = "".join(event.values[0] for event in self.events() if event.kind == "log")
        self.assertIn("styles1 长度", logs)
        self.assertIn(str(MAX_STYLES_LENGTH), logs)

    def test_resume_skips_completed_urls_regardless_of_config(self):
        url = "https://shop.test/products/a"
        path = self.output / "shop.xlsx"
        self.config = ProductTaskConfig(self.folder, (ProductFile(self.folder / "shop.xlsx", "2"),))
        for config, expected_fetches in (
            ("legacy-2", 0),
            ("legacy-1", 0),
            (None, 0),
        ):
            row = {"name": "Old", "link-href": url}
            if config:
                row["_crawl_config"] = config
            workbooks.save_xlsx([row], path)
            with patch.object(product_task, "create_browser"), patch.object(
                product_task, "fetch_exchange_rates", return_value={}
            ), patch.object(
                product_task, "read_targets_from_excel", return_value=[("Category", url)]
            ), patch.object(
                product_task, "fetch_json", return_value=(200, payload())
            ) as fetch, patch.object(
                product_task, "DELAY", 0
            ):
                product_task.run_product_task(self.context, self.config)
            self.assertEqual(fetch.call_count, expected_fetches)
            rows = workbooks.load_xlsx(path)[1]
            self.assertEqual(len(rows), 1)
            self.assertNotIn("_crawl_config", rows[0])
            if expected_fetches:
                self.assertNotIn("2_600x600.jpg", rows[0]["src_links"])

    def test_existing_links_do_not_end_paging_and_categories_save_immediately(self):
        a, b, c = [f"https://shop.test/products/{name}" for name in "abc"]
        workbooks.write_link_workbooks([{"title": "Original", "link": a}], self.folder)
        browser = Mock()

        def navigate(url):
            if "/two" in url:
                self.assertIn(
                    b, [row["link"] for row in workbooks.load_existing_link_rows(self.folder)]
                )
            return True

        browser.latest_tab.get.side_effect = navigate
        with patch.object(link_task, "create_browser", return_value=browser), patch.object(
            link_task, "collect_page_links", side_effect=[[a], [b], [b], [c]]
        ):
            rows = link_task.collect_category_links(
                self.context,
                [
                    ("One", "https://shop.test/collections/one"),
                    ("Two", "https://shop.test/collections/two"),
                ],
                2,
                "",
                self.folder,
            )
        self.assertEqual([row["link"] for row in rows], [a, b, c])
        self.assertEqual(rows[0]["title"], "Original")

    def test_only_changed_domains_are_written(self):
        with patch.object(link_task, "create_browser"), patch.object(
            link_task,
            "collect_page_links",
            side_effect=[[f"https://{name}.test/products/a"] for name in "abc"],
        ), patch.object(
            workbooks, "save_workbook_atomic", wraps=workbooks.save_workbook_atomic
        ) as save:
            link_task.collect_category_links(
                self.context,
                [(name, f"https://{name}.test/collections/all") for name in "abc"],
                1,
                "",
                self.folder,
            )
        self.assertEqual(
            [Path(call.args[1]).name for call in save.call_args_list],
            ["a.test.xlsx", "b.test.xlsx", "c.test.xlsx"],
        )

    def test_link_exception_saves_partial_category(self):
        with patch.object(link_task, "create_browser"), patch.object(
            link_task,
            "collect_page_links",
            side_effect=[["https://shop.test/products/a"], RuntimeError("page failed")],
        ):
            with self.assertRaises(RuntimeError):
                link_task.collect_category_links(
                    self.context,
                    [("Category", "https://shop.test/collections/a")],
                    0,
                    "",
                    self.folder,
                )
        self.assertEqual(len(workbooks.load_existing_link_rows(self.folder)), 1)

    def test_failed_save_retains_all_dirty_domains(self):
        with patch.object(link_task, "create_browser"), patch.object(
            link_task,
            "collect_page_links",
            return_value=["https://a.test/products/a", "https://b.test/products/b"],
        ), patch.object(workbooks, "save_workbook_atomic", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                link_task.collect_category_links(
                    self.context, [("Category", "https://a.test/collections/a")], 1, "", self.folder
                )
        self.assertEqual(set(self.controller.checkpoints.pending), {"links:a.test", "links:b.test"})
        self.controller.checkpoints.retry_all()
        self.assertEqual(len(workbooks.load_existing_link_rows(self.folder)), 2)

    def test_bad_existing_file_aborts_before_browser_start(self):
        path = self.folder / "shop.xlsx"
        path.write_bytes(b"broken workbook")
        with patch.object(link_task, "create_browser") as browser:
            with self.assertRaises(Exception):
                link_task.collect_category_links(self.context, [], 0, "", self.folder)
        browser.assert_not_called()
        self.assertEqual(path.read_bytes(), b"broken workbook")

    def test_controller_captures_result_and_rejects_overlapping_jobs(self):
        import threading

        entered, release = threading.Event(), threading.Event()

        def task(context, config):
            entered.set()
            release.wait(5)
            return controller.TaskResult(PipelineStage.COLLECTING_LINKS, link_count=1)

        config = LinkTaskConfig((), self.folder)
        with patch.object(controller, "run_link_task", side_effect=task):
            self.controller.start(PipelineStage.COLLECTING_LINKS, config)
            try:
                self.assertTrue(entered.wait(3))
                with self.assertRaises(RuntimeError):
                    self.controller.start(PipelineStage.COLLECTING_LINKS, config)
            finally:
                release.set()
                self.wait_worker()
        self.assertEqual(self.controller.state.stage, PipelineStage.WAITING_CONTINUE)
        self.assertIn("completed", [event.kind for event in self.events()])
