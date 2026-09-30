from common import WorkspaceTest
from pathlib import PureWindowsPath
from unittest.mock import Mock, patch

from standard_collection_process.pipeline.models import PipelineStage
from standard_collection_process.storage import checkpoints, paths, workbooks


class StorageTests(WorkspaceTest):
    def test_atomic_failure_preserves_previous_file(self):
        path = self.folder / "saved.xlsx"
        workbooks.save_xlsx([{"name": "old"}], path)
        with patch.object(workbooks.os, "replace", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                workbooks.save_xlsx([{"name": "new"}], path)
        self.assertEqual(workbooks.load_xlsx(path)[1][0]["name"], "old")
        self.assertEqual(list(self.folder.iterdir()), [path])

    def test_no_obsolete_configuration_sheet_is_written(self):
        path = self.folder / "saved.xlsx"
        workbooks.save_xlsx([{"name": "A", "link-href": "url", "_crawl_config": "legacy"}], path)
        self.assertNotIn("_crawl_config", workbooks.load_xlsx(path)[1][0])
        wb = workbooks.openpyxl.load_workbook(path)
        try:
            self.assertEqual(wb.sheetnames, ["Sheet"])
        finally:
            wb.close()

    def test_failed_save_keeps_data_and_blocks_exit_until_retry(self):
        path = self.folder / "retry.xlsx"
        writer = Mock(side_effect=PermissionError("locked"))
        with self.assertRaises(PermissionError):
            self.controller.checkpoints.save("retry", writer, [{"name": "retained"}], path)
        self.controller.request_shutdown()
        with self.assertRaises(PermissionError):
            self.controller.try_close()
        self.assertEqual(self.controller.state.stage, PipelineStage.FAILED)
        self.assertTrue(self.controller.checkpoints.pending)
        writer.side_effect = workbooks.save_xlsx
        self.controller.request_shutdown()
        self.assertTrue(self.controller.try_close())
        self.assertFalse(self.controller.checkpoints.pending)
        self.assertEqual(workbooks.load_xlsx(path)[1][0]["name"], "retained")

    def test_close_waits_for_worker(self):
        self.controller.worker_thread = Mock()
        self.controller.worker_thread.is_alive.return_value = True
        self.assertFalse(self.controller.try_close())

    def test_path_mapping_replaces_all_occurrences(self):
        path = PureWindowsPath("D:/01INPUT_XLSX/category/01INPUT_XLSX-shop.xlsx")
        self.assertEqual(
            paths.map_to_output_path(path),
            PureWindowsPath("D:/02OUTPUT_XLSX/category/02OUTPUT_XLSX-shop.xlsx"),
        )

    def test_existing_links_are_deduplicated(self):
        workbooks.write_link_workbooks(
            [
                {"title": "First", "link": "https://shop.test/products/a"},
                {"title": "Other", "link": "https://shop.test/products/a?variant=1"},
            ],
            self.folder,
        )
        self.assertEqual(
            workbooks.load_existing_link_rows(self.folder),
            [{"title": "First", "link": "https://shop.test/products/a"}],
        )
