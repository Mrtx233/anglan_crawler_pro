from common import WorkspaceTest
from unittest.mock import patch
from standard_collection_process.storage import workbooks
from standard_collection_process.pipeline.conversion_task import run_conversion_task
from standard_collection_process.pipeline.models import ConversionTaskConfig


class FolderMergeTests(WorkspaceTest):
    def test_sources_merge_once_and_preserve_duplicate_rows(self):
        for filename, name in [("a.xlsx", "A"), ("b.XLSX", "B")]:
            workbooks.save_xlsx([{"name": name}, {"name": "Shared"}], self.output / filename)
        for filename in ("old_合并.xlsx", "old_Shopify_原价.xlsx", "~$a.xlsx"):
            (self.output / filename).write_text("ignore", encoding="utf-8")
        (self.output / "xlsx").mkdir()
        (self.output / "xlsx" / "export.xlsx").write_text("ignore", encoding="utf-8")
        for _ in range(2):
            with patch(
                "standard_collection_process.pipeline.conversion_task.convert_styles_file"
            ) as convert:
                run_conversion_task(self.context, ConversionTaskConfig(self.output))
                merged = convert.call_args.args[0]
            self.assertEqual(
                [r["name"] for r in workbooks.load_xlsx(merged)[1]], ["A", "Shared", "B", "Shared"]
            )

    def test_invalid_source_preserves_previous_merge(self):
        target = self.output / "category_合并.xlsx"
        workbooks.save_xlsx([{"name": "old"}], target)
        wb = workbooks.openpyxl.Workbook()
        wb.active.append(["title", "link"])
        wb.save(self.output / "wrong.xlsx")
        wb.close()
        with self.assertRaisesRegex(ValueError, "wrong.xlsx"):
            run_conversion_task(self.context, ConversionTaskConfig(self.output))
        self.assertEqual(workbooks.load_xlsx(target)[1][0]["name"], "old")

    def test_empty_folder_fails_without_export(self):
        self.output.mkdir(parents=True)
        with self.assertRaises(ValueError):
            run_conversion_task(self.context, ConversionTaskConfig(self.output))
        self.assertEqual(list(self.output.iterdir()), [])
