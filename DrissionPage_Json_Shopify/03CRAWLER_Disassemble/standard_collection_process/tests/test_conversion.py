from common import REFERENCE, WorkspaceTest
import random
from unittest.mock import patch

from standard_collection_process.pipeline import conversion_task
from standard_collection_process.pipeline.models import ConversionTaskConfig, PipelineStage
from standard_collection_process.processing import pricing, shopify, woocommerce
from standard_collection_process.storage import workbooks


class ConversionTests(WorkspaceTest):
    def test_wp_string_zero_regular_price_is_preserved(self):
        for option in ("Default Title", "M"):
            for regular in (0, "0", "0.00"):
                with self.subTest(option=option, regular=regular):
                    rows = woocommerce.build_wp_rows(
                        [
                            {
                                "Handle": "test-product",
                                "Option1 Value": option,
                                "Variant Price": "79.50",
                                "Variant Compare At Price": regular,
                            }
                        ],
                        lambda text: None,
                    )
                    for row in rows:
                        if row["Type"] == "variable":
                            self.assertEqual(row["Regular price"], "")
                        else:
                            self.assertEqual(
                                row["Regular price"], "79.50" if regular == 0 else regular
                            )

    def test_all_csv_outputs_match_pre_split_reference(self):
        source = self.output / "fixture.xlsx"
        workbooks.save_xlsx(REFERENCE["product_rows"], source)
        random.seed(2026)
        result = conversion_task.convert_styles_file(source, lambda text: None)
        self.assertTrue(result.exists())
        outputs = {
            str(path.relative_to(self.output)).replace("\\", "/"): path.read_text(
                encoding="utf-8-sig"
            )
            for path in self.output.rglob("*.csv")
        }
        expected = REFERENCE["csvs"]
        self.assertEqual(set(outputs), set(expected))
        for name, content in expected.items():
            if "价格匹配" in name:
                # Re-reading Excel may serialize booleans/numeric cells differently.
                import pandas as pd
                from io import StringIO

                pd.testing.assert_frame_equal(
                    pd.read_csv(StringIO(outputs[name])),
                    pd.read_csv(StringIO(content)),
                    check_dtype=False,
                )
            else:
                self.assertEqual(outputs[name], content)

    def test_matched_wp_generated_with_xlsx_in_directory_and_filename(self):
        source = self.output / "xlsx_archive" / "fixture_xlsx.xlsx"
        workbooks.save_xlsx(REFERENCE["product_rows"], source)
        random.seed(2026)
        conversion_task.convert_styles_file(source, lambda text: None)
        result = source.parent / "csv" / "wp-fixture_xlsx" / "wp-fixture_xlsx_Shopify_价格匹配.csv"
        import pandas as pd
        from io import StringIO

        self.assertTrue(result.exists())
        pd.testing.assert_frame_equal(
            pd.read_csv(result),
            pd.read_csv(
                StringIO(REFERENCE["csvs"]["csv/wp-fixture/wp-fixture_Shopify_价格匹配.csv"])
            ),
            check_dtype=False,
        )

    def test_transformation_kernels_do_not_read_or_write_files(self):
        import pandas as pd

        frame = pd.DataFrame(REFERENCE["product_rows"])
        with patch.object(
            pd, "read_excel", side_effect=AssertionError("read in kernel")
        ), patch.object(pd.DataFrame, "to_excel", side_effect=AssertionError("write in kernel")):
            shopify_frame = shopify.build_shopify_frame(frame)
            original = shopify_frame.copy(deep=True)
            matched, unused = pricing.match_prices(shopify_frame)
            rows = [
                {key: str(value) for key, value in row.items()}
                for row in matched.to_dict("records")
            ]
            self.assertTrue(woocommerce.build_wp_rows(rows, lambda text: None))
            pd.testing.assert_frame_equal(shopify_frame, original)
            self.assertTrue(unused)

    def test_conversion_worker_returns_completion_event(self):
        source = self.output / "fixture.xlsx"
        workbooks.save_xlsx(REFERENCE["product_rows"], source)
        self.controller.start(PipelineStage.CONVERTING, ConversionTaskConfig(self.output))
        self.wait_worker()
        self.assertEqual(self.controller.state.stage, PipelineStage.COMPLETED)
        result = next(event.values[0] for event in self.events() if event.kind == "completed")
        self.assertTrue(result.output_path.endswith("category_合并_Shopify_原价.xlsx"))
