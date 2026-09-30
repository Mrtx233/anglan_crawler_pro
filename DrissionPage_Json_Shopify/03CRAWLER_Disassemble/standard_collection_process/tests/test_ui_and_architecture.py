from common import ROOT, WorkspaceTest
import ast
import sys
import threading
import tkinter as tk
from unittest.mock import Mock, patch

from standard_collection_process.pipeline import link_task
from standard_collection_process.pipeline.models import (
    LinkTaskConfig,
    PipelineStage,
    TaskEvent,
    TaskResult,
)
from standard_collection_process.storage import workbooks
from standard_collection_process.ui.app import ShopifyPipelineApp
from standard_collection_process.ui.theme import UI_COLORS


class ArchitectureTests(WorkspaceTest):
    def test_background_layers_do_not_import_tk_or_ui(self):
        for directory in ("pipeline", "crawler", "processing", "storage", "resources"):
            for path in (ROOT / directory).glob("*.py"):
                tree = ast.parse(path.read_text(encoding="utf-8"))
                for node in ast.walk(tree):
                    if isinstance(node, ast.Import):
                        self.assertFalse(
                            any(a.name.startswith("tkinter") for a in node.names), path
                        )
                    if isinstance(node, ast.ImportFrom):
                        self.assertNotIn("tkinter", node.module or "", path)
                        self.assertFalse((node.module or "").startswith("ui"), path)
                    if isinstance(node, ast.Attribute):
                        self.assertNotEqual(node.attr, "root", path)


class UiTests(WorkspaceTest):
    def setUp(self):
        super().setUp()
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        self.root = tk.Tk()
        self.root.withdraw()
        self.app = ShopifyPipelineApp(self.root, self.controller)
        self.app.folder_var.set(str(self.folder))
        self.addCleanup(self.close_window)

    def close_window(self):
        if not self.app._closed:
            if self.app._event_poll_id is not None:
                self.root.after_cancel(self.app._event_poll_id)
            self.root.destroy()

    def drain(self):
        self.root.after_cancel(self.app._event_poll_id)
        self.app._drain_events()

    def test_conversion_selects_and_passes_folder(self):
        self.output.mkdir(parents=True)
        with patch(
            "standard_collection_process.ui.app.filedialog.askdirectory",
            return_value=str(self.output),
        ):
            self.app._browse_conversion_folder()
        self.assertEqual(self.app.conversion_folder_var.get(), str(self.output))
        with patch.object(self.app, "_launch_task") as launch:
            self.app._start_conversion()
        self.assertEqual(launch.call_args.args[1].folder, self.output)

    def test_initialization_does_not_redirect_global_stdout(self):
        self.assertIs(sys.stdout, self.original_stdout)
        self.assertIs(sys.stderr, self.original_stderr)
        self.assertTrue(self.root.protocol("WM_DELETE_WINDOW"))

    def test_worker_events_update_widgets_only_on_main_thread(self):
        main_thread = threading.get_ident()
        updates = []
        original = self.app.logs.append

        def append(text):
            updates.append(threading.get_ident())
            original(text)

        with patch.object(link_task, "create_browser"), patch.object(
            link_task, "collect_page_links", return_value=["https://shop.test/products/a"]
        ), patch.object(self.app.logs, "append", side_effect=append):
            self.controller.start(
                PipelineStage.COLLECTING_LINKS,
                LinkTaskConfig((("Category", "https://shop.test/collections/a"),), self.folder, 1),
            )
            self.wait_worker()
            self.assertEqual(updates, [])
            self.drain()
        self.assertTrue(updates)
        self.assertEqual(set(updates), {main_thread})
        self.assertEqual(len(self.app.file_rows), 1)
        self.assertEqual(self.controller.state.stage, PipelineStage.WAITING_CONTINUE)

    def test_both_scanning_paths_have_final_image_hint(self):
        workbooks.write_link_workbooks(
            [{"title": "Category", "link": "https://shop.test/products/a"}], self.folder
        )

        def labels(widget):
            values = []
            for child in widget.winfo_children():
                if isinstance(child, tk.Label):
                    values.append(child.cget("text"))
                values.extend(labels(child))
            return values

        self.app._scan_files()
        self.assertIn("如 1,3,-1", labels(self.app.file_card))
        self.app._on_completed(TaskResult(PipelineStage.COLLECTING_LINKS, link_count=1))
        self.assertIn("如 1,3,-1", labels(self.app.file_card))

    def test_failed_event_remains_red_and_pending_save_prevents_close(self):
        self.controller.state.stage = PipelineStage.FAILED
        self.controller.state.task_error = "failed"
        self.controller.events.put(
            TaskEvent("failed", PipelineStage.PROCESSING_PRODUCTS, ("failed",))
        )
        self.drain()
        self.app._apply_stage_controls()
        self.assertEqual(self.app.status_dot.cget("fg"), UI_COLORS["danger"])
        self.controller.checkpoints.pending["file"] = (
            Mock(side_effect=PermissionError("locked")),
            (),
        )
        with patch("standard_collection_process.ui.app.messagebox.showerror") as message:
            self.app._request_shutdown()
        message.assert_called_once()
        self.assertFalse(self.app._closed)
        self.assertEqual(self.app.status_dot.cget("fg"), UI_COLORS["danger"])
        self.assertTrue(self.controller.checkpoints.pending)
