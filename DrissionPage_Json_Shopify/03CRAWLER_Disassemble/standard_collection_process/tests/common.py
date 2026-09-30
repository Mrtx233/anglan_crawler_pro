import copy
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT.parent))

from standard_collection_process.pipeline.controller import PipelineController, TaskContext
from standard_collection_process.pipeline.models import (
    PipelineStage,
    ProductFile,
    ProductTaskConfig,
)

REFERENCE = json.loads((ROOT / "tests/fixtures/reference.json").read_text(encoding="utf-8"))


def payload():
    return copy.deepcopy(REFERENCE["payload"])


class WorkspaceTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name) / "01INPUT_XLSX" / "category"
        self.folder.mkdir(parents=True)
        self.output = Path(self.temp.name) / "02OUTPUT_XLSX" / "category"
        self.controller = PipelineController()
        self.context = TaskContext(
            self.controller.state,
            PipelineStage.PROCESSING_PRODUCTS,
            self.controller.stop_event,
            self.controller.checkpoints,
            self.controller.events,
        )
        self.config = ProductTaskConfig(self.folder, (ProductFile(self.folder / "shop.xlsx", "1"),))

    def wait_worker(self):
        self.controller.worker_thread.join(timeout=10)
        self.assertFalse(self.controller.worker_thread.is_alive(), "Worker did not finish")

    def events(self):
        result = []
        while not self.controller.events.empty():
            result.append(self.controller.events.get_nowait())
        return result
