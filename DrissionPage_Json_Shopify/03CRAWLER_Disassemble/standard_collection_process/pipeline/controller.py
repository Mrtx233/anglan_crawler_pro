from __future__ import annotations

from queue import Queue
import threading
import traceback
from ..pipeline.conversion_task import run_conversion_task
from ..pipeline.link_task import run_link_task
from ..pipeline.models import (
    ConversionTaskConfig,
    LinkTaskConfig,
    PipelineStage,
    ProductTaskConfig,
    TaskEvent,
    TaskResult,
    TaskState,
    next_stage_after_link_collection,
)
from ..pipeline.product_task import run_product_task
from ..storage.checkpoints import CheckpointStore


class TaskContext:
    """Worker-facing services; no GUI callbacks or widgets are accepted."""

    def __init__(self, state, stage, stop_event, checkpoints, events):
        self.state = state
        self.stage = stage
        self.stop_event = stop_event
        self.checkpoints = checkpoints
        self.events = events

    def emit(self, kind, *values):
        self.events.put(TaskEvent(kind, self.stage, tuple(values)))

    def log(self, *values, sep=" ", end="\n"):
        self.emit("log", sep.join(str(value) for value in values) + end)

    def merge_ready(self, path):
        self.state.last_merge_path = str(path)
        self.emit("merge_ready", str(path))


class PipelineController:
    """Own worker lifetime, task state, checkpoint recovery and shutdown."""

    def __init__(self):
        self.state = TaskState()
        self.stop_event = threading.Event()
        self.checkpoints = CheckpointStore()
        self.events = Queue()
        self.worker_thread = None
        self.shutdown_requested = False
        self._running = False

    @property
    def is_running(self):
        return self._running or (self.worker_thread is not None and self.worker_thread.is_alive())

    def start(self, stage, config):
        if self.is_running or self.shutdown_requested or self.checkpoints.pending:
            raise RuntimeError("当前任务或待保存数据尚未处理完成")
        tasks = {
            PipelineStage.COLLECTING_LINKS: (LinkTaskConfig, run_link_task),
            PipelineStage.PROCESSING_PRODUCTS: (ProductTaskConfig, run_product_task),
            PipelineStage.CONVERTING: (ConversionTaskConfig, run_conversion_task),
        }
        expected, target = tasks[stage]
        if not isinstance(config, expected):
            raise TypeError(f"{stage.value} 配置类型错误")
        old_state = self.state
        self.state = TaskState(stage=stage)
        if stage == PipelineStage.CONVERTING:
            self.state.last_merge_path = old_state.last_merge_path
        self.stop_event.clear()
        self._running = True
        self.worker_thread = threading.Thread(
            target=self._execute, args=(stage, target, config), daemon=True
        )
        try:
            self.worker_thread.start()
        except Exception as exc:
            self._running = False
            self.state.stage = PipelineStage.FAILED
            self.state.task_error = str(exc) or type(exc).__name__
            raise

    def _execute(self, stage, target, config):
        context = TaskContext(self.state, stage, self.stop_event, self.checkpoints, self.events)
        try:
            result = target(context, config) or TaskResult(
                stage, output_path=self.state.last_merge_path
            )
            if self.stop_event.is_set():
                self.state.stage = PipelineStage.STOPPED
            elif stage == PipelineStage.COLLECTING_LINKS:
                self.state.stage = next_stage_after_link_collection(result.link_count, False)
            elif stage == PipelineStage.PROCESSING_PRODUCTS:
                if self.state.last_merge_path:
                    self.state.stage = PipelineStage.WAITING_CONVERSION
                else:
                    self.state.stage = (
                        PipelineStage.FAILED
                        if self.state.failed_product_count
                        else PipelineStage.COMPLETED
                    )
            else:
                self.state.stage = PipelineStage.COMPLETED
            context.emit("completed", result)
        except Exception as exc:
            self.state.task_error = str(exc) or type(exc).__name__
            self.state.stage = PipelineStage.FAILED
            context.log(traceback.format_exc())
            context.emit("failed", self.state.task_error)
            if stage != PipelineStage.CONVERTING:
                self.request_shutdown()
                context.emit("shutdown_requested")
        finally:
            self._running = False

    def request_shutdown(self):
        self.shutdown_requested = True
        self.stop_event.set()

    def request_stop(self):
        """Stop the active task while keeping the GUI open for review or restart."""
        self.stop_event.set()

    def try_close(self):
        """Return False while running; retain snapshots and raise if saving fails."""
        if self.is_running:
            return False
        try:
            self.checkpoints.retry_all()
        except Exception as exc:
            self.shutdown_requested = False
            self.state.task_error = f"保存失败：{exc}"
            self.state.stage = PipelineStage.FAILED
            raise
        return True
