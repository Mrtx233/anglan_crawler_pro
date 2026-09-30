from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from pathlib import Path


class PipelineStage(str, Enum):
    IDLE = "idle"
    COLLECTING_LINKS = "collecting_links"
    WAITING_CONTINUE = "waiting_continue"
    PROCESSING_PRODUCTS = "processing_products"
    WAITING_CONVERSION = "waiting_conversion"
    CONVERTING = "converting"
    COMPLETED = "completed"
    FAILED = "failed"
    STOPPED = "stopped"


def next_stage_after_link_collection(
    link_count: int,
    stopped: bool,
) -> PipelineStage:
    if stopped:
        return PipelineStage.STOPPED
    if link_count <= 0:
        return PipelineStage.FAILED
    return PipelineStage.WAITING_CONTINUE


def can_continue_processing(stage: PipelineStage) -> bool:
    return stage == PipelineStage.WAITING_CONTINUE


@dataclass(frozen=True)
class ProductFile:
    path: Path
    skip: str = ""
    keep: str = ""


@dataclass(frozen=True)
class ProductTaskConfig:
    folder: Path
    files: tuple[ProductFile, ...]
    skip_options: tuple[str, ...] = ()


@dataclass(frozen=True)
class LinkTaskConfig:
    pages: tuple[tuple[str, str], ...]
    folder: Path
    max_pages: int = 0
    xpath: str = ""


@dataclass(frozen=True)
class ConversionTaskConfig:
    folder: Path


@dataclass(frozen=True)
class TaskResult:
    stage: PipelineStage
    link_count: int = 0
    output_path: str = ""


@dataclass(frozen=True)
class TaskEvent:
    kind: str
    stage: PipelineStage
    values: tuple = ()


@dataclass
class TaskState:
    stage: PipelineStage = PipelineStage.IDLE
    task_error: str = ""
    failed_product_count: int = 0
    active_file_index: int | None = None
    last_merge_path: str = ""
