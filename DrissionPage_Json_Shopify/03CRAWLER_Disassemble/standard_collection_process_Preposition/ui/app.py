from __future__ import annotations

from pathlib import Path
from queue import Empty
from tkinter import filedialog, messagebox, scrolledtext, ttk
import tkinter as tk
from ..config import SKIP_OPTIONS
from ..pipeline.controller import PipelineController
from ..pipeline.models import (
    ConversionTaskConfig,
    LinkTaskConfig,
    PipelineStage,
    ProductFile,
    ProductTaskConfig,
    can_continue_processing,
)
from ..storage.paths import parse_category_pages
from ..ui.components import UiComponents
from ..ui.log_window import LogWindow
from ..ui.theme import MONO_FONT, UI_COLORS, UI_FONT, configure_styles


class ShopifyPipelineApp(UiComponents):
    """Main-thread UI; task code communicates exclusively through queue events."""

    def __init__(self, root, controller=None):
        self.root = root
        self.controller = controller or PipelineController()
        self._closing = False
        self._closed = False
        self._event_poll_id = None
        self.file_rows = []
        root.title("Shopify_Preposition")
        root.geometry("1040x980")
        root.minsize(940, 760)
        root.configure(bg=UI_COLORS["background"])
        root.protocol("WM_DELETE_WINDOW", self._request_shutdown)
        self.style = ttk.Style(root)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")
        configure_styles(self.style)
        self.folder_var = tk.StringVar(master=root)
        self.status_var = tk.StringVar(master=root, value="请选择分类文件夹后点击读取")
        self.progress_var = tk.DoubleVar(master=root, value=0)
        self.progress_percent_var = tk.StringVar(master=root, value="0%")
        self.file_count_var = tk.StringVar(master=root, value="0 个文件")
        self.max_pages_var = tk.StringVar(master=root)
        self.xpath_var = tk.StringVar(master=root)
        self.conversion_folder_var = tk.StringVar(master=root)
        self.logs = LogWindow(root)
        self._build_ui()
        self._center_window()
        self._apply_stage_controls()
        self._event_poll_id = root.after(50, self._drain_events)

    def _can_start_task(self):
        if (
            self.controller.is_running
            or self.controller.shutdown_requested
            or self.controller.checkpoints.pending
        ):
            messagebox.showerror(
                "任务尚未完成", "请等待任务结束；保存失败时解除文件占用，再关闭窗口重试。"
            )
            return False
        return True

    def _launch_task(self, stage, config, *, total=100, text="", clear_logs=False):
        if clear_logs:
            self.logs._clear_logs()
        try:
            self.controller.start(stage, config)
        except Exception as exc:
            messagebox.showerror("无法启动任务", str(exc))
            self.update_status(0, 0, f"无法启动任务：{exc}")
            self.status_dot.configure(fg=UI_COLORS["danger"])
            self._apply_stage_controls()
            return
        self.update_status(0, total, text)
        self.status_dot.configure(fg=UI_COLORS["primary"])
        self._apply_stage_controls()

    def _start(self):
        if not self._can_start_task() or not self.file_rows:
            return
        config = ProductTaskConfig(
            Path(self.folder_var.get().strip()),
            tuple(
                ProductFile(
                    Path(row["path"]), row["skip_var"].get().strip(), row["keep_var"].get().strip()
                )
                for row in self.file_rows
            ),
            tuple(SKIP_OPTIONS or ()),
        )
        for index in range(len(self.file_rows)):
            self._set_file_status(index, "等待中", "waiting")
        self._launch_task(
            PipelineStage.PROCESSING_PRODUCTS, config, text="正在初始化任务...", clear_logs=True
        )

    def _result_color(self):
        state = self.controller.state
        if (
            state.task_error
            or self.controller.checkpoints.pending
            or state.stage == PipelineStage.FAILED
        ):
            return UI_COLORS["danger"]
        if self.controller.stop_event.is_set() or state.failed_product_count:
            return UI_COLORS["warning"]
        return UI_COLORS["success"]

    def _drain_events(self):
        for _ in range(200):
            try:
                event = self.controller.events.get_nowait()
            except Empty:
                break
            if event.kind == "log":
                self.logs.append(event.values[0])
            elif event.kind == "progress":
                self.update_status(*event.values)
            elif event.kind == "file_status":
                self._set_file_status(*event.values)
            elif event.kind == "merge_ready":
                self.conversion_folder_var.set(str(Path(event.values[0]).parent))
            elif event.kind == "completed":
                self._on_completed(event.values[0])
            elif event.kind == "failed":
                self.update_status(0, 0, f"任务失败：{event.values[0]}")
                self._set_file_status(self.controller.state.active_file_index, "发生异常", "failed")
                self.status_dot.configure(fg=UI_COLORS["danger"])
                self._apply_stage_controls()
            elif event.kind == "shutdown_requested":
                self._request_shutdown()
            if self._closed:
                return
        self._event_poll_id = self.root.after(50, self._drain_events)

    def _on_completed(self, result):
        state = self.controller.state
        if result.stage == PipelineStage.COLLECTING_LINKS:
            if state.stage == PipelineStage.WAITING_CONTINUE:
                self._populate_file_rows()
                self.update_status(0, 100, "链接采集完成；设置图片后点击继续处理")
            elif state.stage == PipelineStage.FAILED:
                self.update_status(0, 100, "未采集到链接，请检查分类地址或 XPath")
        elif (
            result.stage == PipelineStage.PROCESSING_PRODUCTS
            and state.stage == PipelineStage.WAITING_CONVERSION
        ):
            detail = (
                f"；有 {state.failed_product_count} 个商品失败"
                if state.failed_product_count
                else ""
            )
            self.update_status(0, 1, f"第二阶段完成{detail}；请检查来源数据后点击“合并并转换”")
        elif result.stage == PipelineStage.CONVERTING:
            self.update_status(1, 1, f"第三阶段转换完成 · {result.output_path}")
        color = (
            UI_COLORS["warning"]
            if state.stage == PipelineStage.WAITING_CONTINUE
            else self._result_color()
        )
        self.status_dot.configure(fg=color)
        self._apply_stage_controls()

    def _stop(self):
        if not self.controller.is_running:
            return
        self.controller.request_stop()
        self.update_status(0, 0, "正在停止任务并保存已有数据，请稍候...")
        self.status_dot.configure(fg=UI_COLORS["warning"])
        self._apply_stage_controls()

    def _request_shutdown(self):
        if self._closing:
            return
        self._closing = True
        self.controller.request_shutdown()
        detail = f"{self.controller.state.task_error}；" if self.controller.state.task_error else ""
        self.update_status(0, 0, f"{detail}正在保存已有数据，完成后关闭程序...")
        self.status_dot.configure(fg=self._result_color())
        self._apply_stage_controls()
        self._poll_shutdown()

    def _poll_shutdown(self):
        try:
            ready = self.controller.try_close()
        except Exception as exc:
            self._closing = False
            self.status_dot.configure(fg=UI_COLORS["danger"])
            self._apply_stage_controls()
            messagebox.showerror(
                "保存失败，尚未关闭", f"数据仍保留在内存中，请解除文件占用后关闭窗口重试。\n{exc}"
            )
            return
        if not ready:
            self.root.after(100, self._poll_shutdown)
            return
        self._closed = True
        if self._event_poll_id is not None:
            self.root.after_cancel(self._event_poll_id)
        self.root.destroy()

    def _build_ui(self):
        page = tk.Frame(self.root, bg=UI_COLORS["background"])
        page.pack(fill=tk.BOTH, expand=True)

        content = tk.Frame(page, bg=UI_COLORS["background"])
        content.pack(fill=tk.BOTH, expand=True, padx=24, pady=(16, 14))

        body = tk.Frame(content, bg=UI_COLORS["background"])
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self._build_header(body)
        self._build_folder_card(body)
        self._build_link_card(body)
        self._build_file_card(body)
        self.file_card.configure(height=170)
        self._build_conversion_card(body)
        footer = tk.Frame(content, bg=UI_COLORS["background"])
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        self._build_progress_card(footer)

    def _build_header(self, parent):
        colors = UI_COLORS
        header = tk.Frame(parent, bg=colors["background"])
        self._header_frame = header
        header.pack(fill=tk.X, pady=(0, 12))

        logo = tk.Label(
            header,
            text="S",
            width=3,
            bg=colors["primary"],
            fg="#FFFFFF",
            font=(UI_FONT, 15, "bold"),
        )
        logo.pack(side=tk.LEFT, padx=(0, 12), ipady=6)

        title_group = tk.Frame(header, bg=colors["background"])
        title_group.pack(side=tk.LEFT)
        tk.Label(
            title_group,
            text="Shopify 一体化采集工作台",
            bg=colors["background"],
            fg=colors["text"],
            font=(UI_FONT, 18, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="链接采集 → 图片设置 → 商品采集、合并与转换",
            bg=colors["background"],
            fg=colors["text_secondary"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 0))

        self.topmost_btn = self._create_button(
            header, "置顶", self._toggle_topmost, kind="ghost", padx=12, pady=8
        )
        self.topmost_btn.pack(side=tk.RIGHT, padx=(8, 0), pady=4)
        self.log_btn = self._create_button(
            header, "查看运行日志", self.logs.show_logs, kind="ghost", padx=12, pady=8
        )
        self.log_btn.pack(side=tk.RIGHT, padx=(8, 0), pady=4)
        self._build_action_bar(header)

    def _build_link_card(self, parent):
        colors = UI_COLORS
        card = self._create_card(parent, fill=tk.X, pady=(0, 12))
        inner = tk.Frame(card, bg=colors["card"])
        inner.pack(fill=tk.X, padx=18, pady=12)

        heading = tk.Frame(inner, bg=colors["card"])
        heading.pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            heading,
            text="第一阶段 · 详情页链接采集",
            bg=colors["card"],
            fg=colors["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            heading,
            text="每行：标题, URL",
            bg=colors["card"],
            fg=colors["text_muted"],
            font=(UI_FONT, 9),
        ).pack(side=tk.RIGHT)

        content = tk.Frame(inner, bg=colors["card"])
        content.pack(fill=tk.X, pady=(0, 8))
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)

        pages_frame = tk.Frame(content, bg=colors["card"])
        pages_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        controls = tk.Frame(content, bg=colors["card"])
        controls.grid(row=0, column=1, sticky="new")
        controls.grid_columnconfigure(0, weight=1)

        tk.Label(
            pages_frame,
            text="分类地址（每行：标题, URL）",
            bg=colors["card"],
            fg=colors["text_secondary"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W)

        tk.Label(
            controls,
            text="XPath（可留空）",
            bg=colors["card"],
            fg=colors["text_secondary"],
            font=(UI_FONT, 9),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            controls,
            text="最大页数（留空表示不限制）",
            bg=colors["card"],
            fg=colors["text_secondary"],
            font=(UI_FONT, 9),
        ).grid(row=2, column=0, sticky="w", pady=(8, 0))

        self.xpath_entry = ttk.Entry(
            controls,
            textvariable=self.xpath_var,
            style="App.TEntry",
        )
        self.xpath_entry.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        self.max_pages_entry = ttk.Entry(
            controls,
            textvariable=self.max_pages_var,
            style="App.TEntry",
        )
        self.max_pages_entry.grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(3, 0),
        )

        self.pages_text = scrolledtext.ScrolledText(
            pages_frame,
            height=5,
            wrap=tk.NONE,
            bg=colors["card_alt"],
            fg=colors["text"],
            insertbackground=colors["primary"],
            font=(MONO_FONT, 9),
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=colors["border"],
            padx=10,
            pady=7,
        )
        self.pages_text.pack(fill=tk.BOTH, expand=True, pady=(3, 0))

    def _build_conversion_card(self, parent):
        colors = UI_COLORS
        card = self._create_card(parent, fill=tk.X, pady=(0, 10))
        inner = tk.Frame(card, bg=colors["card"])
        inner.pack(fill=tk.X, padx=18, pady=12)

        tk.Label(
            inner,
            text="第三阶段 · Shopify 转换",
            bg=colors["card"],
            fg=colors["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            inner,
            text="选择商品数据文件夹，合并其中的来源 XLSX 后转换",
            bg=colors["card"],
            fg=colors["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 8))

        file_row = tk.Frame(inner, bg=colors["card"])
        file_row.pack(fill=tk.X)
        self.conversion_entry = ttk.Entry(
            file_row,
            textvariable=self.conversion_folder_var,
            style="App.TEntry",
        )
        self.conversion_entry.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            padx=(0, 8),
        )
        self.conversion_browse_btn = self._create_button(
            file_row,
            "选择文件夹",
            self._browse_conversion_folder,
            kind="secondary",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.conversion_browse_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.conversion_start_btn = self._create_button(
            file_row,
            "合并并转换",
            self._start_conversion,
            kind="primary",
            padx=16,
            pady=8,
            font_size=9,
        )
        self.conversion_start_btn.pack(side=tk.LEFT)

    def _build_action_bar(self, parent):
        buttons = parent
        self.stop_btn = self._create_button(
            buttons, "停止", self._stop, kind="danger", padx=14, pady=9
        )
        self.stop_btn.pack(side=tk.RIGHT)
        self.continue_btn = self._create_button(
            buttons,
            "继续处理",
            self._continue_processing,
            kind="primary",
            padx=18,
            pady=9,
        )
        self.continue_btn.pack(side=tk.RIGHT, padx=(0, 8))
        self.existing_btn = self._create_button(
            buttons,
            "仅处理已有 XLSX",
            self._load_existing_files,
            kind="secondary",
            padx=14,
            pady=9,
        )
        self.existing_btn.pack(side=tk.RIGHT, padx=(0, 8))
        self.full_btn = self._create_button(
            buttons,
            "运行完整流程",
            self._start_full_pipeline,
            kind="primary",
            padx=16,
            pady=9,
        )
        self.full_btn.pack(side=tk.RIGHT, padx=(0, 8))

    def _toggle_topmost(self):
        value = not bool(self.root.attributes("-topmost"))
        self.root.attributes("-topmost", value)
        self.topmost_btn.configure(text="取消置顶" if value else "置顶")

    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择 01INPUT_XLSX 下的具体分类目录")
        if path:
            self.folder_var.set(path)

    def _validated_task_folder(self) -> Path | None:
        raw_folder = self.folder_var.get().strip()
        if not raw_folder:
            messagebox.showerror("路径无效", "请先选择链接保存/读取目录。")
            return None
        folder = Path(raw_folder)
        if not folder.is_dir():
            messagebox.showerror("路径无效", "所选分类目录不存在。")
            return None
        return folder

    def _scan_files(self):
        if self.controller.state.stage in {
            PipelineStage.COLLECTING_LINKS,
            PipelineStage.PROCESSING_PRODUCTS,
        }:
            return
        if self._validated_task_folder() is None:
            return
        self._populate_file_rows()
        if self.file_rows:
            self.controller.state.stage = PipelineStage.WAITING_CONTINUE
            self.update_status(
                0,
                100,
                "已读取链接文件；设置跳过图片位置后点击“继续处理”",
            )
            self.status_dot.configure(fg=UI_COLORS["warning"])
        else:
            self.controller.state.stage = PipelineStage.IDLE
        self._apply_stage_controls()

    def _load_existing_files(self):
        if self.controller.state.stage not in {
            PipelineStage.IDLE,
            PipelineStage.COMPLETED,
            PipelineStage.FAILED,
            PipelineStage.STOPPED,
        }:
            return
        self._scan_files()

    def _start_full_pipeline(self):
        if not self._can_start_task():
            return
        if self.controller.state.stage not in {
            PipelineStage.IDLE,
            PipelineStage.COMPLETED,
            PipelineStage.FAILED,
            PipelineStage.STOPPED,
        }:
            return
        pages = parse_category_pages(self.pages_text.get("1.0", tk.END))
        if not pages:
            messagebox.showerror("无法开始", "请先填写有效的分类列表。")
            return
        folder = self._validated_task_folder()
        if folder is None:
            return
        max_pages_raw = self.max_pages_var.get().strip()
        try:
            max_pages = int(max_pages_raw) if max_pages_raw else 0
        except ValueError:
            messagebox.showerror("页数无效", "最大页数必须是整数或留空。")
            return
        if max_pages < 0:
            messagebox.showerror("页数无效", "最大页数不能小于 0。")
            return

        self._launch_task(
            PipelineStage.COLLECTING_LINKS,
            LinkTaskConfig(tuple(pages), folder, max_pages, self.xpath_var.get().strip()),
            total=len(pages),
            text="第一阶段：正在启动链接采集...",
            clear_logs=True,
        )

    def _continue_processing(self):
        if not can_continue_processing(self.controller.state.stage):
            return
        if not self.file_rows:
            messagebox.showerror("无法继续", "当前没有可处理的 XLSX 文件。")
            return
        if self.controller.shutdown_requested or self.controller.checkpoints.pending:
            return
        self.controller.state.last_merge_path = ""
        self._start()

    def _browse_conversion_folder(self):
        if self.controller.state.stage in {
            PipelineStage.COLLECTING_LINKS,
            PipelineStage.PROCESSING_PRODUCTS,
            PipelineStage.CONVERTING,
        }:
            return
        path = filedialog.askdirectory(title="选择包含商品数据 XLSX 的文件夹")
        if path:
            self.conversion_folder_var.set(path)

    def _validated_conversion_folder(self):
        raw_path = self.conversion_folder_var.get().strip()
        folder = Path(raw_path) if raw_path else None
        if folder is None or not folder.is_dir():
            messagebox.showerror("文件夹无效", "请选择存在的商品数据文件夹。")
            return None
        return folder

    def _start_conversion(self):
        if (
            not self.controller.is_running
            and self.controller.checkpoints.pending
            and not self.controller.shutdown_requested
        ):
            try:
                self.controller.checkpoints.retry_all()
            except Exception as exc:
                messagebox.showerror("保存失败", f"数据仍保留，请解除文件占用后重试。\n{exc}")
                return
        if not self._can_start_task():
            return
        if self.controller.state.stage not in {
            PipelineStage.IDLE,
            PipelineStage.WAITING_CONVERSION,
            PipelineStage.COMPLETED,
            PipelineStage.FAILED,
            PipelineStage.STOPPED,
        }:
            return
        folder = self._validated_conversion_folder()
        if folder is None:
            return

        self._launch_task(
            PipelineStage.CONVERTING,
            ConversionTaskConfig(folder),
            total=1,
            text=f"第三阶段：正在合并并转换 {folder.name}",
        )

    def _apply_stage_controls(self):
        if not hasattr(self, "full_btn"):
            return
        idle = self.controller.state.stage in {
            PipelineStage.IDLE,
            PipelineStage.COMPLETED,
            PipelineStage.FAILED,
            PipelineStage.STOPPED,
        }
        waiting = self.controller.state.stage == PipelineStage.WAITING_CONTINUE
        conversion_ready = self.controller.state.stage in {
            PipelineStage.IDLE,
            PipelineStage.WAITING_CONVERSION,
            PipelineStage.COMPLETED,
            PipelineStage.FAILED,
            PipelineStage.STOPPED,
        }
        if self.controller.shutdown_requested or self.controller.checkpoints.pending:
            idle = waiting = conversion_ready = False
        self._set_button_enabled(self.full_btn, idle)
        self._set_button_enabled(self.existing_btn, idle)
        self._set_button_enabled(self.continue_btn, waiting)
        self._set_button_enabled(
            self.stop_btn,
            not self.controller.shutdown_requested
            and self.controller.state.stage
            in {
                PipelineStage.COLLECTING_LINKS,
                PipelineStage.PROCESSING_PRODUCTS,
            },
        )
        self._set_button_enabled(self.browse_btn, idle)
        self._set_button_enabled(self.scan_btn, idle)
        self._set_button_enabled(self.conversion_browse_btn, conversion_ready)
        retry_merge = (
            bool(self.controller.checkpoints.pending)
            and not self.controller.is_running
            and not self.controller.shutdown_requested
        )
        self._set_button_enabled(self.conversion_start_btn, conversion_ready or retry_merge)

        normal_or_disabled = tk.NORMAL if idle else tk.DISABLED
        self.folder_entry.configure(state=normal_or_disabled)
        self.xpath_entry.configure(state=normal_or_disabled)
        self.max_pages_entry.configure(state=normal_or_disabled)
        self.pages_text.configure(state=normal_or_disabled)
        self.conversion_entry.configure(state=tk.NORMAL if conversion_ready else tk.DISABLED)
        for row in self.file_rows:
            row["skip_entry"].configure(state=tk.NORMAL if waiting else tk.DISABLED)
            row["keep_entry"].configure(state=tk.NORMAL if waiting else tk.DISABLED)
