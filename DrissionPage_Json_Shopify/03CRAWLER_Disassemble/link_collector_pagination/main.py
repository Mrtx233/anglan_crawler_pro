# ======================== GUI 入口 ========================
#
# 职责: 只负责三件事——画界面、校验输入组装 CollectConfig、
#       起线程并把进度/日志/保存报告转发回界面。不含采集与文件逻辑。
# 依赖: config、logger、collector（唯一允许导入 tkinter 的模块）。
#
# 运行: cd link_collector_pagination && python main.py
# ========================================================

import os
import sys
import threading
import time

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import collector
import config
import logger

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
)

# 关窗等待保存的最长秒数；超时则强制退出并打日志
CLOSE_WAIT_TIMEOUT = 60
# 轮询采集线程的间隔（毫秒）
CLOSE_POLL_INTERVAL_MS = 200


class LinkCollectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title(config.APP_TITLE)
        self.root.geometry(config.WINDOW_GEOMETRY)
        self.root.minsize(*config.WINDOW_MIN_SIZE)
        self.root.configure(bg=config.UI_COLORS["background"])

        self.is_running = False
        self.stop_event = threading.Event()
        self.runner = None
        self._worker_thread = None
        self._closing = False

        self.max_pages_var = tk.StringVar(value="")
        self.xpath_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.save_summary_var = tk.StringVar(value="")

        self._build_ui()

        self._log_sink = self._append_log
        logger.add_sink(self._log_sink)
        logger.install_stdout_redirect()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ======================== 界面构建 ========================
    def _build_ui(self):
        colors = config.UI_COLORS
        ui_font = config.UI_FONT
        mono_font = config.MONO_FONT

        main = tk.Frame(self.root, bg=colors["background"])
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        # ── 标题 ──
        header = tk.Frame(main, bg=colors["background"])
        header.pack(fill=tk.X, pady=(0, 12))
        tk.Label(
            header, text=config.APP_TITLE,
            bg=colors["background"], fg=colors["text"],
            font=(ui_font, 18, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            header, text=config.APP_SUBTITLE,
            bg=colors["background"], fg=colors["text_secondary"],
            font=(ui_font, 9),
        ).pack(side=tk.LEFT, padx=(16, 0), pady=(8, 0))

        # ── 控制区 ──
        ctrl_card = tk.Frame(
            main, bg=colors["card"],
            highlightbackground=colors["border"],
            highlightthickness=1, bd=0,
        )
        ctrl_card.pack(fill=tk.X, pady=(0, 10))
        ctrl_inner = tk.Frame(ctrl_card, bg=colors["card"])
        ctrl_inner.pack(fill=tk.X, padx=16, pady=12)

        _entry_style = dict(
            bg=colors["card_alt"], fg=colors["text"],
            insertbackground=colors["primary"],
            font=(ui_font, 10), relief=tk.FLAT, bd=0,
            highlightthickness=1, highlightbackground=colors["border"],
            highlightcolor=colors["primary"],
        )

        # 第一行: 保存路径 + XPath + 最大页数
        row1 = tk.Frame(ctrl_inner, bg=colors["card"])
        row1.pack(fill=tk.X, pady=(0, 8))

        col_path = tk.Frame(row1, bg=colors["card"])
        col_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Label(
            col_path, text="保存路径",
            bg=colors["card"], fg=colors["text_secondary"],
            font=(ui_font, 9),
        ).pack(anchor=tk.W)
        path_row = tk.Frame(col_path, bg=colors["card"])
        path_row.pack(fill=tk.X)
        tk.Entry(
            path_row, textvariable=self.output_var, **_entry_style
        ).pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)
        tk.Button(
            path_row, text="浏览", command=self._browse_output,
            bg=colors["card_alt"], fg=colors["text_secondary"],
            activebackground="#EEF2F7", relief=tk.FLAT, bd=0,
            font=(ui_font, 9), padx=14, pady=4, cursor="hand2",
        ).pack(side=tk.LEFT, padx=(6, 0))

        col_xpath = tk.Frame(row1, bg=colors["card"])
        col_xpath.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        tk.Label(
            col_xpath, text="XPath",
            bg=colors["card"], fg=colors["text_secondary"],
            font=(ui_font, 9),
        ).pack(anchor=tk.W)
        tk.Entry(
            col_xpath, textvariable=self.xpath_var,
            **{**_entry_style, "font": (mono_font, 10)},
        ).pack(fill=tk.X, ipady=6)

        col_pages = tk.Frame(row1, bg=colors["card"])
        col_pages.pack(side=tk.LEFT, fill=tk.X, expand=False, padx=(0, 0))
        tk.Label(
            col_pages, text="最大页数",
            bg=colors["card"], fg=colors["text_secondary"],
            font=(ui_font, 9),
        ).pack(anchor=tk.W)
        tk.Entry(
            col_pages, textvariable=self.max_pages_var,
            width=10, **_entry_style,
        ).pack(fill=tk.X, ipady=6)

        # 第二行: 按钮 + 进度 + 保存摘要
        row2 = tk.Frame(ctrl_inner, bg=colors["card"])
        row2.pack(fill=tk.X)

        self.stop_btn = tk.Button(
            row2, text="停止", command=self._stop,
            bg=colors["card"], fg=colors["danger"],
            activebackground=colors["danger_light"],
            relief=tk.FLAT, bd=1, font=(ui_font, 10, "bold"),
            padx=20, pady=8, state=tk.DISABLED, cursor="hand2",
            highlightthickness=1, highlightbackground="#FECACA",
            highlightcolor="#FECACA",
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.start_btn = tk.Button(
            row2, text="开始采集", command=self._start,
            bg=colors["primary"], fg="#FFFFFF",
            activebackground=colors["primary_hover"],
            activeforeground="#FFFFFF",
            relief=tk.FLAT, bd=0, font=(ui_font, 10, "bold"),
            padx=28, pady=8, cursor="hand2",
        )
        self.start_btn.pack(side=tk.LEFT)

        self.progress_var = tk.StringVar(value="就绪")
        tk.Label(
            row2, textvariable=self.progress_var,
            bg=colors["card"], fg=colors["text_muted"],
            font=(ui_font, 9),
        ).pack(side=tk.LEFT, padx=(16, 0), pady=(8, 0))

        # 保存摘要：与进度同排，但在右侧；不抢焦点
        tk.Label(
            row2, textvariable=self.save_summary_var,
            bg=colors["card"], fg=colors["text_secondary"],
            font=(ui_font, 9),
        ).pack(side=tk.RIGHT, padx=(8, 0), pady=(8, 0))

        # ── 分类列表 ──
        pages_card = tk.Frame(
            main, bg=colors["card"],
            highlightbackground=colors["border"],
            highlightthickness=1, bd=0,
        )
        pages_card.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        pages_inner = tk.Frame(pages_card, bg=colors["card"])
        pages_inner.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)

        tk.Label(
            pages_inner, text="分类列表（每行: 标题, URL）",
            bg=colors["card"], fg=colors["text"],
            font=(ui_font, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            pages_inner, text=config.CATEGORY_PLACEHOLDER,
            bg=colors["card"], fg=colors["text_muted"],
            font=(ui_font, 9),
        ).pack(anchor=tk.W, pady=(2, 8))

        self.pages_text = scrolledtext.ScrolledText(
            pages_inner, height=8, wrap=tk.NONE,
            bg=colors["card_alt"], fg=colors["text"],
            insertbackground=colors["primary"],
            font=(mono_font, 9), relief=tk.FLAT, bd=0,
            highlightthickness=1, highlightbackground=colors["border"],
            highlightcolor=colors["primary"],
            padx=10, pady=8,
        )
        self.pages_text.pack(fill=tk.BOTH, expand=True)

        # ── 日志区 ──
        log_card = tk.Frame(
            main, bg=colors["log_background"],
            highlightbackground="#1E293B",
            highlightthickness=1, bd=0,
        )
        log_card.pack(fill=tk.BOTH, expand=True)

        log_header = tk.Frame(log_card, bg=colors["log_panel"], height=36)
        log_header.pack(fill=tk.X)
        log_header.pack_propagate(False)
        tk.Label(
            log_header, text="运行日志",
            bg=colors["log_panel"], fg="#F8FAFC",
            font=(ui_font, 10, "bold"),
        ).pack(side=tk.LEFT, padx=14, pady=8)
        tk.Button(
            log_header, text="清空", command=self._clear_log,
            bg="#1E293B", fg="#CBD5E1", activebackground="#334155",
            activeforeground="#FFFFFF", relief=tk.FLAT, bd=0,
            font=(ui_font, 8), padx=10, pady=4, cursor="hand2",
        ).pack(side=tk.RIGHT, padx=10, pady=6)

        self.log_text = scrolledtext.ScrolledText(
            log_card, state=tk.DISABLED,
            bg=colors["log_background"], fg=colors["log_text"],
            insertbackground="#FFFFFF", selectbackground="#1D4ED8",
            selectforeground="#FFFFFF",
            font=(mono_font, 9), relief=tk.FLAT, bd=0,
            padx=12, pady=10, wrap=tk.WORD,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ======================== 界面回调 ========================
    def _browse_output(self):
        path = filedialog.askdirectory(title="选择保存目录")
        if path:
            self.output_var.set(path)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _append_log(self, text):
        if not text:
            return
        try:
            self.root.after(0, self._write_log, text)
        except (tk.TclError, RuntimeError):
            pass

    def _write_log(self, text):
        try:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        except tk.TclError:
            pass

    def _on_progress(self, index, total, title):
        text = f"[{index}/{total}] {title[:30]}"
        try:
            self.root.after(0, self.progress_var.set, text)
        except (tk.TclError, RuntimeError):
            pass

    def _on_save_report(self, report):
        """采集线程的保存报告回调，转发到主线程更新摘要/弹窗。"""

        def _apply():
            text = report.describe()
            self.save_summary_var.set(text)

            # 关窗过程中不弹窗，只写日志；避免阻塞退出
            if self._closing:
                logger.log(f"[保存摘要] {text}")
                return

            if report.error:
                messagebox.showerror("保存失败", text)
                return

            recovery = list(report.recovery_files) + list(
                report.json_recovery_files
            )
            if recovery:
                details = [text]
                if report.recovery_files:
                    details.append(
                        "恢复文件:\n" + "\n".join(report.recovery_files)
                    )
                if report.json_recovery_files:
                    details.append(
                        "索引备份:\n" + "\n".join(report.json_recovery_files)
                    )
                messagebox.showwarning(
                    "保存完成（含恢复文件）", "\n\n".join(details)
                )
            else:
                logger.log(f"[保存摘要] {text}")

        try:
            self.root.after(0, _apply)
        except (tk.TclError, RuntimeError):
            pass

    # ======================== 输入校验与组装 ========================
    def _build_config(self):
        raw_text = self.pages_text.get("1.0", tk.END)
        categories = collector.parse_categories(raw_text)
        if not categories:
            messagebox.showerror("错误", "请先填写分类列表")
            return None

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = os.path.join(
                PROJECT_ROOT, config.DEFAULT_OUTPUT_DIRNAME
            )
            self.output_var.set(output_path)

        max_pages_raw = self.max_pages_var.get().strip()
        if max_pages_raw:
            try:
                max_pages = int(max_pages_raw)
                if max_pages <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "错误", "最大页数必须为正整数，或留空表示不限制"
                )
                return None
        else:
            max_pages = 0

        return config.CollectConfig(
            categories=categories,
            output_dir=output_path,
            max_pages=max_pages,
            xpath=self.xpath_var.get().strip(),
        )

    # ======================== 任务控制 ========================
    def _start(self):
        if self.is_running:
            return

        cfg = self._build_config()
        if cfg is None:
            return

        self.is_running = True
        self._closing = False
        self.stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_var.set("正在启动...")
        self.save_summary_var.set("")
        self._clear_log()

        self.runner = collector.CollectorRunner(
            cfg,
            stop_event=self.stop_event,
            on_progress=self._on_progress,
            on_save_report=self._on_save_report,
        )
        self._worker_thread = threading.Thread(
            target=self._worker, daemon=True, name="collector-worker"
        )
        self._worker_thread.start()

    def _worker(self):
        try:
            self.runner.run()
        finally:
            self.is_running = False
            if not self._closing:
                try:
                    self.root.after(0, self._reset_buttons)
                except (tk.TclError, RuntimeError):
                    pass

    def _stop(self):
        if not self.is_running or self.runner is None:
            return
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_var.set("正在停止并保存已采集结果...")
        logger.log("\n[停止] 已放弃当前页，正在保存此前已采集的链接...")
        threading.Thread(target=self.runner.interrupt, daemon=True).start()

    def _reset_buttons(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress_var.set("就绪")
        # 保存摘要保留，供用户查看

    # ======================== 关闭与退出 ========================
    def _on_close(self):
        """关闭窗口：运行中先确认，再请求停止并等待保存完成后销毁。"""
        if self._closing:
            return

        if self.is_running:
            if not messagebox.askyesno(
                "确认退出",
                "采集任务正在进行中。\n"
                "退出将停止采集并保存此前已完成分页的结果，"
                "当前正在加载的页面会被放弃。\n\n确定退出吗？",
            ):
                return

            self._closing = True
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)
            self.progress_var.set("正在停止并保存，请勿关闭窗口...")
            logger.log("\n[关闭] 正在停止采集并等待保存完成...")

            try:
                self.runner.interrupt()
            except Exception:
                pass

            self._wait_worker_then_shutdown(
                deadline=time.monotonic() + CLOSE_WAIT_TIMEOUT
            )
            return

        self._shutdown()

    def _wait_worker_then_shutdown(self, deadline):
        """轮询等待采集线程结束，再销毁窗口；主线程保持响应。"""
        thread = self._worker_thread
        if thread is None or not thread.is_alive():
            self._shutdown()
            return

        if time.monotonic() > deadline:
            logger.log(
                f"[警告] 等待保存超过 {CLOSE_WAIT_TIMEOUT} 秒，强制退出，"
                "可能丢失未落盘数据"
            )
            self._shutdown()
            return

        try:
            self.root.after(
                CLOSE_POLL_INTERVAL_MS,
                lambda: self._wait_worker_then_shutdown(deadline),
            )
        except (tk.TclError, RuntimeError):
            self._shutdown()

    def _shutdown(self):
        """还原 stdout 并注销日志 sink，避免窗口销毁后写入失效控件。"""
        try:
            logger.remove_sink(self._log_sink)
        finally:
            logger.restore_stdout()
        try:
            self.root.destroy()
        except tk.TclError:
            pass


def main():
    root = tk.Tk()
    LinkCollectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()