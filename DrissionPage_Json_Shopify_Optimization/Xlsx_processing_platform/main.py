"""Excel 处理平台：一个文件夹入口，后台完成合并与导出。"""

import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from processing_runner import process_folder


class ProcessingApp:
    def __init__(self, root):
        self.root = root
        self.events = queue.Queue()
        self.running = False
        self.closing = False
        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="就绪")
        root.title("Excel 数据处理平台")
        root.geometry("780x460")
        root.minsize(620, 360)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        body = ttk.Frame(root, padding=20)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="Excel 数据处理平台", font=("TkDefaultFont", 18, "bold")).pack(anchor=tk.W, pady=(0, 18))
        path_row = ttk.Frame(body)
        path_row.pack(fill=tk.X)
        ttk.Label(path_row, text="数据文件夹").pack(side=tk.LEFT, padx=(0, 10))
        self.path_entry = ttk.Entry(path_row, textvariable=self.folder_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.browse_btn = ttk.Button(path_row, text="选择文件夹", command=self._browse)
        self.browse_btn.pack(side=tk.LEFT, padx=(10, 0))

        actions = ttk.Frame(body)
        actions.pack(fill=tk.X, pady=14)
        ttk.Label(actions, textvariable=self.status_var).pack(side=tk.LEFT)
        self.start_btn = ttk.Button(actions, text="合并并转换", command=self._start)
        self.start_btn.pack(side=tk.RIGHT)
        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(0, 12))
        self.log_text = scrolledtext.ScrolledText(body, state=tk.DISABLED, wrap=tk.WORD, height=12)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.path_entry.focus_set()
        root.after(100, self._poll_events)

    def _browse(self):
        folder = filedialog.askdirectory(title="选择数据文件夹")
        if folder:
            self.folder_var.set(folder)

    def _start(self):
        if self.running:
            return
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showerror("缺少路径", "请选择数据文件夹。")
            return
        self.running = True
        self.status_var.set("正在处理…")
        for widget in (self.path_entry, self.browse_btn, self.start_btn):
            widget.configure(state=tk.DISABLED)
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.progress.start(12)
        # 工作线程不操作 Tk；关闭窗口时等待完成，避免中断写文件。
        threading.Thread(target=self._worker, args=(folder,), daemon=False).start()

    def _worker(self, folder):
        try:
            report = process_folder(folder, lambda text: self.events.put(("log", str(text))))
            self.events.put(("done", report))
        except Exception as exc:
            self.events.put(("log", traceback.format_exc()))
            self.events.put(("error", str(exc)))

    def _poll_events(self):
        while True:
            try:
                kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "log":
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, value + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
            else:
                self.running = False
                self.progress.stop()
                for widget in (self.path_entry, self.browse_btn, self.start_btn):
                    widget.configure(state=tk.NORMAL)
                if kind == "done":
                    self.status_var.set(f"完成 · {value.input_files} 个文件 · {value.product_rows} 条商品")
                else:
                    self.status_var.set("处理失败")
                    messagebox.showerror("处理失败", value)
                    # 失败时留在窗口，用户可以检查日志和已有输出。
                    self.closing = False
        if self.closing and not self.running:
            self.root.destroy()
            return
        self.root.after(100, self._poll_events)

    def _on_close(self):
        if self.running:
            self.closing = True
            self.status_var.set("处理完成后关闭…")
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    ProcessingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
