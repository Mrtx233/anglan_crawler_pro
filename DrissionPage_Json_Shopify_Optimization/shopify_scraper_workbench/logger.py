# ======================== 日志层 ========================
#
# 职责: 把各业务模块的 print 输出安全转发到 GUI 日志文本框。
# 依赖: 标准库（tkinter 采用延迟导入，保证本模块可脱离 GUI 使用与单测）。
#
# 设计说明:
#   原实现在 main.py 启动时执行 sys.stdout = TextRedirector(widget)，
#   此后所有模块的 print 自动进日志框。拆分后这一机制完整保留，
#   业务模块无需改任何 print 语句。
# ========================================================

import sys
import threading

# 保存真实 stdout/stderr，供无 GUI 场景回退以及退出时还原
_REAL_STDOUT = sys.stdout
_REAL_STDERR = sys.stderr

# GUI 日志接收器（由 install_gui_sink 设置，clear_gui_sink 清除）
_gui_sink = None
_lock = threading.RLock()

class _LazyTk:
    """延迟导入 tkinter，使 logger 模块在无 GUI 环境下也能导入。

    仅暴露 TextRedirector 用到的属性；tkinter 不可用时提供等价常量，
    保证命令行/单测场景不会因 import 失败而崩溃。
    """

    _cache = None

    def _load(self):
        if self._cache is None:
            try:
                import tkinter as tk
                self._cache = tk
            except ImportError:
                class _Fallback:
                    NORMAL = "normal"
                    DISABLED = "disabled"
                    END = "end"

                    class TclError(Exception):
                        pass

                self._cache = _Fallback
        return self._cache

    def __getattr__(self, name):
        return getattr(self._load(), name)


tk = _LazyTk()


class TextRedirector:
    """将 print 输出安全地转发到日志文本框。"""

    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        if not string:
            return
        try:
            self.text_widget.after(0, self._write, string)
        except (tk.TclError, RuntimeError):
            pass

    def _write(self, string):
        try:
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, string)
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        except tk.TclError:
            pass

    def flush(self):
        pass


# ======================== sink 管理 ========================
def install_gui_sink(write_fn):
    """设置 GUI 日志接收器。write_fn 接收一段字符串。"""
    global _gui_sink
    with _lock:
        _gui_sink = write_fn


def clear_gui_sink():
    """清除 GUI 日志接收器（窗口销毁前调用）。"""
    global _gui_sink
    with _lock:
        _gui_sink = None


def log(msg=""):
    """主动输出一行日志。

    业务模块可直接用 print（会被 stdout 重定向捕获），
    本函数供不便用 print 的场景（如已还原 stdout 后）使用。
    """
    text = "" if msg is None else str(msg)
    with _lock:
        sink = _gui_sink
    if sink is not None:
        try:
            sink(text)
            return
        except Exception:
            pass
    try:
        _REAL_STDOUT.write(text if text.endswith("\n") else text + "\n")
    except Exception:
        pass


def log_exc(prefix=""):
    """打印当前异常堆栈，替代各处 import traceback + print_exc()。"""
    import traceback
    text = traceback.format_exc()
    if prefix:
        text = f"{prefix}{text}"
    print(text.rstrip("\n"))


def install_stdout_redirect(text_widget):
    """把 sys.stdout / sys.stderr 重定向到 GUI 文本框。

    与原实现 main.py 中的 sys.stdout = TextRedirector(self.log_text) 等价。
    """
    redirector = TextRedirector(text_widget)
    sys.stdout = redirector
    sys.stderr = redirector
    return redirector


def restore_stdout():
    """还原 sys.stdout / sys.stderr。程序退出前调用。"""
    sys.stdout = _REAL_STDOUT
    sys.stderr = _REAL_STDERR


def get_real_stdout():
    return _REAL_STDOUT


def get_real_stderr():
    return _REAL_STDERR
