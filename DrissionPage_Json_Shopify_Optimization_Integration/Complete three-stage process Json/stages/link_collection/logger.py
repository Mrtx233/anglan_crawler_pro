# ======================== 日志层 ========================
#
# 职责: 全项目统一的日志出口，替代原实现中散落的 print()。
# 依赖: 仅标准库（依赖图叶子，禁止导入其他业务模块）。
#
# 设计:
#   - 多 sink 广播: 任何模块调用 log()，由已注册的 sink 决定去向。
#   - 线程无关: 本模块不感知线程；跨线程安全由 sink 自行保证
#     （GUI sink 内部用 widget.after(0, ...) 转发到主线程）。
#   - TextRedirector 保留，仅用于兜底捕获第三方库
#     （DrissionPage / openpyxl）直接 print 的输出。
#   - 无 sink 时回退到真实 stdout，保证命令行/单测场景可用。
# ========================================================

import sys
import threading
import traceback
from typing import Callable, List

# sink 签名: Callable[[str], None]
_sinks: List[Callable[[str], None]] = []

# _sinks 的读写锁：采集线程与 GUI 主线程可能并发 add/remove/遍历
_lock = threading.RLock()

# 保存真实的 stdout，用于无 sink 时回退，以及退出时还原
_REAL_STDOUT = sys.stdout
_REAL_STDERR = sys.stderr


def add_sink(fn: Callable[[str], None]) -> None:
    """注册一个日志接收器（重复注册同一函数会被忽略）。"""
    if fn is None:
        return
    with _lock:
        if fn not in _sinks:
            _sinks.append(fn)


def remove_sink(fn: Callable[[str], None]) -> None:
    """注销日志接收器；未注册时静默忽略。"""
    with _lock:
        if fn in _sinks:
            _sinks.remove(fn)


def log(msg: str = "") -> None:
    """输出一行日志到全部已注册 sink；无 sink 时回退真实 stdout。

    自动补全换行（与 print 行为一致）。单个 sink 抛异常不影响其他 sink，
    避免 GUI 已销毁时采集线程崩溃。
    """
    text = "" if msg is None else str(msg)
    if not text.endswith("\n"):
        text += "\n"

    with _lock:
        targets = list(_sinks)

    if not targets:
        try:
            _REAL_STDOUT.write(text)
        except Exception:
            pass
        return

    for fn in targets:
        try:
            fn(text)
        except Exception:
            # sink 失效（如窗口已关闭）时忽略，保证采集主流程不中断
            pass


def log_exc(prefix: str = "") -> None:
    """打印当前异常的完整堆栈，替代各处 import traceback + print_exc()。"""
    text = traceback.format_exc()
    if prefix:
        text = f"{prefix}{text}"
    log(text.rstrip("\n"))


def install_stdout_redirect() -> None:
    """把 sys.stdout / sys.stderr 重定向到 log()。

    仅用于兜底捕获第三方库的 print 输出；本项目自身代码应直接调 log()。
    """
    redirector = TextRedirector()
    sys.stdout = redirector
    sys.stderr = redirector


def restore_stdout() -> None:
    """还原 sys.stdout / sys.stderr。程序退出前调用。"""
    sys.stdout = _REAL_STDOUT
    sys.stderr = _REAL_STDERR


class TextRedirector:
    """文件类对象，把写入内容转发给 log()。"""

    def write(self, string):
        """接收一次写入并按行转发给 log()，保持每行一条日志。"""
        if not string:
            return
        # 第三方库可能一次 write 多行，按行拆分保持日志粒度
        for line in str(string).splitlines():
            log(line)

    def flush(self):
        """占位实现：本重定向器无缓冲，无需刷新。"""
        pass

    def isatty(self):
        """始终返回 False，让第三方库按“非终端”路径输出。"""
        return False
