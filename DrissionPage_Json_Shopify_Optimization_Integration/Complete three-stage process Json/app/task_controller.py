"""串行后台任务、线程安全事件队列和界面日志。"""
import contextlib
import queue
import threading
import traceback



class Cancelled(Exception):
    pass


def checkpoint(stop_event):
    """步骤边界的停止检查点：信号已置位时抛 Cancelled 结束当前任务。"""
    if stop_event.is_set():
        raise Cancelled('已在步骤边界停止，已完成的文件保留')


class TaskController:
    def __init__(self):
        """所有任务共用一个控制器：事件队列、停止信号、当前后台线程。"""
        self.events = queue.Queue()
        self.stop_event = threading.Event()
        self.thread = None
        self.stage = None
        # 当前运行器提供的“尽力中断当前浏览器请求”回调（可为 None）
        self._interrupt = None

    @property
    def running(self):
        """后台线程是否仍存活，用于判断任务是否真正结束。"""
        return self.thread is not None and self.thread.is_alive()

    def emit(self, kind, value):
        """线程安全地向界面投递一个 (kind, value) 事件。"""
        self.events.put((kind, value))

    def start(self, stage, operation):
        """在新的非守护线程中运行 operation(controller, log)。

        已有任务在跑时抛 RuntimeError；每次启动都会清空停止信号。
        """
        if self.running:
            raise RuntimeError('请等待当前任务结束')
        self.stage = stage
        self.stop_event.clear()
        self._interrupt = None
        self.thread = threading.Thread(target=self._run, args=(stage, operation), daemon=False)
        self.thread.start()

    def stop(self):
        """请求协作式停止，并在独立守护线程里尽力中断当前浏览器加载。"""
        self.stop_event.set()
        if self._interrupt:
            threading.Thread(target=self._interrupt, daemon=True).start()

    def _run(self, stage, operation):
        """后台线程主体：重定向 stdout/stderr 到日志，归一化任务状态并广播 done。"""
        result = None
        status = '完成'
        try:
            def log(text):
                """去掉行尾换行后投递一条界面日志（空串忽略）。"""
                line = str(text).rstrip('\n')
                if not line:
                    return
                self.emit('log', (stage, line))

            class Output:
                """把第三方库（DrissionPage/openpyxl）的 print 转发为日志行。"""
                def write(self, text):
                    """转写非空内容（纯空白行忽略，避免日志里出现大量空行）。"""
                    if text.strip():
                        log(text)
                def flush(self):
                    """占位实现：本输出对象无缓冲。"""
                    pass

            with contextlib.redirect_stdout(Output()), contextlib.redirect_stderr(Output()):
                try:
                    log(f'开始 {stage}')
                    result = operation(self, log)
                    status = result.get('status', '完成') if result else '完成'
                    if self.stop_event.is_set() and status == '完成':
                        status = '已停止'
                except Cancelled as exc:
                    status = '已停止'
                    log(str(exc))
                except Exception:
                    status = '失败'
                    log(traceback.format_exc())
                log(f'任务状态：{status}')
        except Exception as exc:
            status = '失败'
            self.emit('log', (stage, f'任务执行失败：{exc}'))
        finally:
            self._interrupt = None
            self.emit('done', (stage, status, result))
