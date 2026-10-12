# -*- coding: utf-8 -*-
"""后台线程执行 + 主线程刷新进度。

Tkinter 不是线程安全的：worker 只把进度写进共享字典（只保留最新值，天然节流），
主线程用 after 轮询刷新窗口并在结束时回调。
"""
import threading
import traceback

POLL_MS = 80


class ProgressProxy(object):
    """worker 线程侧的进度代理（与 core.progress.Progress 接口一致）。"""

    def __init__(self, state):
        self._state = state

    def set_indeterminate(self, text=''):
        self._state['mode'] = 'indeterminate'
        if text:
            self._state['text'] = text

    def set_determinate(self, maximum, text=''):
        self._state['mode'] = 'determinate'
        self._state['max'] = max(1, int(maximum or 1))
        self._state['value'] = 0
        if text:
            self._state['text'] = text

    def update(self, current, maximum=None, text=None):
        if maximum is not None:
            self._state['max'] = max(1, int(maximum))
        self._state['value'] = current
        if text:
            self._state['text'] = text

    def close(self):
        pass


class TaskRunner(object):
    """把 func(progress) 丢到后台线程跑，完成后在主线程回调 on_done / on_error。"""

    def __init__(self, root):
        self.root = root
        self._state = {}
        self._result = {}
        self._done = threading.Event()
        self._running = False

    def busy(self):
        return self._running

    def start(self, func, on_done=None, on_error=None, progress=None):
        if self._running:
            return False
        self._running = True
        self._state = {'mode': 'determinate', 'max': 1, 'value': 0, 'text': '准备中…'}
        self._result = {}
        self._done = threading.Event()
        proxy = ProgressProxy(self._state)
        if progress is not None:
            progress.set_determinate(1, '准备中…')

        def worker():
            try:
                self._result['value'] = func(proxy)
                self._result['error'] = None
            except Exception as e:  # 任何异常都回传给 UI，不让线程静默死掉
                self._result['error'] = e
                self._result['trace'] = traceback.format_exc()
            finally:
                self._done.set()

        threading.Thread(target=worker, daemon=True).start()
        self._poll(on_done, on_error, progress)
        return True

    def _poll(self, on_done, on_error, progress):
        st = self._state
        if progress is not None and not getattr(progress, 'closed', False):
            if st.get('mode') == 'indeterminate':
                progress.set_indeterminate(st.get('text') or '')
            else:
                progress.update(st.get('value', 0), st.get('max', 1), st.get('text'))
        if self._done.is_set():
            if progress is not None:
                progress.close()
            self._running = False
            err = self._result.get('error')
            if err is not None:
                if on_error:
                    on_error(err, self._result.get('trace', ''))
            elif on_done:
                on_done(self._result.get('value'))
            return
        self.root.after(POLL_MS, lambda: self._poll(on_done, on_error, progress))
