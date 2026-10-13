# -*- coding: utf-8 -*-
"""进度回调协议：与 UI 解耦，便于 headless 运行与测试。"""


class Progress(object):
    """进度接口。UI 实现见 ui/progress.py，worker 代理见 ui/worker.py。"""

    def set_indeterminate(self, text=''):
        pass

    def set_determinate(self, maximum, text=''):
        pass

    def update(self, current, maximum=None, text=None):
        pass

    def close(self):
        pass


class NullProgress(Progress):
    """无界面运行时的空实现；传入 logger 时按 10% 节流记日志。"""

    def __init__(self, logger=None):
        self.logger = logger
        self._last = -1

    def set_indeterminate(self, text=''):
        if self.logger and text:
            self.logger(text)

    def set_determinate(self, maximum, text=''):
        self._last = -1
        if self.logger and text:
            self.logger(text)

    def update(self, current, maximum=None, text=None):
        if not self.logger:
            return
        if maximum:
            pct = int(current * 100 / max(1, maximum))
            if pct == self._last:
                return
            self._last = pct
            self.logger('{} {}/{} ({}%)'.format(text or '进度', current, maximum, pct))
        elif text:
            self.logger(text)
