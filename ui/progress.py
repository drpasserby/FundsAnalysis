# -*- coding: utf-8 -*-
"""Tk 进度窗口。

只允许主线程调用（worker 线程请用 ui.worker.ProgressProxy）。
重复调用同一种模式是幂等的，便于被 80ms 轮询驱动。
"""
import tkinter as tk
from tkinter import ttk


class ProgressWindow(object):
    def __init__(self, parent=None, title='处理中', text='准备中…'):
        self.closed = False
        self.mode = 'determinate'
        self.win = tk.Toplevel(parent) if parent is not None else tk.Tk()
        self.win.title(title)
        self.win.resizable(False, False)
        try:
            self.win.attributes('-topmost', True)
        except Exception:
            pass
        try:
            self.win.protocol('WM_DELETE_WINDOW', lambda: None)  # 处理中禁止关闭
        except Exception:
            pass

        self.text_var = tk.StringVar(value=text)
        self.percent_var = tk.StringVar(value='')
        body = ttk.Frame(self.win, padding=16)
        body.pack(fill='both', expand=True)
        ttk.Label(body, textvariable=self.text_var, width=46, anchor='w').pack(fill='x')
        self.bar = ttk.Progressbar(body, mode='determinate', maximum=1, value=0, length=420)
        self.bar.pack(fill='x', pady=(10, 4))
        ttk.Label(body, textvariable=self.percent_var, anchor='e').pack(fill='x')
        self._center(parent)
        self._pump()

    # ---- 内部 ----
    def _center(self, parent):
        try:
            self.win.update_idletasks()
            if parent is not None:
                px, py = parent.winfo_rootx(), parent.winfo_rooty()
                pw, ph = parent.winfo_width(), parent.winfo_height()
            else:
                px = py = 0
                pw, ph = self.win.winfo_screenwidth(), self.win.winfo_screenheight()
            w, h = self.win.winfo_width(), self.win.winfo_height()
            x = px + max(0, (pw - w) // 2)
            y = py + max(0, (ph - h) // 3)
            self.win.geometry('+{}+{}'.format(x, y))
        except Exception:
            pass

    def _pump(self):
        if self.closed:
            return
        try:
            self.win.update()
        except Exception:
            pass

    # ---- 进度接口（与 core.progress.Progress 一致）----
    def set_indeterminate(self, text=''):
        if self.closed:
            return
        if text:
            self.text_var.set(text)
        if self.mode != 'indeterminate':
            self.mode = 'indeterminate'
            try:
                self.bar.stop()
                self.bar.configure(mode='indeterminate')
                self.bar.start(12)
            except Exception:
                pass
            self.percent_var.set('')
        self._pump()

    def set_determinate(self, maximum, text=''):
        if self.closed:
            return
        if text:
            self.text_var.set(text)
        try:
            if self.mode != 'determinate':
                self.bar.stop()
                self.bar.configure(mode='determinate')
                self.mode = 'determinate'
            self.bar.configure(maximum=max(1, int(maximum or 1)))
        except Exception:
            pass
        self._pump()

    def update(self, current, maximum=None, text=None):
        if self.closed:
            return
        if maximum is not None:
            self.set_determinate(maximum)
        if self.mode != 'determinate':
            return
        if text:
            self.text_var.set(text)
        try:
            top = float(self.bar['maximum'] or 1)
            cur = min(max(0.0, float(current)), top)
            self.bar['value'] = cur
            self.percent_var.set('{}%'.format(int(cur * 100 / top)))
        except Exception:
            pass
        self._pump()

    def close(self):
        if self.closed:
            return
        self.closed = True
        try:
            self.bar.stop()
        except Exception:
            pass
        try:
            self.win.destroy()
        except Exception:
            pass
