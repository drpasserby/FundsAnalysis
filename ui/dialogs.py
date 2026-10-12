# -*- coding: utf-8 -*-
"""共享对话框：选择文件/目录、错误提示、结果面板。"""
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

FT_FLOW = [('流水表格', '*.xlsx *.xls *.tsv *.txt *.csv'), ('Excel 文件', '*.xlsx *.xls'),
           ('文本流水', '*.tsv *.txt *.csv'), ('所有文件', '*.*')]
FT_WECHAT = [('微信账单文本', '*.txt'), ('所有文件', '*.*')]
FT_XLSX = [('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')]
FT_TSV = [('统一流水', '*.tsv *.txt *.csv'), ('所有文件', '*.*')]


def ask_open_file(parent, title='选择文件', filetypes=None, initialdir=None):
    return filedialog.askopenfilename(parent=parent, title=title,
                                      filetypes=filetypes or FT_FLOW, initialdir=initialdir) or None


def ask_open_files(parent, title='选择文件（可多选）', filetypes=None, initialdir=None):
    return list(filedialog.askopenfilenames(parent=parent, title=title,
                                            filetypes=filetypes or FT_FLOW, initialdir=initialdir) or [])


def ask_directory(parent, title='选择目录', initialdir=None):
    return filedialog.askdirectory(parent=parent, title=title, initialdir=initialdir) or None


def show_info(parent, message, title='提示'):
    messagebox.showinfo(title, message, parent=parent)


def show_error(parent, title, err, trace=''):
    msg = '{}'.format(err)
    if trace:
        msg += '\n\n技术细节（已写入运行日志）：\n' + trace.strip().splitlines()[-1]
    messagebox.showerror(title, msg, parent=parent)


def open_path(path):
    """在系统文件管理器中打开文件或目录。"""
    if not path or not os.path.exists(path):
        return False
    try:
        if hasattr(os, 'startfile'):
            os.startfile(path)  # noqa - Windows
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
        return True
    except Exception:
        return False


def ask_columns(parent, columns, title='选择需要写入的列', locked=None, note='',
                selected=None):
    """勾选列对话框（返回选中的列名列表；用户取消返回 None）。

    columns ：全部候选列（按展示顺序）
    locked  ：强制勾选且不可取消的列（如前 4 列必需列）
    selected：默认勾选的列；None 表示除 locked 外全部默认勾选
    """
    columns = list(columns)
    locked = list(locked or [])
    result = {'ok': False, 'cols': []}

    dlg = tk.Toplevel(parent) if parent is not None else tk.Tk()
    dlg.title(title)
    try:
        dlg.attributes('-topmost', True)
        dlg.after(300, lambda: dlg.attributes('-topmost', False))
    except Exception:
        pass

    if note:
        ttk.Label(dlg, text=note, wraplength=420, justify='left',
                  foreground='#555555').pack(fill='x', padx=12, pady=(10, 0))
    ttk.Label(dlg, text='请勾选需要写入的列（前 4 列是必需列，始终保留）：').pack(
        anchor='w', padx=12, pady=(8, 4))

    box = ttk.Frame(dlg)
    box.pack(fill='both', expand=True, padx=12)
    canvas = tk.Canvas(box, highlightthickness=0, height=280)
    sb = ttk.Scrollbar(box, orient='vertical', command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
    canvas.create_window((0, 0), window=inner, anchor='nw')
    canvas.configure(yscrollcommand=sb.set)
    canvas.pack(side='left', fill='both', expand=True)
    sb.pack(side='right', fill='y')

    vars_ = []
    for c in columns:
        is_locked = c in locked
        default = True if selected is None else (c in selected)
        v = tk.BooleanVar(value=True if is_locked else bool(default))
        cb = ttk.Checkbutton(inner, text=c + ('（必需）' if is_locked else ''), variable=v)
        if is_locked:
            cb.state(['disabled'])
        cb.pack(anchor='w', fill='x')
        vars_.append((c, v, is_locked))

    def set_all(flag):
        for _c, v, is_locked in vars_:
            if not is_locked:
                v.set(flag)

    def on_ok():
        picked = [c for c, v, _l in vars_ if v.get()]
        if not picked:
            messagebox.showwarning('提示', '至少需要保留一列。', parent=dlg)
            return
        result['ok'] = True
        result['cols'] = picked
        dlg.destroy()

    btns = ttk.Frame(dlg)
    btns.pack(fill='x', padx=12, pady=10)
    ttk.Button(btns, text='全选', command=lambda: set_all(True)).pack(side='left')
    ttk.Button(btns, text='全不选', command=lambda: set_all(False)).pack(side='left', padx=6)
    ttk.Button(btns, text='取消', command=dlg.destroy).pack(side='right')
    ttk.Button(btns, text='确定', command=on_ok).pack(side='right', padx=6)
    dlg.bind('<Escape>', lambda e: dlg.destroy())

    center(dlg, parent)
    try:
        dlg.grab_set()
    except Exception:
        pass
    dlg.wait_window()
    return result['cols'] if result['ok'] else None


def center(win, parent=None):
    try:
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        if parent is not None:
            x = parent.winfo_rootx() + max(0, (parent.winfo_width() - w) // 2)
            y = parent.winfo_rooty() + max(0, (parent.winfo_height() - h) // 3)
        else:
            x = (win.winfo_screenwidth() - w) // 2
            y = (win.winfo_screenheight() - h) // 3
        win.geometry('+{}+{}'.format(x, y))
    except Exception:
        pass


class ResultDialog(object):
    """任务完成后的结果面板：列出产物并可直接打开目录 / 日志 / 进入下一步。"""

    def __init__(self, parent, title='完成', summary='', items=None, out_dir=None,
                 log_file=None, next_label=None, on_next=None, extra_buttons=None):
        self.parent = parent
        self.on_next = on_next
        self.out_dir = out_dir
        self.log_file = log_file
        self.win = tk.Toplevel(parent) if parent is not None else tk.Tk()
        self.win.title(title)
        self.win.resizable(True, True)
        try:
            self.win.attributes('-topmost', True)
        except Exception:
            pass

        body = ttk.Frame(self.win, padding=14)
        body.pack(fill='both', expand=True)

        if summary:
            ttk.Label(body, text=summary, justify='left', wraplength=560).pack(fill='x')

        box = ttk.Frame(body)
        box.pack(fill='both', expand=True, pady=(10, 8))
        self.text = tk.Text(box, width=76, height=max(4, min(12, len(items or []) + 2)),
                            wrap='none', relief='solid', borderwidth=1)
        self.text.pack(side='left', fill='both', expand=True)
        sb = ttk.Scrollbar(box, orient='vertical', command=self.text.yview)
        sb.pack(side='right', fill='y')
        self.text.configure(yscrollcommand=sb.set)
        lines = []
        for label, path in (items or []):
            lines.append('{}：{}'.format(label, path))
        if log_file:
            lines.append('运行日志：{}'.format(log_file))
        self.text.insert('1.0', '\n'.join(lines) if lines else '（无产物）')
        self.text.configure(state='disabled')

        bar = ttk.Frame(body)
        bar.pack(fill='x')
        if out_dir:
            ttk.Button(bar, text='打开输出目录', command=lambda: open_path(out_dir)).pack(side='left')
        if log_file:
            ttk.Button(bar, text='打开运行日志', command=lambda: open_path(log_file)).pack(side='left', padx=6)
        ttk.Button(bar, text='复制路径', command=self._copy).pack(side='left')
        for label, cb in (extra_buttons or []):
            ttk.Button(bar, text=label, command=cb).pack(side='left', padx=6)
        ttk.Button(bar, text='关闭', command=self.win.destroy).pack(side='right')
        if next_label and on_next:
            ttk.Button(bar, text=next_label, command=self._next).pack(side='right', padx=6)

        self.win.bind('<Escape>', lambda e: self.win.destroy())
        center(self.win, parent)
        try:
            self.win.transient(parent) if parent is not None else None
            self.win.focus_set()
        except Exception:
            pass

    def _copy(self):
        try:
            self.win.clipboard_clear()
            self.win.clipboard_append(self.text.get('1.0', 'end').strip())
        except Exception:
            pass

    def _next(self):
        self.win.destroy()
        if self.on_next:
            self.on_next()
