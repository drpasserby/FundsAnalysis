# -*- coding: utf-8 -*-
"""
资金流水工具套件 v2.0.0v1
作者：wulvxinchen
"""
import os
import sys
import traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

import tkinter as tk
from tkinter import ttk

from core import APP_NAME, AUTHOR, GITHUB, __version__
from core import deps
from core import log as logger
from core import paths
from core import registry
from ui import dialogs
from ui.progress import ProgressWindow
from ui.worker import TaskRunner

BG = '#f4f5f7'
CARD_BG = '#ffffff'
FG_DIM = '#5a5f66'
FG_NOTE = '#8a6d3b'
FONT_UI = 'Microsoft YaHei UI'


class SuiteApp(object):
    """主窗口：四个功能卡片 + 一键跑全程。"""

    def __init__(self, root):
        self.root = root
        self.state = {'last_out_dir': '', 'last_input': '', 'last_style': 'classic',
                      'last_policy': 'spend'}
        root.title('{} v{}'.format(APP_NAME, __version__))
        root.configure(bg=BG)
        self._header()
        self._cards()
        self._bottom()
        self.root.after(150, self._check_env)

    # ---------- 环境自检 ----------
    def _check_env(self):
        """启动自检：缺依赖时把原因和解决办法直接说清楚，避免"点哪儿都报错"。"""
        miss = deps.missing()
        if miss:
            self.status.config(text='缺少运行依赖：{}（点「环境自检」看详情与解决办法）'.format('、'.join(miss)),
                               fg='#c0392b')
            logger.log('依赖自检未通过：缺少 {}；{}'.format('、'.join(miss), deps.interpreter_info()))
            dialogs.show_error(self.root, '缺少运行依赖', deps.advice())
            return False
        opt = deps.missing_optional()
        tip = '就绪（Python {}，依赖齐全）'.format(deps.python_version())
        if opt:
            tip += '；可选依赖 {} 缺失（不影响 .xlsx）'.format('、'.join(opt))
        self.status.config(text=tip)
        logger.log('依赖自检通过：Python {}'.format(deps.python_version()))
        return True

    # ---------- 界面 ----------
    def _header(self):
        head = tk.Frame(self.root, bg=BG)
        head.pack(fill='x', padx=18, pady=(16, 4))
        tk.Label(head, text=APP_NAME, font=(FONT_UI, 17, 'bold'), bg=BG).pack(anchor='w')
        tk.Label(head, text='版本 v{}    ·    {}    ·    GPLv3    ·    {}'.format(
            __version__, AUTHOR, GITHUB), bg=BG, fg=FG_DIM, font=('', 9)).pack(anchor='w', pady=(3, 0))
        tk.Label(head, text='四个功能互不干扰、可单独使用；产物按固定命名落在同一个输出目录，'
                            '下一步会自动读取。建议顺序：A → B → C → D。',
                 bg=BG, fg=FG_DIM, font=('', 9), justify='left').pack(anchor='w', pady=(6, 0))

    def _cards(self):
        wrap = tk.Frame(self.root, bg=BG)
        wrap.pack(fill='both', expand=True, padx=18, pady=8)
        for spec in registry.TOOLS:
            self._card(wrap, spec.title, spec.subtitle,
                       lambda s=spec: self._open(s.id), enabled=spec.enabled, note=spec.note)
        try:
            from tools.alipay import NOTE, STATUS
        except Exception:
            STATUS, NOTE = '未启用', ''
        self._card(wrap, '支付宝导入（预留）',
                   '{}：{}'.format(STATUS, NOTE), None, enabled=False, note='')

    def _card(self, parent, title, subtitle, command, enabled=True, note=''):
        card = tk.Frame(parent, bg=CARD_BG, highlightbackground='#dcdfe4', highlightthickness=1)
        card.pack(fill='x', pady=4)
        left = tk.Frame(card, bg=CARD_BG)
        left.pack(side='left', fill='both', expand=True, padx=12, pady=10)
        tk.Label(left, text=title, bg=CARD_BG, font=(FONT_UI, 11, 'bold'), anchor='w').pack(anchor='w')
        tk.Label(left, text=subtitle, bg=CARD_BG, fg=FG_DIM, font=('', 9), justify='left',
                 wraplength=520, anchor='w').pack(anchor='w', pady=(3, 0))
        if note:
            tk.Label(left, text=note, bg=CARD_BG, fg=FG_NOTE, font=('', 9), justify='left',
                     wraplength=520, anchor='w').pack(anchor='w', pady=(3, 0))
        btn = ttk.Button(card, text='进入' if enabled else '未启用', command=command)
        btn.pack(side='right', padx=12)
        btn.state(['!disabled'] if enabled else ['disabled'])
        return card

    def _bottom(self):
        bar = tk.Frame(self.root, bg=BG)
        bar.pack(fill='x', padx=18, pady=(4, 12))
        ttk.Button(bar, text='一键跑全程（A→B→C）…', command=self._pipeline_window).pack(side='left')
        ttk.Button(bar, text='打开最近输出目录', command=self._open_recent).pack(side='left', padx=6)
        ttk.Button(bar, text='关于 / 使用说明', command=self._about).pack(side='left')
        ttk.Button(bar, text='环境自检', command=self._env_report).pack(side='left', padx=6)
        self.status = tk.Label(self.root, text='就绪', bg=BG, fg=FG_DIM, anchor='w', font=('', 9))
        self.status.pack(fill='x', padx=18, pady=(0, 10))

    # ---------- 行为 ----------
    def _open(self, tid):
        try:
            registry.load_tool(tid).open_window(self.root, self.state)
        except ImportError as e:
            logger.log('打开功能 {} 失败（缺少依赖）：{}'.format(tid, e))
            dialogs.show_error(self.root, '缺少运行依赖', deps.advice('原始错误：{}'.format(e)))
        except Exception as e:
            trace = traceback.format_exc()
            logger.log('打开功能 {} 失败：{}'.format(tid, e))
            logger.log(trace)
            dialogs.show_error(self.root, '打开失败', e, trace)

    def _open_recent(self):
        d = self.state.get('last_out_dir') or paths.find_recent_output_dir(paths.default_output_base())
        if d and os.path.isdir(d):
            self.state['last_out_dir'] = d
            dialogs.open_path(d)
            self.status.config(text='已打开：{}'.format(d))
        else:
            dialogs.show_info(self.root, '还没有找到输出目录（形如 输出_20250101_1200）。', '提示')

    def _env_report(self):
        dialogs.show_info(self.root,
                          deps.diagnostic_text() + '\n套件版本：v{}'.format(__version__),
                          '环境自检')

    def _about(self):
        miss = deps.missing()
        dep_state = '齐全' if not miss else '缺少 {}（见上方提示的处理办法）'.format('、'.join(miss))
        msg = ('使用顺序（其中 A/B/C 可一键跑全程）：\n'
               '  A 导入与清洗：选择微信账单 TenpayTrades.txt（或其所在文件夹），'
               '生成统一格式的 01_统一流水.tsv 与 01_清洗报告.txt。\n'
               '  B 多来源合并：选择一份或多份统一格式流水，合并去重后生成 02_合并流水.tsv。\n'
               '  C 流水总表与透视表：生成 03_流水总表.xlsx / .csv 与 03_数据透视表.xlsx'
               '（含数据透视表、按主体汇总、人物关系图三张表）。\n'
               '  D 资金流向分析：生成 04_资金流向图与流水分析报告.html'
               '（先看资金流向图，再由按钮进入分析报告）。\n\n'
               '输出：默认写入输入文件所在目录下的「输出_时间戳」文件夹；所有产物集中在同一目录，'
               '下一步会自动读取，也可以手动选择。\n\n'
               '数据格式契约：前 4 列固定为「用户方 | 支出/收入 | 客户方 | 金额（元）」，'
               '其余列（交易时间、交易类型、订单号等）为可选扩展列。\n\n'
               '免责声明：本工具的分析结果仅供参考，不构成审计结论、鉴定意见或法律意见。\n\n'
               '{}\n依赖状态：{}\n'
               '作者：{}    {}'.format(deps.interpreter_info().replace('\n', '；'),
                                    dep_state, AUTHOR, GITHUB))
        dialogs.show_info(self.root, msg, '关于与使用说明')

    def _pipeline_window(self):
        win = tk.Toplevel(self.root)
        win.title('一键跑全程 · 导入 → 合并 → 总表与透视')
        win.resizable(False, False)
        try:
            win.transient(self.root)
        except Exception:
            pass
        body = ttk.Frame(win, padding=14)
        body.pack(fill='both', expand=True)

        ttk.Label(body, text='微信账单文件或文件夹（TenpayTrades.txt；可添加多个）',
                  font=('', 10, 'bold')).pack(anchor='w')
        listbox = tk.Listbox(body, width=76, height=6, activestyle='none')
        listbox.pack(fill='both', expand=True, pady=(4, 6))
        selected = []

        def refresh():
            listbox.delete(0, 'end')
            for p in selected:
                listbox.insert('end', p)

        def add_files():
            for p in dialogs.ask_open_files(win, '选择微信账单 txt（可多选）', dialogs.FT_WECHAT):
                if p not in selected:
                    selected.append(p)
            refresh()

        def add_folder():
            d = dialogs.ask_directory(win, '选择包含账单的文件夹（递归查找）', self.state.get('last_dir'))
            if d and d not in selected:
                selected.append(d)
                self.state['last_dir'] = d
                refresh()

        def clear():
            del selected[:]
            refresh()

        row = ttk.Frame(body)
        row.pack(fill='x')
        ttk.Button(row, text='添加账单文件…', command=add_files).pack(side='left')
        ttk.Button(row, text='添加文件夹…', command=add_folder).pack(side='left', padx=6)
        ttk.Button(row, text='清空', command=clear).pack(side='left')

        opt = ttk.LabelFrame(body, text='选项', padding=8)
        opt.pack(fill='x', pady=10)
        var_self = tk.BooleanVar(value=True)
        var_dedupe = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text='去除自转自流水（规则 R1）', variable=var_self).pack(anchor='w')
        ttk.Checkbutton(opt, text='合并时去除重复记录（规则 R5）', variable=var_dedupe).pack(anchor='w')

        ttk.Label(body, text='输出目录（留空则按第一个输入自动新建 输出_时间戳）').pack(anchor='w')
        var_out = tk.StringVar(value=self.state.get('last_out_dir') or '')
        row2 = ttk.Frame(body)
        row2.pack(fill='x', pady=(4, 10))
        ttk.Entry(row2, textvariable=var_out, width=56).pack(side='left', fill='x', expand=True)
        ttk.Button(row2, text='浏览…', command=lambda: var_out.set(
            dialogs.ask_directory(win, '选择输出目录', var_out.get() or None) or var_out.get())
        ).pack(side='left', padx=(6, 0))

        runner = TaskRunner(win)

        def start():
            if not selected:
                dialogs.show_error(win, '未选择输入', '请先添加微信账单文件或文件夹。')
                return
            if runner.busy():
                return
            files = [p for p in selected if os.path.isfile(p)]
            folders = [p for p in selected if os.path.isdir(p)]
            out_dir = var_out.get().strip().strip('"') or None
            pw = ProgressWindow(win, title='一键跑全程', text='准备中…')

            def job(progress):
                from core import pipeline
                return pipeline.run_all(files=files, folders=folders, out_dir=out_dir,
                                        progress=progress, drop_self=var_self.get(),
                                        dedupe=var_dedupe.get())

            def done(res):
                self.state['last_out_dir'] = res['out_dir']
                pv = res['results'].get('pivot')
                self.state['last_input'] = pv['xlsx'] if pv else res['final']
                self.status.config(text='最近输出：{}'.format(res['out_dir']))
                dialogs.ResultDialog(
                    win, title='一键跑全程完成', summary=res['summary'], items=res['products'],
                    out_dir=res['out_dir'], log_file=res['log'],
                    next_label='继续：资金流向分析', on_next=lambda: self._open('flow'))

            def fail(err, trace):
                logger.log('一键跑全程失败：{}'.format(err))
                dialogs.show_error(win, '执行失败', err, trace)

            runner.start(job, on_done=done, on_error=fail, progress=pw)

        bottom = ttk.Frame(body)
        bottom.pack(fill='x')
        ttk.Button(bottom, text='开始', command=start).pack(side='left')
        ttk.Button(bottom, text='关闭', command=win.destroy).pack(side='right')
        dialogs.center(win, self.root)
        return win


def main():
    root = tk.Tk()
    SuiteApp(root)
    root.update_idletasks()
    dialogs.center(root, None)
    root.mainloop()


if __name__ == '__main__':
    main()
