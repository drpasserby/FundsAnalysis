# -*- coding: utf-8 -*-
"""
资金流水工具套件 v2.0.5
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
    """主窗口：五个功能卡片。"""

    def __init__(self, root):
        self.root = root
        self.state = {'last_out_dir': '', 'last_input': '', 'last_style': 'classic',
                      'last_policy': 'spend'}
        root.title('{} v{}'.format(APP_NAME, __version__))
        root.configure(bg=BG)
        self._header()
        self._cards()
        self._bottom()

    # ---------- 界面 ----------
    def _header(self):
        head = tk.Frame(self.root, bg=BG)
        head.pack(fill='x', padx=18, pady=(16, 4))
        tk.Label(head, text=APP_NAME, font=(FONT_UI, 17, 'bold'), bg=BG).pack(anchor='w')
        tk.Label(head, text='版本 v{}    ·    {}    ·    GPLv3    ·    {}'.format(
            __version__, AUTHOR, GITHUB), bg=BG, fg=FG_DIM, font=('', 9)).pack(anchor='w', pady=(3, 0))

    def _cards(self):
        wrap = tk.Frame(self.root, bg=BG)
        wrap.pack(fill='both', expand=True, padx=18, pady=8)
        for spec in registry.TOOLS:
            self._card(wrap, spec.title, spec.subtitle,
                       lambda s=spec: self._open(s.id), enabled=spec.enabled, note=spec.note)

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
        ttk.Button(bar, text='打开最近输出目录', command=self._open_recent).pack(side='left')
        ttk.Button(bar, text='关于 / 使用说明', command=self._about).pack(side='left', padx=6)
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

    def _about(self):
        msg = ('五个功能（可单独使用，按 1 → 2 → 3 → 4 → 5 的顺序最省事）：\n'
               '  1 微信流水导入：选择微信账单 TenpayTrades.txt（或其所在文件夹），'
               '生成 01_统一流水.tsv 与 01_清洗报告.txt。\n'
               '  2 支付宝流水导入：选择支付宝「账户明细」导出的 csv（或其所在文件夹），'
               '生成 01_统一流水_支付宝.tsv 与 01_清洗报告_支付宝.txt，列格式与微信一致、'
               '来源列标注「支付宝」。\n'
               '  3 多来源合并：选择一份或多份统一格式流水，合并去重后生成 02_合并流水.tsv。\n'
               '  4 流水总表与数据透视表生成：生成 03_流水总表.xlsx / .csv 与 03_数据透视表.xlsx'
               '（含数据透视表、按主体汇总、人物关系图三张表）。\n'
               '  5 资金流向分析图和报告生成：生成 04_资金流向图与流水分析报告.html'
               '（先看资金流向图，再由按钮进入分析报告）。\n\n'
               '推荐用法（看「某某和谁有往来」）：\n'
               '  ① 先在 03_数据透视表.xlsx 的「人物关系图」表里填姓名与角色；\n'
               '  ② 打开「资金流向分析图和报告生成」，输入直接选这个 03_数据透视表.xlsx'
               '（透视表本身就是「谁 ↔ 谁、多少钱」的成对数据，不需要逐笔时间）；\n'
               '  ③ 点「载入「人物关系图」人物」，默认只保留双方都在名单里的往来，'
               '还可再设金额下限收窄范围。\n\n'
               '输出：默认写入输入文件所在目录下的「输出_时间戳」文件夹；产物按固定名命名、'
               '集中在同一目录，下一步会自动读取，也可以手动选择。\n\n'
               '数据格式契约：前 4 列固定为「用户方 | 支出/收入 | 客户方 | 金额（元）」，'
               '其余列（交易时间、交易类型、订单号等）为可选扩展列。\n\n'
               '免责声明：本工具的分析结果仅供参考，不构成审计结论、鉴定意见或法律意见。\n\n'
               '版本：v{}    作者：{}    {}'.format(__version__, AUTHOR, GITHUB))
        dialogs.show_info(self.root, msg, '关于与使用说明')


def main():
    root = tk.Tk()
    SuiteApp(root)
    root.update_idletasks()
    dialogs.center(root, None)
    root.mainloop()


if __name__ == '__main__':
    main()
