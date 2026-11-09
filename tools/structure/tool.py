# -*- coding: utf-8 -*-
"""试验性功能：无人物关系图时，估算资金层级、团体与净流向并生成 Markdown。"""
import os
import tkinter as tk
from tkinter import ttk

from core import log as logger
from core import paths
from core.progress import NullProgress
from tools.structure.analyze import build_model
from tools.structure.input import read_input
from tools.structure.report import render_markdown

REPORT_NAME = '05_资金流水走向大致分析与人物关系预估.md'


def run_report(input_file, out_dir=None, progress=None, min_amount=0.0,
               max_subjects=60, session=True):
    """headless 入口：读取流水 → 估算结构 → 写 Markdown 报告。"""
    progress = progress or NullProgress(logger.log)
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(input_file))
    os.makedirs(out_dir, exist_ok=True)
    logger.log('试验性资金结构分析：{} → {}'.format(input_file, out_dir))
    progress.set_determinate(3, '正在读取流水…')
    df, input_info = read_input(input_file)
    progress.update(1, 3, '正在估算主体层级与资金团体…')
    model = build_model(df, min_amount=min_amount, max_subjects=max_subjects)
    progress.update(2, 3, '正在生成 Markdown 报告…')
    text = render_markdown(model, source_file=input_file, input_info=input_info)
    report = paths.product(out_dir, REPORT_NAME)
    paths.backup_if_exists(report)
    with open(report, 'w', encoding='utf-8', newline='\n') as f:
        f.write(text)
    stats = model['stats']
    logger.log('已写出：{}（{} 个核心主体 / {} 个团体 / {} 条净流向通道）'.format(
        report, stats['subjects_core'], stats['groups'], stats['net_channels']))
    if session:
        paths.update_session(out_dir, 'structure', 'done', outputs={'report': report},
                             args={'min_amount': float(min_amount or 0.0),
                                   'max_subjects': int(max_subjects),
                                   'input_kind': input_info.get('kind')},
                             note='试验性估算：{} 行有效流水 / {} 个主体 / 核心 {} 人'.format(
                                 stats['rows_valid'], stats['subjects_scope'], stats['subjects_core']))
    progress.update(3, 3, '完成')
    return {'out_dir': out_dir, 'report': report, 'stats': stats, 'model': model,
            'input_info': input_info, 'log': logger.log_path()}


def open_window(parent, state=None):
    """功能窗口：选择流水总表/统一流水 → 生成试验性 Markdown 分析。"""
    state = state if state is not None else {}
    from ui import dialogs
    from ui.progress import ProgressWindow
    from ui.worker import TaskRunner

    win = tk.Toplevel(parent)
    win.title('（试验性）资金流水走向大致分析和生成人物关系图')
    win.resizable(False, False)
    try:
        win.transient(parent)
    except Exception:
        pass
    body = ttk.Frame(win, padding=14)
    body.pack(fill='both', expand=True)

    ttk.Label(body, text='输入流水总表或统一流水', font=('', 10, 'bold')).pack(anchor='w')
    ttk.Label(body, text='支持 03_流水总表.xlsx / 03_数据透视表.xlsx / 统一流水 tsv、csv，'
                         '也兼容微信原始总表（交易金额(分) 会自动换算为元）。',
              wraplength=590, foreground='#555555', justify='left').pack(anchor='w', pady=(2, 4))
    var_in = tk.StringVar(value=state.get('last_input') if os.path.isfile(str(state.get('last_input') or '')) else '')
    row = ttk.Frame(body)
    row.pack(fill='x', pady=(4, 10))
    ttk.Entry(row, textvariable=var_in, width=66).pack(side='left', fill='x', expand=True)
    ttk.Button(row, text='浏览…', command=lambda: var_in.set(
        dialogs.ask_open_file(win, '选择流水总表或统一流水', dialogs.FT_FLOW) or var_in.get())
    ).pack(side='left', padx=(6, 0))

    opt = ttk.LabelFrame(body, text='核心图范围', padding=8)
    opt.pack(fill='x', pady=(0, 10))
    var_min = tk.StringVar(value='0')
    var_max = tk.StringVar(value='60')
    ttk.Label(opt, text='最小通道金额（元，0 = 不过滤）').grid(row=0, column=0, sticky='w')
    ttk.Entry(opt, textvariable=var_min, width=14).grid(row=0, column=1, sticky='w', padx=(8, 24))
    ttk.Label(opt, text='核心主体上限').grid(row=0, column=2, sticky='w')
    ttk.Spinbox(opt, from_=10, to=500, textvariable=var_max, width=10).grid(
        row=0, column=3, sticky='w', padx=(8, 0))
    ttk.Label(opt, text='P0 = 每个团体的枢纽候选；P1/P2/P3 = 与 P0 的图距离。'
                        '资金方向按双向对抵后的净额估计；低流量主体会被省略以避免图形拥挤。',
              wraplength=570, foreground='#8a6d3b', justify='left').grid(
        row=1, column=0, columnspan=4, sticky='w', pady=(8, 0))

    ttk.Label(body, text='输出目录（留空则自动新建 输出_时间戳）').pack(anchor='w')
    var_out = tk.StringVar(value=state.get('last_out_dir') or '')
    row2 = ttk.Frame(body)
    row2.pack(fill='x', pady=(4, 10))
    ttk.Entry(row2, textvariable=var_out, width=66).pack(side='left', fill='x', expand=True)
    ttk.Button(row2, text='浏览…', command=lambda: var_out.set(
        dialogs.ask_directory(win, '选择输出目录', var_out.get() or None) or var_out.get())
    ).pack(side='left', padx=(6, 0))
    ttk.Label(body, text='产物：' + REPORT_NAME + '\n'
                         '报告含 P0–P3 人员分层、资金团体、人员架构树、Mermaid 资金走向图与候选链路。\n'
                         '试验性估算仅供摸排线索，不能替代人物关系核实或司法会计鉴定。',
              wraplength=590, foreground='#555555', justify='left').pack(anchor='w')

    runner = TaskRunner(win)

    def start():
        src = var_in.get().strip().strip('"')
        if not src or not os.path.isfile(src):
            dialogs.show_error(win, '输入无效', '请选择一个存在的流水总表或统一流水文件。')
            return
        try:
            min_amount = float(var_min.get().strip() or 0)
            max_subjects = int(var_max.get().strip() or 60)
            if min_amount < 0 or max_subjects < 2:
                raise ValueError
        except ValueError:
            dialogs.show_error(win, '参数无效', '最小通道金额需为非负数字，核心主体上限至少为 2。')
            return
        if runner.busy():
            return
        out_dir = var_out.get().strip().strip('"')
        if not out_dir:
            out_dir = paths.new_output_dir(paths.default_output_base(src))
        logger.init_log(out_dir)
        state['last_input'] = src
        state['last_out_dir'] = out_dir
        pw = ProgressWindow(win, title='正在估算资金结构', text='准备中…')

        def job(progress):
            return run_report(src, out_dir=out_dir, progress=progress,
                              min_amount=min_amount, max_subjects=max_subjects)

        def done(res):
            s = res['stats']
            summary = ('有效流水 {} 行，{} 个主体、{} 条通道。\n'
                       '报告展示 {} 个核心主体，估算为 {} 个资金团体、{} 条净流向通道。\n'
                       'P0–P3 仅为图结构估算，请结合原始流水和人物关系人工复核。').format(
                           s['rows_valid'], s['subjects_scope'], s['channels_scope'],
                           s['subjects_core'], s['groups'], s['net_channels'])
            if s['subjects_omitted']:
                summary += '\n另有 {} 个低流量主体未进入核心图。'.format(s['subjects_omitted'])
            dialogs.ResultDialog(win, title='试验性资金结构分析完成', summary=summary,
                                 items=[('Markdown 分析报告', res['report'])],
                                 out_dir=res['out_dir'], log_file=res['log'])

        def fail(err, trace):
            logger.log('失败：{}'.format(err))
            dialogs.show_error(win, '分析失败', err, trace)

        runner.start(job, on_done=done, on_error=fail, progress=pw)

    bottom = ttk.Frame(body)
    bottom.pack(fill='x', pady=(10, 0))
    ttk.Button(bottom, text='开始分析', command=start).pack(side='left')
    ttk.Button(bottom, text='关闭', command=win.destroy).pack(side='right')
    dialogs.center(win, parent)
    return win


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        result = run_report(sys.argv[1], out_dir=sys.argv[2] if len(sys.argv) > 2 else None)
        print('OK', result['report'])
