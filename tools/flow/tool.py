# -*- coding: utf-8 -*-
"""功能 D：资金流向分析 —— UI 窗口与 headless 入口。

输入：统一格式流水（.xlsx/.xls/.tsv/.txt/.csv），兼容 v1.3.2 的 4 列 Excel。
输出：`04_资金流向图与流水分析报告.html`（图 + 报告合并页，自包含、离线可用）。
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk

from core import log as logger
from core import paths
from core.progress import NullProgress
from core.schema import read_table
from tools.flow.analyze import (POLICIES, POLICY_LABELS, build_contract, build_graphs,
                                validate_and_clean)
from tools.flow.render import generate_html

DEFAULT_TITLE = '资金流水分析演示图'


def run_flow(input_path, out_dir=None, title=DEFAULT_TITLE, style='classic', policy='spend',
             progress=None, source_label=None, session=True):
    """headless 可调用：读取 → 校验 → 构图 → 契约 → 生成 HTML。返回结果字典。"""
    progress = progress or NullProgress(logger.log)
    if policy not in POLICY_LABELS:
        policy = 'spend'
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(input_path))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    html_path = paths.product(out_dir, paths.P_FLOW_HTML)

    progress.set_indeterminate('正在读取流水数据…')
    df = read_table(input_path)
    logger.log('读取完成：{} 行'.format(len(df)))

    progress.set_indeterminate('正在校验与清洗…')
    clean_df, data_messages = validate_and_clean(df)
    for m in data_messages:
        logger.log('[校验] ' + m.replace('\n', ' '))
    if clean_df.empty:
        raise ValueError('没有可用的有效数据。')

    progress.set_determinate(3, '正在构图…')
    G = build_graphs(clean_df, policy=policy)
    progress.update(1, 3, '正在构建聚合契约…')
    contract = build_contract(G)
    audit = contract.get('audit') or []
    layout_note = contract.get('layout_note')
    if layout_note:
        logger.log('[布局] ' + layout_note)
    progress.update(2, 3, '正在渲染 HTML…')
    html = generate_html(contract, title=title, hide_other=True, default_mode=style,
                         policy_label=POLICY_LABELS[policy],
                         source_file=source_label or input_path)
    paths.backup_if_exists(html_path)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    progress.update(3, 3, '完成')

    nodes = len(contract.get('nodes') or [])
    edges = len(contract.get('edges') or [])
    if session:
        paths.update_session(out_dir, 'flow', 'done',
                             outputs={'html': html_path, 'input': input_path},
                             args={'style': style, 'policy': policy, 'title': title},
                             note='{} 行 → {} 主体 / {} 通道'.format(len(clean_df), nodes, edges))
    logger.log('已生成：{}（{} 主体 / {} 通道）'.format(html_path, nodes, edges))
    return {'out_dir': out_dir, 'html': html_path, 'audit': audit, 'messages': data_messages,
            'rows': len(clean_df), 'nodes': nodes, 'edges': edges, 'log': logger.log_path(),
            'layout_note': layout_note}


def guess_input(state):
    """接力：优先取同一输出目录里上游（C 总表 / B 合并 / A 统一流水）的产物。"""
    prev = (state or {}).get('last_out_dir')
    if prev:
        for name in (paths.P_PIVOT_XLSX, paths.P_MERGE_TSV, paths.P_IMPORT_TSV):
            p = paths.product(prev, name)
            if os.path.exists(p):
                return p
    return (state or {}).get('last_input') or ''


def audit_text(audit):
    """把一致性结果拼成与 1.3.2 相同口径的中文提示（返回 None 表示无差异）。"""
    if not audit:
        return None
    total_diff = sum(r['diff'] for r in audit)
    lines = ['数据一致性检查：发现 {} 处方向记录不一致，差异合计 {:.2f} 元。'.format(len(audit), total_diff)]
    for r in audit[:10]:
        lines.append('{} → {}：支出方 {} / 收入方 {}（差 {}，{:.1f}%）'.format(
            r['a'], r['b'], r['spend'], r['recv'], r['diff'], r['ratio']))
    if len(audit) > 10:
        lines.append('……共 {} 处，完整明细与参考分析见 {} 的报告页'.format(len(audit), paths.P_FLOW_HTML))
    return '\n'.join(lines)


def open_window(parent, state=None):
    """功能窗口：选择流水文件 + 风格/口径 → 生成 HTML。"""
    state = state if state is not None else {}
    win = tk.Toplevel(parent)
    win.title('资金流向分析 · 资金流水工具套件')
    win.resizable(False, False)
    try:
        win.transient(parent)
    except Exception:
        pass

    body = ttk.Frame(win, padding=14)
    body.pack(fill='both', expand=True)

    ttk.Label(body, text='输入流水文件（统一格式：用户方 | 支出/收入 | 客户方 | 金额）',
              font=('', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky='w')
    var_in = tk.StringVar(value=guess_input(state))
    ent_in = ttk.Entry(body, textvariable=var_in, width=58)
    ent_in.grid(row=1, column=0, columnspan=2, sticky='we', pady=(4, 2))

    ttk.Label(body, text='输出目录（留空则自动新建 输出_时间戳）').grid(row=2, column=0, columnspan=3,
                                                          sticky='w', pady=(8, 0))
    var_out = tk.StringVar(value=state.get('last_out_dir') or '')
    ttk.Entry(body, textvariable=var_out, width=58).grid(row=3, column=0, columnspan=2, sticky='we', pady=(4, 2))

    def browse_in():
        from ui import dialogs
        d = os.path.dirname(var_in.get()) if var_in.get() else None
        p = dialogs.ask_open_file(win, '选择流水文件', dialogs.FT_FLOW, d)
        if p:
            var_in.set(p)
            if not var_out.get():
                var_out.set(os.path.dirname(os.path.abspath(p)))

    def browse_out():
        from ui import dialogs
        d = dialogs.ask_directory(win, '选择输出目录', var_out.get() or None)
        if d:
            var_out.set(d)

    ttk.Button(body, text='浏览…', command=browse_in).grid(row=1, column=2, sticky='w', padx=(6, 0))
    ttk.Button(body, text='浏览…', command=browse_out).grid(row=3, column=2, sticky='w', padx=(6, 0))

    ttk.Separator(body, orient='horizontal').grid(row=4, column=0, columnspan=3, sticky='we', pady=10)

    ttk.Label(body, text='界面风格').grid(row=5, column=0, sticky='w')
    var_style = tk.StringVar(value=state.get('last_style') or 'classic')
    sf = ttk.Frame(body)
    sf.grid(row=6, column=0, columnspan=3, sticky='w', pady=(2, 8))
    ttk.Radiobutton(sf, text='经典', value='classic', variable=var_style).pack(side='left')
    ttk.Radiobutton(sf, text='iOS', value='ios', variable=var_style).pack(side='left', padx=(12, 0))

    ttk.Label(body, text='聚合口径（支出与收入记录不一致时以谁为准）').grid(row=7, column=0, columnspan=3, sticky='w')
    var_policy = tk.StringVar(value=state.get('last_policy') or 'spend')
    pf = ttk.Frame(body)
    pf.grid(row=8, column=0, columnspan=3, sticky='w', pady=(2, 10))
    for key, label in POLICIES:
        ttk.Radiobutton(pf, text=label, value=key, variable=var_policy).pack(side='left', padx=(0, 10))

    var_hint = tk.StringVar(value='说明：分析结论由规则推导，仅供参考，不构成审计结论或法律意见。')
    ttk.Label(body, textvariable=var_hint, wraplength=520, foreground='#555555',
              justify='left').grid(row=9, column=0, columnspan=3, sticky='w', pady=(0, 10))

    bar = ttk.Frame(body)
    bar.grid(row=10, column=0, columnspan=3, sticky='we')

    from ui import dialogs
    from ui.progress import ProgressWindow
    from ui.worker import TaskRunner
    runner = TaskRunner(win)

    def start():
        path = var_in.get().strip().strip('"')
        if not path or not os.path.exists(path):
            dialogs.show_error(win, '输入无效', '请先选择一个存在的数据文件。')
            return
        if runner.busy():
            return
        out_dir = var_out.get().strip().strip('"')
        if out_dir and not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as e:
                dialogs.show_error(win, '无法创建输出目录', e)
                return
        if not out_dir:
            out_dir = paths.new_output_dir(paths.default_output_base(path))
        logger.init_log(out_dir)
        logger.log('功能 D 资金流向分析：输入 {}'.format(path))
        state['last_input'] = path
        state['last_out_dir'] = out_dir
        state['last_style'] = var_style.get()
        state['last_policy'] = var_policy.get()
        pw = ProgressWindow(win, title='正在生成资金流向图', text='准备中…')

        def job(progress):
            return run_flow(path, out_dir=out_dir, style=var_style.get(), policy=var_policy.get(),
                            progress=progress)

        def done(res):
            audit = res.get('audit') or []
            msg = audit_text(audit)
            if msg:
                messagebox.showwarning('数据一致性提示', msg, parent=win)
            summary = '已生成：{} 主体 / {} 通道，共 {} 行有效记录。\n产物：{}'.format(
                res['nodes'], res['edges'], res['rows'], os.path.basename(res['html']))
            if res.get('layout_note'):
                summary += '\n说明：' + res['layout_note'] + '。'
            if res['nodes'] >= 500:
                summary += '\n提示：主体数较多，浏览器打开大图可能略慢，可用鼠标拖拽或滚轮缩放查看。'
            dialogs.ResultDialog(
                win, title='资金流向分析完成', summary=summary,
                items=[('资金流向图与报告', res['html']), ('输入数据', path)],
                out_dir=res['out_dir'], log_file=res['log'])

        def fail(err, trace):
            logger.log('失败：{}'.format(err))
            dialogs.show_error(win, '生成失败', err, trace)

        runner.start(job, on_done=done, on_error=fail, progress=pw)

    ttk.Button(bar, text='开始生成', command=start).pack(side='left')
    ttk.Button(bar, text='关闭', command=win.destroy).pack(side='right')

    win.bind('<Return>', lambda e: start())
    dialogs.center(win, parent)
    try:
        win.focus_set()
    except Exception:
        pass
    return win


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        r = run_flow(sys.argv[1], out_dir=sys.argv[2] if len(sys.argv) > 2 else None)
        print('OK', r['html'])
