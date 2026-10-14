# -*- coding: utf-8 -*-
"""功能 D：资金流向分析 —— UI 窗口与 headless 入口。

输入：功能 C 的汇总件 `03_数据透视表.xlsx`（用户方 ↔ 客户方 + 方向 + 求和金额，推荐），
或统一格式逐笔明细（.xlsx/.xls/.tsv/.txt/.csv），兼容 v1.3.2 的 4 列 Excel。
输出：`04_资金流向图与流水分析报告.html`（图 + 报告合并页，自包含、离线可用）。
"""
import os
import tkinter as tk
from tkinter import messagebox, ttk

from core import log as logger
from core import paths
from core.progress import NullProgress
from core.schema import read_table
from tools.flow import select
from tools.flow.analyze import (POLICIES, POLICY_LABELS, build_contract, build_graphs,
                                validate_and_clean)
from tools.flow.render import generate_html

DEFAULT_TITLE = '资金流水分析演示图'


def run_flow(input_path, out_dir=None, title=DEFAULT_TITLE, style='classic', policy='spend',
             progress=None, source_label=None, session=True, filters=None):
    """headless 可调用：识别输入 → 读取 → 筛选（可选）→ 校验 → 构图 → 契约 → 生成 HTML。

    input_path：功能 C 的汇总件 `03_数据透视表.xlsx`（用「数据透视表」sheet：用户方 ↔ 客户方
                的合计金额与笔数，研判只看两人之间的总流水，不需要时间），或逐笔明细
                （03_流水总表.xlsx / 02_合并流水.tsv / 01_统一流水.tsv …）。
    filters   ：{'names': [...], 'mode': 'strict'|'expand', 'min_amount': 数值,
                 'time_from': 'YYYY-MM-DD', 'time_to': 'YYYY-MM-DD',
                 'use_relation': True 表示 names 为空时自动取「人物关系图」全部姓名}。键都可省略。
                默认 mode='strict'：只保留**双方都在名单里**的往来。
    """
    progress = progress or NullProgress(logger.log)
    if policy not in POLICY_LABELS:
        policy = 'spend'
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(input_path))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    html_path = paths.product(out_dir, paths.P_FLOW_HTML)

    progress.set_indeterminate('正在识别输入文件…')
    kind, input_file, src_note = select.resolve_input(input_path)
    logger.log('[输入] ' + src_note)

    progress.set_indeterminate('正在读取流水数据…')
    if kind == select.KIND_PIVOT:
        # C 的汇总透视表：用户方 ↔ 客户方 + 方向 + 求和金额，直接就是流向图要的边数据
        df = select.load_pivot_detail(input_file)
        logger.log('按「数据透视表」sheet 读取：{} 行（汇总口径，无逐笔时间）'.format(len(df)))
    else:
        df = read_table(input_file)
        logger.log('读取完成：{} 行'.format(len(df)))

    finfo = None
    if filters:
        names = list(filters.get('names') or [])
        if not names and filters.get('use_relation'):
            rel = select.load_relation(input_file if kind == select.KIND_PIVOT else '')
            names = select.relation_names(rel)
            if names:
                logger.log('[名单] 取自「人物关系图」：{} 人'.format(len(names)))
            else:
                logger.log('[名单] 「人物关系图」里还没有姓名，本次不按人物筛选')
        progress.set_indeterminate('正在按关注对象筛选…')
        df, finfo = select.apply_filters(
            df, names=names, min_amount=filters.get('min_amount'),
            time_from=filters.get('time_from'), time_to=filters.get('time_to'),
            mode=filters.get('mode') or select.MODE_STRICT)
        if select.is_active(finfo):
            logger.log('[筛选] ' + select.summary(finfo))
            if finfo.get('no_time_column'):
                logger.log('[筛选] 输入没有交易时间列（汇总透视表），日期范围未生效')
            if df.empty:
                raise ValueError('筛选后没有剩下任何记录（{}）。\n'
                                 '请放宽关注人物、金额下限或日期范围。'.format(select.summary(finfo)))
        else:
            finfo = None

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
    src_label = source_label or input_file
    if finfo:
        src_label = '{}（已筛选：{}）'.format(src_label, select.summary(finfo))
    html = generate_html(contract, title=title, hide_other=True, default_mode=style,
                         policy_label=POLICY_LABELS[policy],
                         source_file=src_label)
    paths.backup_if_exists(html_path)
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)
    progress.update(3, 3, '完成')

    nodes = len(contract.get('nodes') or [])
    edges = len(contract.get('edges') or [])
    if session:
        note = '{} 行 → {} 主体 / {} 通道'.format(len(clean_df), nodes, edges)
        if finfo:
            note = '{}；{}'.format(note, select.summary(finfo))
        paths.update_session(out_dir, 'flow', 'done',
                             outputs={'html': html_path, 'input': input_file},
                             args={'style': style, 'policy': policy, 'title': title,
                                   'filters': finfo},
                             note=note)
    logger.log('已生成：{}（{} 主体 / {} 通道）'.format(html_path, nodes, edges))
    return {'out_dir': out_dir, 'html': html_path, 'audit': audit, 'messages': data_messages,
            'rows': len(clean_df), 'nodes': nodes, 'edges': edges, 'log': logger.log_path(),
            'layout_note': layout_note, 'input': input_file, 'filter': finfo}


def guess_input(state):
    """接力：优先取同一输出目录里功能 C 的产物。

    C 的 `03_数据透视表.xlsx` 本身就是合法输入（汇总口径直接可用，且「人物关系图」sheet
    就是筛选名单的来源），所以这里优先推荐它，其次才是 B 的合并流水与 A 的统一流水。
    """
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
    """功能窗口：选择流水文件 + 选关注的人 + 风格/口径 → 生成 HTML。"""
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

    ttk.Label(body, text='输入流水文件（逐笔明细：用户方 | 支出/收入 | 客户方 | 金额）',
              font=('', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky='w')
    var_in = tk.StringVar(value=guess_input(state))
    ent_in = ttk.Entry(body, textvariable=var_in, width=58)
    ent_in.grid(row=1, column=0, columnspan=2, sticky='we', pady=(4, 2))

    var_src = tk.StringVar(value='')
    lbl_src = ttk.Label(body, textvariable=var_src, wraplength=520, justify='left',
                        foreground='#555555')
    lbl_src.grid(row=2, column=0, columnspan=3, sticky='w')

    ttk.Label(body, text='输出目录（留空则自动新建 输出_时间戳）').grid(row=3, column=0, columnspan=3,
                                                          sticky='w', pady=(8, 0))
    var_out = tk.StringVar(value=state.get('last_out_dir') or '')
    ttk.Entry(body, textvariable=var_out, width=58).grid(row=4, column=0, columnspan=2,
                                                         sticky='we', pady=(4, 2))

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
    ttk.Button(body, text='浏览…', command=browse_out).grid(row=4, column=2, sticky='w', padx=(6, 0))

    def refresh_src(*_a):
        p = var_in.get().strip().strip('"')
        info = select.detect_source(p) if p else {'kind': select.KIND_UNKNOWN,
                                                  'note': '尚未选择文件'}
        txt = info.get('note') or ''
        if info.get('kind') == select.KIND_PIVOT:
            var_src.set('输入：' + txt)
            lbl_src.configure(foreground='#1a6b1a')
        else:
            var_src.set('输入：' + txt)
            lbl_src.configure(foreground='#555555')

    var_in.trace_add('write', refresh_src)
    refresh_src()

    ttk.Separator(body, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky='we', pady=10)

    flt = {'names': [], 'mode': select.MODE_STRICT, 'min_amount': 0.0,
           'time_from': '', 'time_to': '', 'use_relation': False}
    fbox = ttk.LabelFrame(body, text='筛选：只看关注的人之间的流水')
    fbox.grid(row=6, column=0, columnspan=3, sticky='we')

    var_flt = tk.StringVar(value='未启用（分析全部流水）')
    ttk.Label(fbox, textvariable=var_flt, wraplength=520, justify='left',
              foreground='#555555').pack(anchor='w', padx=8, pady=(6, 2))

    fbtns = ttk.Frame(fbox)
    fbtns.pack(fill='x', padx=8, pady=(0, 8))

    def refresh_flt():
        if flt['names'] or flt['min_amount'] or flt['time_from'] or flt['time_to']:
            parts = []
            if flt['names']:
                parts.append('关注人物 {} 个（{}）'.format(
                    len(flt['names']), dict(select.MODE_LABELS).get(flt['mode'], '')))
                if flt.get('use_relation'):
                    parts.append('名单取自「人物关系图」')
            if flt['min_amount']:
                parts.append('金额 ≥ {:g} 元'.format(flt['min_amount']))
            if flt['time_from'] or flt['time_to']:
                parts.append('日期 {} ~ {}'.format(flt['time_from'] or '起', flt['time_to'] or '止'))
            var_flt.set('已启用：' + '；'.join(parts))
        else:
            var_flt.set('未启用（分析全部流水）')

    def pick_people():
        from ui import dialogs
        path = var_in.get().strip().strip('"')
        if not path or not os.path.exists(path):
            dialogs.show_error(win, '输入无效', '请先选择一个存在的数据文件。')
            return
        win.configure(cursor='watch')
        try:
            win.update_idletasks()
        except Exception:
            pass
        try:
            info = select.detect_source(path)
            is_pivot = info.get('kind') == select.KIND_PIVOT
            rel_path = path if is_pivot else paths.product(
                os.path.dirname(os.path.abspath(path)), paths.P_PIVOT_XLSX)
            relation = select.load_relation(rel_path)
            subjects = select.load_summary(path) if is_pivot else None
            if subjects is None or subjects.empty:
                subjects = select.subject_table(read_table(path))
        except Exception as e:
            win.configure(cursor='')
            dialogs.show_error(win, '读取主体列表失败', e)
            return
        win.configure(cursor='')
        if subjects is None or subjects.empty:
            dialogs.show_info(win, '这个文件里没有可用的主体，请确认选的是流水明细或 C 的汇总件。', '无法筛选')
            return
        got = dialogs.ask_people(
            win, subjects, relation=relation, selected=flt['names'], min_amount=flt['min_amount'],
            time_from=flt['time_from'], time_to=flt['time_to'], mode=flt['mode'],
            note='勾选关注的人：默认只保留**双方都在名单里**的往来（人物来源可以是「人物关系图」，'
                 '也可以直接在下表里挑）。当前列出 {} 个主体（按笔数降序）。'.format(len(subjects)))
        if got is None:
            return
        flt.update(got)
        flt['use_relation'] = False
        refresh_flt()

    def use_relation_all():
        """一键把 03_数据透视表.xlsx「人物关系图」里的全部姓名作为关注对象（双方都要在名单里）。"""
        from ui import dialogs
        path = var_in.get().strip().strip('"')
        if not path or not os.path.exists(path):
            dialogs.show_error(win, '输入无效', '请先选择一个存在的数据文件。')
            return
        relation = select.load_relation(path)
        if relation.empty:
            relation = select.load_relation(paths.product(
                os.path.dirname(os.path.abspath(path)), paths.P_PIVOT_XLSX))
        names = select.relation_names(relation)
        if not names:
            dialogs.show_info(
                win,
                '这个文件里的「人物关系图」还没有姓名。\n'
                '请先在 03_数据透视表.xlsx 的「人物关系图」sheet 里填「姓名 / 角色」（内容可以来自其它研判），'
                '保存后再回到这里点一次。',
                '人物关系图是空的')
            return
        flt['names'] = names
        flt['mode'] = select.MODE_STRICT
        flt['use_relation'] = True
        refresh_flt()
        logger.log('[名单] 已载入「人物关系图」{} 人（只看名单内双方的往来）'.format(len(names)))

    def clear_flt():
        flt.update({'names': [], 'mode': select.MODE_STRICT, 'min_amount': 0.0,
                    'time_from': '', 'time_to': '', 'use_relation': False})
        refresh_flt()

    ttk.Button(fbtns, text='载入「人物关系图」人物', command=use_relation_all).pack(side='left')
    ttk.Button(fbtns, text='选择关注的人…', command=pick_people).pack(side='left', padx=(8, 0))
    ttk.Button(fbtns, text='清空筛选', command=clear_flt).pack(side='left', padx=(8, 0))

    ttk.Label(body, text='界面风格').grid(row=7, column=0, sticky='w', pady=(10, 0))
    var_style = tk.StringVar(value=state.get('last_style') or 'classic')
    sf = ttk.Frame(body)
    sf.grid(row=8, column=0, columnspan=3, sticky='w', pady=(2, 8))
    ttk.Radiobutton(sf, text='经典', value='classic', variable=var_style).pack(side='left')
    ttk.Radiobutton(sf, text='iOS', value='ios', variable=var_style).pack(side='left', padx=(12, 0))

    ttk.Label(body, text='聚合口径（支出与收入记录不一致时以谁为准）').grid(row=9, column=0, columnspan=3, sticky='w')
    var_policy = tk.StringVar(value=state.get('last_policy') or 'spend')
    pf = ttk.Frame(body)
    pf.grid(row=10, column=0, columnspan=3, sticky='w', pady=(2, 10))
    for key, label in POLICIES:
        ttk.Radiobutton(pf, text=label, value=key, variable=var_policy).pack(side='left', padx=(0, 10))

    var_hint = tk.StringVar(value='说明：分析结论由规则推导，仅供参考，不构成审计结论或法律意见。')
    ttk.Label(body, textvariable=var_hint, wraplength=520, foreground='#555555',
              justify='left').grid(row=11, column=0, columnspan=3, sticky='w', pady=(0, 10))

    bar = ttk.Frame(body)
    bar.grid(row=12, column=0, columnspan=3, sticky='we')

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
        state['last_people'] = list(flt['names'])
        pw = ProgressWindow(win, title='正在生成资金流向图', text='准备中…')

        def job(progress):
            return run_flow(path, out_dir=out_dir, style=var_style.get(), policy=var_policy.get(),
                            progress=progress, filters=dict(flt))

        def done(res):
            audit = res.get('audit') or []
            msg = audit_text(audit)
            if msg:
                messagebox.showwarning('数据一致性提示', msg, parent=win)
            summary = '已生成：{} 主体 / {} 通道，共 {} 行有效记录。\n产物：{}'.format(
                res['nodes'], res['edges'], res['rows'], os.path.basename(res['html']))
            if res.get('filter'):
                summary += '\n筛选：' + select.summary(res['filter'])
            if res.get('layout_note'):
                summary += '\n说明：' + res['layout_note'] + '。'
            if res['nodes'] >= 500:
                summary += '\n提示：主体数较多，浏览器打开大图可能略慢，可用鼠标拖拽或滚轮缩放查看。'
            dialogs.ResultDialog(
                win, title='资金流向分析完成', summary=summary,
                items=[('资金流向图与报告', res['html']), ('输入数据', res.get('input') or path)],
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
