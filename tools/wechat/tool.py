# -*- coding: utf-8 -*-
"""功能 A：导入与清洗 —— 微信原始账单 → 统一格式流水。

产物：`01_统一流水.tsv`（统一格式中间文件）+ `01_清洗报告.txt`。
"""
import io
import os
import tkinter as tk
from tkinter import ttk

from core import log as logger
from core import paths
from core.progress import NullProgress
from core.schema import ALL_COLUMNS, CORE_COLUMNS, EXT_COLUMNS, COL_AMOUNT, COL_CP, COL_DIR, COL_USER, write_tsv
from tools.wechat import parse


def collect_files(files=None, folders=None, progress=None):
    """汇总待解析的账单文件（文件 + 文件夹递归扫描）。"""
    out = []
    for p in (files or []):
        if p and os.path.isfile(p):
            out.append(p)
    for d in (folders or []):
        if d and os.path.isdir(d):
            out.extend(parse.find_bill_files(d))
    seen = set()
    uniq = []
    for p in sorted(out):
        k = os.path.normcase(os.path.abspath(p))
        if k not in seen:
            seen.add(k)
            uniq.append(p)
    return uniq


def _is_blank_column(s):
    """整列是否为空（空串 / NaN 都算空）。

    注意：pandas 3.x 的 str 类型里 NaN 仍是缺失值，`astype(str)` 不再得到字面量 'nan'，
    所以这里必须同时判断 `isna()`，否则整列 NaN 会被当成有值。
    """
    if len(s) == 0:
        return True
    v = s.astype(str).str.strip()
    return bool((s.isna() | (v == '') | (v == 'nan')).all())


def _empty_message(targets, stats, report):
    """清洗后无有效数据时，逐条列出规则的过滤行数，便于判断是账单格式问题还是规则问题。"""
    lines = ['没有可用的有效数据：{} 个文件共 {} 行原始记录，全部被清洗规则过滤。'.format(
        len(targets), stats['rows_raw'])]
    lines.append('  · 方向不是支出/收入：{} 行（R2）'.format(stats['dropped_direction']))
    if stats.get('direction_values'):
        lines.append('    实际取到的「借贷类型」取值：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['direction_values'].items()))
    lines.append('  · 金额缺失或非正：{} 行（R7）'.format(stats['dropped_amount']))
    lines.append('  · 用户方或客户方为空：{} 行（R7）'.format(stats['dropped_empty']))
    lines.append('  · 自转自：{} 行（R1）'.format(stats['dropped_self']))
    if stats.get('dropped_nontransfer'):
        lines.append('  · 非转账类：{} 行（R4，档位「{}」）'.format(
            stats['dropped_nontransfer'], parse.R4_LABELS.get(stats.get('r4_mode'), '')))
    if stats.get('usage_values'):
        lines.append('    实际取到的「交易用途类型 / 交易类型」取值：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['usage_values'].items()))
    if stats['missing_columns']:
        lines.append('  · 缺少必要列：{}'.format('、'.join(stats['missing_columns'])))
    lines.append('诊断报告已写出：{}'.format(report))
    lines.append('请确认账单为微信/财付通导出的 txt（表头含「用户侧账号名称/借贷类型/'
                 '对手侧账户名称/交易金额(分)」）。')
    return '\n'.join(lines)


def run_import(files=None, folders=None, out_dir=None, progress=None,
               drop_self=True, dedupe=False, columns=None, session=True, r4=None):
    """headless 可调用：解析 → 清洗 → 写统一流水 tsv 与清洗报告。返回结果字典。

    columns：要写入产物的列名列表（None = 全部列）；前 4 个必需列始终保留，
    顺序始终按统一格式排列（账单里的其余原始列本来就不会进入统一格式）。
    r4：非转账类处理档位（parse.R4_OFF / R4_TRANSFER / R4_P2P，None = 不过滤）。
    """
    progress = progress or NullProgress(logger.log)
    progress.set_indeterminate('正在查找账单文件…')
    targets = collect_files(files, folders)
    if not targets:
        raise ValueError('没有找到可解析的微信账单文件。\n请选择包含 TenpayTrades.txt 的文件夹，'
                         '或直接选择账单 txt 文件。')
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(targets[0]))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    logger.log('功能 A 导入与清洗：{} 个文件 → {}'.format(len(targets), out_dir))

    progress.set_determinate(max(1, len(targets)), '正在解析账单…')
    df, stats = parse.build_frame(targets, progress=progress, drop_self=drop_self, dedupe=dedupe,
                                  r4=r4 or parse.R4_OFF)
    if not df.empty:
        stats['selected_columns'] = list(columns) if columns else []
        if columns:
            picked = set(columns) | set(CORE_COLUMNS)
            keep = [c for c in ALL_COLUMNS if c in picked]
            dropped = [c for c in ALL_COLUMNS if c not in picked]
            if dropped:
                logger.log('按列选择丢弃：{}'.format('、'.join(dropped)))
            df = df[keep]
        # 整批为空的扩展列（例如这份账单没有「交易状态」）自动省略，避免产物里全是空列
        empty_cols = [c for c in df.columns if c in EXT_COLUMNS and _is_blank_column(df[c])]
        if empty_cols:
            df = df.drop(columns=empty_cols)
            stats['empty_columns'] = empty_cols
            logger.log('整批为空、已省略的扩展列：{}'.format('、'.join(empty_cols)))
        stats['columns_kept'] = list(df.columns)
    tsv = paths.product(out_dir, paths.P_IMPORT_TSV)
    report = paths.product(out_dir, paths.P_IMPORT_REPORT)
    if df.empty:
        with io.open(report, 'w', encoding='utf-8') as f:
            f.write(parse.format_report(stats, targets))
        logger.log('导入失败：清洗后无有效数据，诊断报告已写出 {}'.format(report))
        raise ValueError(_empty_message(targets, stats, report))

    progress.set_indeterminate('正在写出统一流水…')
    paths.backup_if_exists(tsv)
    write_tsv(df, tsv)
    with io.open(report, 'w', encoding='utf-8') as f:
        f.write(parse.format_report(stats, targets, tsv))
    logger.log('已写出：{}（{} 行）；{}'.format(tsv, len(df), report))

    if session:
        paths.update_session(out_dir, 'import', 'done',
                             outputs={'tsv': tsv, 'report': report},
                             args={'files': len(targets), 'drop_self': drop_self, 'dedupe': dedupe,
                                   'r4': parse.R4_LABELS.get(r4 or parse.R4_OFF, r4 or parse.R4_OFF),
                                   'columns': stats.get('columns_kept') or '全部'},
                             note='{} 个文件 / {} 行有效记录 / {} 个主体'.format(
                                 stats['files_ok'], len(df), stats['subjects']))
    progress.update(1, 1, '完成')
    return {'out_dir': out_dir, 'tsv': tsv, 'report': report, 'stats': stats,
            'rows': len(df), 'files': targets, 'log': logger.log_path()}


def open_window(parent, state=None):
    """功能窗口：选择微信账单（文件夹/文件）→ 生成统一流水。"""
    state = state if state is not None else {}
    from ui import dialogs
    from ui.progress import ProgressWindow
    from ui.worker import TaskRunner

    win = tk.Toplevel(parent)
    win.title('导入与清洗（微信账单）· 资金流水工具套件')
    win.resizable(False, False)
    try:
        win.transient(parent)
    except Exception:
        pass
    body = ttk.Frame(win, padding=14)
    body.pack(fill='both', expand=True)

    ttk.Label(body, text='微信账单文件（TenpayTrades.txt，可多选或整个文件夹）',
              font=('', 10, 'bold')).pack(anchor='w')
    listbox = tk.Listbox(body, width=78, height=7, activestyle='none')
    listbox.pack(fill='both', expand=True, pady=(4, 6))
    selected = []

    def refresh():
        listbox.delete(0, 'end')
        for p in selected:
            listbox.insert('end', p)

    def add_files():
        ps = dialogs.ask_open_files(win, '选择微信账单 txt（可多选）', dialogs.FT_WECHAT)
        for p in ps:
            if p not in selected:
                selected.append(p)
        refresh()

    def add_folder():
        d = dialogs.ask_directory(win, '选择包含账单的文件夹（递归查找）', state.get('last_dir'))
        if d and d not in selected:
            selected.append(d)
            state['last_dir'] = d
            refresh()

    def clear():
        del selected[:]
        refresh()

    btn = ttk.Frame(body)
    btn.pack(fill='x')
    ttk.Button(btn, text='添加账单文件…', command=add_files).pack(side='left')
    ttk.Button(btn, text='添加文件夹…', command=add_folder).pack(side='left', padx=6)
    ttk.Button(btn, text='清空', command=clear).pack(side='left')

    opt = ttk.LabelFrame(body, text='清洗选项', padding=8)
    opt.pack(fill='x', pady=10)
    var_self = tk.BooleanVar(value=True)
    var_dedupe = tk.BooleanVar(value=False)
    ttk.Checkbutton(opt, text='去除自转自流水（用户方 = 客户方，规则 R1）',
                    variable=var_self).pack(anchor='w')
    ttk.Checkbutton(opt, text='去除重复记录（订单号优先，规则 R5）',
                    variable=var_dedupe).pack(anchor='w')

    var_r4 = tk.StringVar(value=parse.R4_OFF)
    ttk.Label(opt, text='非转账类处理（规则 R4，微信账单的「交易用途类型」有'
                        '转账/提现/消费/退款/充值/手续费/信用卡还款 七类）：').pack(anchor='w', pady=(6, 0))
    for mode in parse.R4_CHOICES:
        ttk.Radiobutton(opt, text=parse.R4_LABELS[mode], value=mode,
                        variable=var_r4).pack(anchor='w', padx=(16, 0))

    picked = {'cols': state.get('columns') or list(ALL_COLUMNS)}
    lbl_cols = ttk.Label(opt, text='')

    def refresh_cols():
        cols = picked['cols']
        extra = [c for c in cols if c in EXT_COLUMNS]
        if len(cols) >= len(ALL_COLUMNS):
            lbl_cols.configure(text='将写入：全部 {} 列（{}）'.format(
                len(ALL_COLUMNS), '、'.join(ALL_COLUMNS)))
        else:
            lbl_cols.configure(text='将写入：4 个必需列 + {} 个扩展列（{}）'.format(
                len(extra), '、'.join(extra) if extra else '无'))

    def choose_cols():
        sel = dialogs.ask_columns(
            win, ALL_COLUMNS, title='选择需要写入的列', locked=CORE_COLUMNS,
            selected=picked['cols'],
            note='账单原始有 27 列，套件只保留统一格式的前 4 个必需列 + 以下扩展列；'
                 '不勾选的扩展列不会写进 01_统一流水.tsv，后续总表 / 流向图也不会出现。')
        if sel:
            picked['cols'] = [c for c in ALL_COLUMNS if c in set(sel)]
            state['columns'] = picked['cols']
            refresh_cols()

    ttk.Button(opt, text='选择需要写入的列…', command=choose_cols).pack(anchor='w', pady=(4, 0))
    lbl_cols.pack(anchor='w')
    refresh_cols()

    ttk.Label(opt, text='说明：非「支出/收入」记录、空值与非正金额记录一律去除（规则 R2/R7）；'
                        '「交易金额(分)」自动换算为元并给出对账（规则 R6）；'
                        '整批为空的扩展列会自动省略（例如这份账单没有「交易状态」）；'
                        '文本字段会清掉不可见字符（零宽字符、双向控制符）。',
              wraplength=520, foreground='#555555', justify='left').pack(anchor='w', pady=(4, 0))
    ttk.Label(opt, text='支付宝账单入口预留：待提供真实样本后启用（本版不解析支付宝）。',
              wraplength=520, foreground='#8a6d3b', justify='left').pack(anchor='w', pady=(4, 0))

    ttk.Label(body, text='输出目录（留空则自动新建 输出_时间戳）').pack(anchor='w')
    var_out = tk.StringVar(value=state.get('last_out_dir') or '')
    row = ttk.Frame(body)
    row.pack(fill='x', pady=(4, 10))
    ttk.Entry(row, textvariable=var_out, width=58).pack(side='left', fill='x', expand=True)
    ttk.Button(row, text='浏览…', command=lambda: var_out.set(
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
        out_dir = var_out.get().strip().strip('"')
        if not out_dir:
            base = os.path.dirname(os.path.abspath(files[0])) if files else folders[0]
            out_dir = paths.new_output_dir(base)
        logger.init_log(out_dir)
        state['last_out_dir'] = out_dir
        state['last_input'] = files[0] if files else ''
        pw = ProgressWindow(win, title='正在导入与清洗', text='准备中…')

        def job(progress):
            return run_import(files=files, folders=folders, out_dir=out_dir,
                              progress=progress, drop_self=var_self.get(),
                              dedupe=var_dedupe.get(), columns=picked['cols'],
                              r4=var_r4.get())

        def done(res):
            s = res['stats']
            summary = ('已解析 {} 个文件，得到 {} 行有效记录，涉及 {} 个主体。\n'
                       '金额（分）合计 {:.2f} → {:.2f} 元，入表 {:.2f} 元。'.format(
                           s['files_ok'], res['rows'], s['subjects'],
                           s['fen_total'], s['fen_total'] / 100.0, s['yuan_total']))
            if s.get('r4_mode') and s['r4_mode'] != parse.R4_OFF:
                summary += '\n已按「{}」过滤掉 {} 行非转账类记录。'.format(
                    parse.R4_LABELS.get(s['r4_mode'], ''), s['dropped_nontransfer'])
            if s.get('biz_type_values'):
                summary += '\n写入列：{} 列（整批为空的列已省略{}）。'.format(
                    len(s.get('columns_kept') or []),
                    '：' + '、'.join(s['empty_columns']) if s.get('empty_columns') else '')
            if s['files_ok'] > 1 and not var_dedupe.get():
                summary += ('\n提示：本批成功解析 {} 个文件且未开启 R5 去重——若其中存在重叠导出'
                            '（同一账户多次导出同一批流水），金额会重复计入；'
                            '可勾选「去除重复记录」后重跑对比。'.format(s['files_ok']))
            if s.get('recv_diff_rows'):
                summary += ('\n提示：{} 行「交易金额」与账单里的「对手方接收金额」不一致'
                            '（多为微信红包），已按「交易金额」入表，明细见清洗报告第五节。'
                            .format(s['recv_diff_rows']))
            dialogs.ResultDialog(
                win, title='导入与清洗完成', summary=summary,
                items=[('统一流水', res['tsv']), ('清洗报告', res['report'])],
                out_dir=res['out_dir'], log_file=res['log'],
                next_label='下一步：多来源合并', on_next=lambda: _next_step(parent, state, 'merge'))

        def fail(err, trace):
            logger.log('失败：{}'.format(err))
            dialogs.show_error(win, '导入失败', err, trace)

        runner.start(job, on_done=done, on_error=fail, progress=pw)

    bar = ttk.Frame(body)
    bar.pack(fill='x')
    ttk.Button(bar, text='开始导入', command=start).pack(side='left')
    ttk.Button(bar, text='关闭', command=win.destroy).pack(side='right')
    dialogs.center(win, parent)
    return win


def _next_step(parent, state, tid):
    from core import registry
    try:
        mod = registry.load_tool(tid)
        mod.open_window(parent, state)
    except Exception as e:
        logger.log('打开下一步失败：{}'.format(e))


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        r = run_import(folders=[sys.argv[1]] if os.path.isdir(sys.argv[1]) else None,
                       files=[sys.argv[1]] if os.path.isfile(sys.argv[1]) else None,
                       out_dir=sys.argv[2] if len(sys.argv) > 2 else None)
        print('OK', r['tsv'], r['rows'], '行')
