# -*- coding: utf-8 -*-
"""功能 B：多来源合并 —— 统一格式流水（微信/支付宝/历史产物）→ `02_合并流水.tsv`。

输入：任意数量的统一格式表（`01_统一流水.tsv`、`03_流水总表.xlsx`/`.csv`，或任何
      前 4 列同为 `用户方 | 支出/收入 | 客户方 | 金额` 的表）。
处理：逐个 read_table → 纵向拼接 → 去重（R5，可关）→ 按交易时间稳定排序。
输出：`02_合并流水.tsv`（统一格式；「来源」列空则自动填文件名，便于追溯）。
"""
import io
import os
import tkinter as tk
from tkinter import ttk

import pandas as pd

from core import log as logger
from core import paths
from core.clean import dedupe_frame, sort_by_time
from core.progress import NullProgress
from core.schema import (COL_AMOUNT, COL_CP, COL_DIR, COL_ORDER, COL_SOURCE, COL_USER,
                         CSV_EXTS, DIR_RECV, DIR_SPEND, EXCEL_EXTS, TSV_EXTS,
                         read_table, write_tsv)

TABLE_EXTS = set(EXCEL_EXTS) | set(TSV_EXTS) | set(CSV_EXTS)


def collect_tables(inputs=None, folders=None):
    """汇总待合并的表：显式文件 + 文件夹递归（按扩展名过滤，去重）。"""
    out = []
    for p in (inputs or []):
        if p and os.path.isfile(p):
            out.append(p)
    for d in (folders or []):
        if d and os.path.isdir(d):
            for root, _dirs, names in os.walk(d):
                for fn in sorted(names):
                    if os.path.splitext(fn)[1].lower() in TABLE_EXTS:
                        out.append(os.path.join(root, fn))
    seen, uniq = set(), []
    for p in sorted(out):
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    return uniq


CONFLICT_LIMIT = 10


def detect_conflicts(df, limit=CONFLICT_LIMIT):
    """检查同一订单号在多来源/两侧记录里是否对不上（微信 + 支付宝合并后的冲突提示）。

    一笔交易在两侧账单里各有一条记录（付款方记「支出」、收款方记「收入」）属正常；
    这里只挑出「同一订单号自相矛盾」的组，供人工复核：
      · 两侧金额不一致：同一订单号的支出合计与收入合计不同
      · 同向重复：同一订单号在同一方向出现多行
      · 主体不一致：同一订单号出现多组不同的（用户方, 客户方）
    """
    info = {'groups': 0, 'amount_mismatch': 0, 'same_dir_dup': 0, 'party_mismatch': 0,
            'samples': []}
    if df is None or df.empty or COL_ORDER not in df.columns:
        return info
    order = df[COL_ORDER].astype(str).str.strip()
    sub = df[order != '']
    if sub.empty:
        return info
    keys = order[order != '']
    for oid, grp in sub.groupby(keys):
        if len(grp) < 2:
            continue
        dirs = grp[COL_DIR].astype(str).str.strip()
        amounts = pd.to_numeric(grp[COL_AMOUNT], errors='coerce').fillna(0.0)
        spend = float(amounts[dirs == DIR_SPEND].sum())
        recv = float(amounts[dirs == DIR_RECV].sum())
        parties = sorted(set(zip(grp[COL_USER].astype(str), grp[COL_CP].astype(str))))
        # 一笔交易的两侧记录是 (A→B) 与 (B→A)，属正常；只有出现多组不同「主体对」才算冲突
        pairs = set(frozenset((u, c)) for u, c in parties)
        kinds = []
        if spend > 0 and recv > 0 and abs(spend - recv) > 0.01:
            info['amount_mismatch'] += 1
            kinds.append('两侧金额不一致：支出合计 {:.2f} / 收入合计 {:.2f}（差 {:.2f}）'.format(
                spend, recv, abs(spend - recv)))
        same_dir = len(grp) - int(dirs.nunique())
        if same_dir > 0:
            info['same_dir_dup'] += 1
            kinds.append('同一方向重复 {} 行'.format(same_dir))
        if len(pairs) > 1:
            info['party_mismatch'] += 1
            kinds.append('主体组合 {} 组'.format(len(pairs)))
        if not kinds:
            continue
        info['groups'] += 1
        if len(info['samples']) < limit:
            who = '；'.join('{} → {}'.format(u, c) for u, c in parties[:3])
            info['samples'].append('订单号 {}：{}（{}）'.format(oid, '；'.join(kinds), who))
    return info


def format_report(stats, tables, skipped, conflicts, out_tsv=None):
    """把合并口径、来源分布与订单号一致性检查整理成文本（02_合并说明.txt）。"""
    import time
    lines = []
    lines.append('资金流水工具套件 · 多来源合并说明')
    lines.append('生成时间：' + time.strftime('%Y-%m-%d %H:%M:%S'))
    lines.append('')
    lines.append('一、输入文件（{} 个）'.format(len(tables)))
    for p in tables:
        lines.append('  - {}'.format(os.path.basename(p)))
    if skipped:
        lines.append('  无法解析、已跳过：{} 个'.format(len(skipped)))
    lines.append('')
    lines.append('二、合并结果')
    lines.append('  拼接后原始行：{}'.format(stats['rows_raw']))
    lines.append('  去重（R5）：{} 行'.format(stats['dropped_dup']))
    lines.append('  入表行：{}'.format(stats['rows_kept']))
    counts = stats.get('source_counts') or {}
    if counts:
        lines.append('  「来源」分布：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in counts.items()))
    lines.append('  合并后按交易时间稳定排序')
    lines.append('')
    lines.append('三、订单号一致性检查（多来源合并后的冲突提示）')
    if not conflicts or not conflicts.get('groups'):
        lines.append('  未发现同一订单号自相矛盾的记录。')
    else:
        lines.append('  对不上记录组：{} 组（其中两侧金额不一致 {} 组、同向重复 {} 组、'
                     '主体不一致 {} 组）'.format(
                         conflicts['groups'], conflicts['amount_mismatch'],
                         conflicts['same_dir_dup'], conflicts['party_mismatch']))
        lines.append('  说明：一笔交易在两侧账单里各有一条记录（付款方记「支出」、收款方记'
                     '「收入」）属正常，「两侧金额不一致」通常说明某一侧口径不同或数据有问题。')
        for s in conflicts.get('samples') or []:
            lines.append('    · ' + s)
        if conflicts['groups'] > len(conflicts.get('samples') or []):
            lines.append('    ……共 {} 组，仅显示前 {} 组'.format(
                conflicts['groups'], CONFLICT_LIMIT))
    lines.append('')
    lines.append('四、跳过/失败的文件')
    if not skipped:
        lines.append('  无。')
    for p, err in skipped:
        lines.append('  - {}：{}'.format(os.path.basename(p), err))
    lines.append('')
    lines.append('五、说明与局限')
    lines.append('  本报告由程序按固定规则自动生成，仅用于数据整理与后续分析，'
                 '不构成审计结论、鉴定意见或法律意见。')
    lines.append('  订单号检查只做「同一订单号内部自相矛盾」的提示，不会修改任何金额；'
                 '是否采纳请以原始账单为准。')
    if out_tsv:
        lines.append('  产物：{}'.format(os.path.basename(out_tsv)))
    return '\n'.join(lines) + '\n'


def run_merge(inputs=None, folders=None, out_dir=None, progress=None, dedupe=True, session=True):
    """headless 可调用：读取多张表 → 合并 → 去重 → 排序 → 写 02_合并流水.tsv。"""
    progress = progress or NullProgress(logger.log)
    tables = collect_tables(inputs, folders)
    if not tables:
        raise ValueError('没有找到可合并的流水文件。\n支持 .tsv / .txt / .csv / .xlsx / .xls，'
                         '可多选文件或选择存储这些文件的文件夹。')
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(tables[0]))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    logger.log('功能 B 多来源合并：{} 个输入文件 → {}'.format(len(tables), out_dir))

    frames, skipped = [], []
    total = len(tables)
    progress.set_determinate(total, '正在读取流水…')
    for i, p in enumerate(tables):
        progress.update(i, total, '读取 {}'.format(os.path.basename(p)))
        try:
            one = read_table(p)
        except Exception as e:
            skipped.append((p, str(e)))
            logger.log('跳过 {}：{}'.format(p, e))
            continue
        if one is None or one.empty:
            skipped.append((p, '无数据行'))
            logger.log('跳过 {}：无数据行'.format(p))
            continue
        src = one[COL_SOURCE].astype(str).str.strip()
        blank = src.eq('')
        if bool(blank.any()):
            one[COL_SOURCE] = src.where(~blank, os.path.splitext(os.path.basename(p))[0])
        logger.log('读取 {}：{} 行'.format(os.path.basename(p), len(one)))
        frames.append(one)
    if not frames:
        raise ValueError('全部 {} 个输入文件都无法解析，没有可合并的数据。'.format(len(skipped)))

    progress.set_indeterminate('正在合并与去重…')
    df = pd.concat(frames, ignore_index=True)
    rows_raw = len(df)
    conflicts = detect_conflicts(df)
    dropped_dup = 0
    if dedupe:
        df, dropped_dup = dedupe_frame(df)
    df = sort_by_time(df)

    counts = df[COL_SOURCE].astype(str).value_counts()
    source_counts = dict((str(k), int(v)) for k, v in counts.items())
    sources = sorted(source_counts)

    tsv = paths.product(out_dir, paths.P_MERGE_TSV)
    report = paths.product(out_dir, paths.P_MERGE_REPORT)
    stats = {'files_ok': len(frames), 'files_failed': len(skipped), 'rows_raw': rows_raw,
             'rows_kept': len(df), 'dropped_dup': dropped_dup, 'sources': sources,
             'source_counts': source_counts, 'conflicts': conflicts}
    paths.backup_if_exists(tsv)
    write_tsv(df, tsv)
    paths.backup_if_exists(report)
    with io.open(report, 'w', encoding='utf-8') as f:
        f.write(format_report(stats, tables, skipped, conflicts, tsv))
    logger.log('已写出：{}（{} 行，去重 {} 行，来源 {} 个）'.format(tsv, len(df), dropped_dup, len(sources)))
    if conflicts.get('groups'):
        logger.log('订单号一致性检查：{} 组记录对不上（详见 {}）'.format(
            conflicts['groups'], os.path.basename(report)))

    if session:
        note = '{} 个文件 / {} 行 → {} 行（去重 {}）'.format(
            len(frames), rows_raw, len(df), dropped_dup)
        if conflicts.get('groups'):
            note += '；订单号冲突 {} 组'.format(conflicts['groups'])
        paths.update_session(out_dir, 'merge', 'done',
                             outputs={'tsv': tsv, 'report': report},
                             args={'inputs': len(tables), 'dedupe': dedupe}, note=note)
    progress.update(1, 1, '完成')
    return {'out_dir': out_dir, 'tsv': tsv, 'report': report, 'stats': stats, 'rows': len(df),
            'skipped': skipped, 'conflicts': conflicts, 'log': logger.log_path()}


def open_window(parent, state=None):
    """功能窗口：选择多份统一格式流水 → 合并为一个 tsv。"""
    state = state if state is not None else {}
    from ui import dialogs
    from ui.progress import ProgressWindow
    from ui.worker import TaskRunner

    win = tk.Toplevel(parent)
    win.title('多来源合并 · 资金流水工具套件')
    win.resizable(False, False)
    try:
        win.transient(parent)
    except Exception:
        pass
    body = ttk.Frame(win, padding=14)
    body.pack(fill='both', expand=True)

    ttk.Label(body, text='待合并的统一格式流水（可多选，或选择文件夹递归查找）',
              font=('', 10, 'bold')).pack(anchor='w')
    ttk.Label(body, text='支持：01_统一流水.tsv、03_流水总表.xlsx / .csv，以及任何前 4 列为'
                         '「用户方 | 支出/收入 | 客户方 | 金额」的表。',
              wraplength=560, foreground='#555555', justify='left').pack(anchor='w', pady=(2, 4))
    listbox = tk.Listbox(body, width=78, height=7, activestyle='none')
    listbox.pack(fill='both', expand=True, pady=(4, 6))
    selected = []

    def refresh():
        listbox.delete(0, 'end')
        for p in selected:
            listbox.insert('end', p)

    def add_files():
        for p in dialogs.ask_open_files(win, '选择流水文件（可多选）', dialogs.FT_TSV):
            if p not in selected:
                selected.append(p)
        refresh()

    def add_folder():
        d = dialogs.ask_directory(win, '选择文件夹（递归查找流水表）', state.get('last_dir'))
        if d and d not in selected:
            selected.append(d)
            state['last_dir'] = d
            refresh()

    bar = ttk.Frame(body)
    bar.pack(fill='x')
    ttk.Button(bar, text='添加流水文件…', command=add_files).pack(side='left')
    ttk.Button(bar, text='添加文件夹…', command=add_folder).pack(side='left', padx=6)
    ttk.Button(bar, text='清空', command=lambda: (selected.__delitem__(slice(None)), refresh())
               ).pack(side='left')

    opt = ttk.LabelFrame(body, text='合并选项', padding=8)
    opt.pack(fill='x', pady=10)
    var_dedupe = tk.BooleanVar(value=True)
    ttk.Checkbutton(opt, text='去除重复记录（订单号优先，规则 R5）', variable=var_dedupe).pack(anchor='w')
    ttk.Label(opt, text='合并后按交易时间稳定排序；「来源」列留空的记录自动填入文件名。',
              foreground='#555555').pack(anchor='w', pady=(4, 0))

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
            dialogs.show_error(win, '未选择输入', '请先添加要合并的流水文件或文件夹。')
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
        pw = ProgressWindow(win, title='正在合并', text='准备中…')

        def job(progress):
            return run_merge(inputs=files, folders=folders, out_dir=out_dir,
                             progress=progress, dedupe=var_dedupe.get())

        def done(res):
            s = res['stats']
            extra = ''
            if res['skipped']:
                extra = '\n跳过 {} 个无法解析的文件（详见运行日志）。'.format(len(res['skipped']))
            cf = res.get('conflicts') or {}
            if cf.get('groups'):
                extra += ('\n⚠ 发现 {} 组同一订单号对不上（两侧金额不一致 {} 组、同向重复 {} 组、'
                          '主体不一致 {} 组），详见「合并说明」。').format(
                              cf['groups'], cf['amount_mismatch'], cf['same_dir_dup'],
                              cf['party_mismatch'])
            dialogs.ResultDialog(
                win, title='合并完成',
                summary='读取 {} 个文件 / {} 行，去重 {} 行，合并结果 {} 行，来源 {} 个。{}'.format(
                    s['files_ok'], s['rows_raw'], s['dropped_dup'], res['rows'],
                    len(s['sources']), extra),
                items=[('合并流水', res['tsv']), ('合并说明', res['report'])],
                out_dir=res['out_dir'], log_file=res['log'],
                next_label='下一步：流水总表与透视表',
                on_next=lambda: _next_step(parent, state, 'pivot'))

        def fail(err, trace):
            logger.log('失败：{}'.format(err))
            dialogs.show_error(win, '合并失败', err, trace)

        runner.start(job, on_done=done, on_error=fail, progress=pw)

    bottom = ttk.Frame(body)
    bottom.pack(fill='x')
    ttk.Button(bottom, text='开始合并', command=start).pack(side='left')
    ttk.Button(bottom, text='关闭', command=win.destroy).pack(side='right')
    dialogs.center(win, parent)
    return win


def _next_step(parent, state, tid):
    from core import registry
    try:
        registry.load_tool(tid).open_window(parent, state)
    except Exception as e:
        logger.log('打开下一步失败：{}'.format(e))


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        args = sys.argv[1:]
        folders = [a for a in args if os.path.isdir(a)]
        files = [a for a in args if os.path.isfile(a)]
        r = run_merge(inputs=files, folders=folders)
        print('OK', r['tsv'], r['rows'], '行')
