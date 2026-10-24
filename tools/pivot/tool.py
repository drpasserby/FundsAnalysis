# -*- coding: utf-8 -*-
"""功能 C：流水总表与数据透视表 —— 统一格式流水 → 汇总 xlsx / csv / 透视表。

输入：统一格式表（推荐 `02_合并流水.tsv`，也可直接是 `01_统一流水.tsv` 或总表 xlsx）。
输出（同一输出目录）：
  `03_流水总表.xlsx`   —— 统一格式全量明细（sheet「流水总表」）
  `03_流水总表.csv`    —— 同一份明细的 csv
  `03_数据透视表.xlsx` —— sheet「数据透视表」（用户方 / 用户侧角色 / 支出·收入 / 客户方 /
                          对手侧角色 / 求和项:金额(元) / 笔数；两个角色列写的是 VLOOKUP 公式）
                          sheet「按主体汇总」（每个主体的支出、收入、净流入，按净流入降序）
                          sheet「人物关系图」（只有表头「序号 / 姓名 / 角色」，留空供人工填写；
                          填好后「数据透视表」的两个角色列会自动带出角色）
"""
import os
import tkinter as tk
from tkinter import ttk

import pandas as pd

from core import log as logger
from core import paths
from core.excel_style import style_worksheet
from core.progress import NullProgress
from core.schema import (COL_AMOUNT, COL_CP, COL_DIR, COL_USER, DIR_RECV, DIR_SPEND, HINT_COLUMNS,
                         blank_columns, read_table, write_csv, write_xlsx)

SHEET_DETAIL = '流水总表'
SHEET_PIVOT = '数据透视表'
SHEET_SUMMARY = '按主体汇总'
SHEET_RELATION = '人物关系图'
AMOUNT_HEAD = '求和项:金额(元)'
COUNT_HEAD = '笔数'
COL_USER_ROLE = '用户侧角色'
COL_CP_ROLE = '对手侧角色'
COLS_PIVOT = [COL_USER, COL_USER_ROLE, COL_DIR, COL_CP, COL_CP_ROLE, AMOUNT_HEAD, COUNT_HEAD]
COLS_SUMMARY = ['主体', '支出合计(元)', '支出笔数', '收入合计(元)', '收入笔数', '净流入(元)']
COLS_RELATION = ['序号', '姓名', '角色']


def role_formula(cell, relation_sheet=SHEET_RELATION):
    """角色列的 VLOOKUP 公式：在「人物关系图」的 B(姓名) : C(角色) 里查名字，带出角色。"""
    return '=VLOOKUP({0},{1}!B:C,2,0)'.format(cell, relation_sheet)


def _ensure_clean_index(df):
    """分组前的防御：索引唯一 + 列名唯一。

    重复列名会让 `df[col]` 返回 DataFrame，pandas 分组随即报
    `InvalidIndexError: Reindexing only valid with uniquely valued Index objects`；
    重复索引则在 reindex 类操作上报同样的错。
    """
    if not df.columns.is_unique:
        logger.log('[WARN] 检测到重复列名，已按首次出现保留')
        df = df.loc[:, ~df.columns.duplicated()]
    if not df.index.is_unique:
        logger.log('[WARN] 检测到重复索引，已重置')
        df = df.reset_index(drop=True)
    return df


def build_pivot(df, relation_sheet=SHEET_RELATION):
    """按 用户方 / 支出·收入 / 客户方 汇总金额与笔数，并写入角色 VLOOKUP 公式。

    列顺序刻意与旧版微信工具一致：用户侧账号名称 | 用户侧角色 | 借贷类型 | 对手侧账户名称 |
    对手侧角色 | 求和项:交易金额(分)，因此 A 列、D 列是姓名，可直接沿用
    `=VLOOKUP(A6,人物关系图!B:C,2,0)` 这类写法。
    """
    df = _ensure_clean_index(df)
    g = df.groupby([COL_USER, COL_DIR, COL_CP], as_index=False)[COL_AMOUNT].agg(['sum', 'count'])
    g.columns = [COL_USER, COL_DIR, COL_CP, AMOUNT_HEAD, COUNT_HEAD]
    g[AMOUNT_HEAD] = g[AMOUNT_HEAD].round(2)
    g[COUNT_HEAD] = g[COUNT_HEAD].astype(int)
    g = g.sort_values([COL_USER, COL_DIR, COL_CP]).reset_index(drop=True)

    rows = []
    for i, r in enumerate(g.itertuples(index=False), start=2):
        rows.append([r[0], role_formula('A{}'.format(i), relation_sheet), r[1], r[2],
                     role_formula('D{}'.format(i), relation_sheet), r[3], r[4]])
    return pd.DataFrame(rows, columns=COLS_PIVOT)


def build_summary(df):
    """按主体汇总：支出/收入合计与笔数、净流入（按净流入降序）。"""
    df = _ensure_clean_index(df)
    subjects = sorted(set(df[COL_USER].astype(str)) | set(df[COL_CP].astype(str)))
    spend = df[df[COL_DIR] == DIR_SPEND].groupby(COL_USER)[COL_AMOUNT].agg(['sum', 'count'])
    recv = df[df[COL_DIR] == DIR_RECV].groupby(COL_USER)[COL_AMOUNT].agg(['sum', 'count'])
    rows = []
    for s in subjects:
        if s in spend.index:
            sp_sum, sp_cnt = spend.loc[s, 'sum'], spend.loc[s, 'count']
        else:
            sp_sum, sp_cnt = 0.0, 0
        if s in recv.index:
            rc_sum, rc_cnt = recv.loc[s, 'sum'], recv.loc[s, 'count']
        else:
            rc_sum, rc_cnt = 0.0, 0
        rows.append([s, round(float(sp_sum), 2), int(sp_cnt), round(float(rc_sum), 2), int(rc_cnt),
                     round(float(rc_sum) - float(sp_sum), 2)])
    out = pd.DataFrame(rows, columns=COLS_SUMMARY)
    return out.sort_values('净流入(元)', ascending=False).reset_index(drop=True)


def build_relation():
    """人物关系图模板：只有表头「序号 / 姓名 / 角色」，内容留空由人工填写。"""
    return pd.DataFrame([], columns=COLS_RELATION)


def run_pivot(input_path, out_dir=None, progress=None, session=True):
    """headless 可调用：读统一格式流水 → 写 03_流水总表.xlsx / .csv / 03_数据透视表.xlsx。"""
    progress = progress or NullProgress(logger.log)
    if not input_path or not os.path.isfile(input_path):
        raise ValueError('没有可用的输入流水文件：{}'.format(input_path))
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(input_path))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    logger.log('功能 C 流水总表与透视表：{} → {}'.format(input_path, out_dir))

    progress.set_indeterminate('正在读取流水…')
    df = read_table(input_path)
    df = _ensure_clean_index(df)
    logger.log('[INFO] 输入 {}：shape={} 列名={}'.format(
        os.path.basename(input_path), tuple(df.shape), list(df.columns)))
    if df.empty:
        raise ValueError('输入文件没有数据行：{}'.format(input_path))
    logger.log('读取 {} 行'.format(len(df)))

    empty_hint = blank_columns(df, HINT_COLUMNS)
    if empty_hint:
        logger.log('[WARN] 输入里这些列整列为空：{}'.format('、'.join(empty_hint)))
        logger.log('[WARN] 多半是导入时没有勾选这些列（功能 A 的「选择需要写入的列」），'
                   '总表里对应列也会空白；需要的话请回功能 A 重新导入并勾选，再走一次合并与总表。')

    progress.set_indeterminate('正在汇总…')
    pivot = build_pivot(df)
    summary = build_summary(df)
    relation = build_relation()
    subjects = len(set(df[COL_USER].astype(str)) | set(df[COL_CP].astype(str)))

    xlsx = paths.product(out_dir, paths.P_PIVOT_XLSX)
    csv = paths.product(out_dir, paths.P_PIVOT_CSV)
    pvx = paths.product(out_dir, paths.P_PIVOT_PIVOT)
    for p in (xlsx, csv, pvx):
        paths.backup_if_exists(p)

    progress.set_indeterminate('正在写出总表…')
    write_xlsx(df, xlsx, sheet_name=SHEET_DETAIL)
    write_csv(df, csv)

    progress.set_indeterminate('正在写出透视表…')
    with pd.ExcelWriter(pvx, engine='openpyxl') as w:
        pivot.to_excel(w, sheet_name=SHEET_PIVOT, index=False)
        summary.to_excel(w, sheet_name=SHEET_SUMMARY, index=False)
        relation.to_excel(w, sheet_name=SHEET_RELATION, index=False)
        for ws in w.sheets.values():
            style_worksheet(ws)
        try:
            # 角色列是公式，打开 Excel 时强制重算（避免显示为 0 或 #N/A 缓存值）
            w.book.calculation.fullCalcOnLoad = True
        except Exception:
            pass
    logger.log('已写出：{}、{}、{}'.format(os.path.basename(xlsx), os.path.basename(csv),
                                          os.path.basename(pvx)))

    stats = {'rows': len(df), 'subjects': subjects, 'pivot_rows': len(pivot),
             'channels': len(pivot[[COL_USER, COL_CP]].drop_duplicates()),
             'empty_columns': empty_hint}
    if session:
        paths.update_session(out_dir, 'pivot', 'done',
                             outputs={'xlsx': xlsx, 'csv': csv, 'pivot': pvx},
                             args={'input': os.path.basename(input_path)},
                             note='{} 行 / {} 个主体 / {} 条通道'.format(
                                 stats['rows'], stats['subjects'], stats['channels']))
    progress.update(1, 1, '完成')
    return {'out_dir': out_dir, 'xlsx': xlsx, 'csv': csv, 'pivot_xlsx': pvx,
            'stats': stats, 'rows': len(df), 'log': logger.log_path()}


def guess_input(state=None):
    """从上次输出目录里猜测可用的输入流水（合并流水优先，其次统一流水，最后总表）。"""
    state = state or {}
    out_dir = state.get('last_out_dir') or paths.find_recent_output_dir(paths.default_output_base())
    if not out_dir:
        return ''
    for name in (paths.P_MERGE_TSV, paths.P_IMPORT_TSV, paths.P_PIVOT_XLSX):
        p = paths.product(out_dir, name)
        if os.path.isfile(p):
            return p
    return ''


def open_window(parent, state=None):
    """功能窗口：选择统一格式流水 → 生成总表与透视表。"""
    state = state if state is not None else {}
    from ui import dialogs
    from ui.progress import ProgressWindow
    from ui.worker import TaskRunner

    win = tk.Toplevel(parent)
    win.title('流水总表与透视表 · 资金流水工具套件')
    win.resizable(False, False)
    try:
        win.transient(parent)
    except Exception:
        pass
    body = ttk.Frame(win, padding=14)
    body.pack(fill='both', expand=True)

    ttk.Label(body, text='输入统一格式流水', font=('', 10, 'bold')).pack(anchor='w')
    ttk.Label(body, text='推荐使用上一步的 02_合并流水.tsv；也可直接选 01_统一流水.tsv 或已有的流水总表。',
              wraplength=560, foreground='#555555', justify='left').pack(anchor='w', pady=(2, 4))
    var_in = tk.StringVar(value=state.get('last_input') if os.path.isfile(str(state.get('last_input') or '')) else guess_input(state))
    row = ttk.Frame(body)
    row.pack(fill='x', pady=(4, 10))
    ttk.Entry(row, textvariable=var_in, width=58).pack(side='left', fill='x', expand=True)
    ttk.Button(row, text='浏览…', command=lambda: var_in.set(
        dialogs.ask_open_file(win, '选择统一格式流水', dialogs.FT_TSV) or var_in.get())
    ).pack(side='left', padx=(6, 0))

    ttk.Label(body, text='输出目录（留空则自动新建 输出_时间戳）').pack(anchor='w')
    var_out = tk.StringVar(value=state.get('last_out_dir') or '')
    row2 = ttk.Frame(body)
    row2.pack(fill='x', pady=(4, 10))
    ttk.Entry(row2, textvariable=var_out, width=58).pack(side='left', fill='x', expand=True)
    ttk.Button(row2, text='浏览…', command=lambda: var_out.set(
        dialogs.ask_directory(win, '选择输出目录', var_out.get() or None) or var_out.get())
    ).pack(side='left', padx=(6, 0))

    ttk.Label(body, text='将生成：03_流水总表.xlsx（明细）、03_流水总表.csv、'
                         '03_数据透视表.xlsx（数据透视表 / 按主体汇总 / 人物关系图）。\n'
                         '「数据透视表」含「用户侧角色 / 对手侧角色」两列，写的是 '
                         'VLOOKUP 公式；在「人物关系图」里填好「姓名 + 角色」（只需两列）后，'
                         '角色会自动带出。',
              wraplength=560, foreground='#555555', justify='left').pack(anchor='w')

    runner = TaskRunner(win)

    def start():
        src = var_in.get().strip().strip('"')
        if not src or not os.path.isfile(src):
            dialogs.show_error(win, '输入无效', '请选择一个存在的统一格式流水文件。')
            return
        if runner.busy():
            return
        out_dir = var_out.get().strip().strip('"')
        if not out_dir:
            out_dir = paths.new_output_dir(os.path.dirname(os.path.abspath(src)))
        logger.init_log(out_dir)
        state['last_input'] = src
        state['last_out_dir'] = out_dir
        pw = ProgressWindow(win, title='正在汇总', text='准备中…')

        def job(progress):
            return run_pivot(src, out_dir=out_dir, progress=progress)

        def done(res):
            s = res['stats']
            summary = ('共 {} 行明细，{} 个主体，{} 条通道。\n'
                       '「数据透视表」的 B 列（用户侧角色）与 E 列（对手侧角色）已写入 '
                       'VLOOKUP 公式；打开「人物关系图」填好「姓名 / 角色」后角色自动带出。').format(
                           s['rows'], s['subjects'], s['channels'])
            if s.get('empty_columns'):
                summary += '\n注意：输入文件里「{}」整列为空，总表对应列也会是空白；' \
                           '多半是导入时没勾选这些列，请回功能 A 重新导入并勾选后再走一遍。'.format(
                               '、'.join(s['empty_columns']))
            dialogs.ResultDialog(
                win, title='总表与透视表完成',
                summary=summary,
                items=[('流水总表 (xlsx)', res['xlsx']), ('流水总表 (csv)', res['csv']),
                       ('数据透视表', res['pivot_xlsx'])],
                out_dir=res['out_dir'], log_file=res['log'],
                next_label='下一步：资金流向分析',
                on_next=lambda: _next_step(parent, state, 'flow'))

        def fail(err, trace):
            logger.log('失败：{}'.format(err))
            dialogs.show_error(win, '汇总失败', err, trace)

        runner.start(job, on_done=done, on_error=fail, progress=pw)

    bottom = ttk.Frame(body)
    bottom.pack(fill='x', pady=(10, 0))
    ttk.Button(bottom, text='开始生成', command=start).pack(side='left')
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
        r = run_pivot(sys.argv[1], out_dir=sys.argv[2] if len(sys.argv) > 2 else None)
        print('OK', r['xlsx'], r['rows'], '行')
