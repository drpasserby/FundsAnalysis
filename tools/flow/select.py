# -*- coding: utf-8 -*-
"""功能 D 的「输入识别 + 人物筛选」（纯逻辑，无 UI 依赖，可被自检脚本直接调用）。

两种输入都能直接分析
--------------------
* `03_流水总表.xlsx` —— **逐笔明细**（13 列统一格式，带交易时间）；
* `03_数据透视表.xlsx` —— 功能 C 的**汇总透视表**，第一个 sheet「数据透视表」本身就是
  `用户方 | 用户侧角色 | 支出/收入 | 客户方 | 对手侧角色 | 求和项:金额(元) | 笔数`，
  即「谁 ↔ 谁、方向、合计金额、笔数」，**正好是资金流向图需要的边数据**（研判只看两人之间的
  汇总流水，不需要时间）。因此这里把它的前 4 个有效列投影成流向分析要的
  `用户方 | 支出/收入 | 客户方 | 金额`，不再要求用户改选明细。

另外两个 sheet 供筛选使用：「按主体汇总」（姓名/笔数/收支合计）与「人物关系图」（姓名/角色）。

本模块负责四件事：

1. `detect_source()`：判断输入是逐笔明细还是汇总透视表；
2. `resolve_input()` / `load_pivot_detail()`：给出实际分析用的文件与数据帧；
3. `load_relation()` / `relation_names()` / `subject_table()` / `load_summary()`：
   读「人物关系图」的姓名-角色，或列出流水里的主体，供人工挑选关注对象；
4. `apply_filters()`：按「关注的人 + 金额下限 + 时间范围」裁剪，
   默认只看选中人物**彼此之间**的往来（双方都要在名单里，`MODE_STRICT`），
   也可放宽成保留选中人物的全部往来（含直接对手，`MODE_EXPAND`）。
"""
import os

import pandas as pd

from core import paths
from core.schema import (COL_AMOUNT, COL_CP, COL_DIR, COL_TIME, COL_USER, DIR_SPEND,
                         clean_text_value)

KIND_DETAIL = 'detail'
KIND_PIVOT = 'pivot'
KIND_UNKNOWN = 'unknown'

SHEET_DETAIL = '流水总表'
SHEET_PIVOT = '数据透视表'
SHEET_SUMMARY = '按主体汇总'
SHEET_RELATION = '人物关系图'

MODE_EXPAND = 'expand'
MODE_STRICT = 'strict'
MODE_LABELS = [
    (MODE_STRICT, '只看选中人物之间的往来（两边都要在名单里）'),
    (MODE_EXPAND, '保留选中人物的往来（含名单外的直接对手）'),
]

COL_SUBJECT = '主体'
COL_COUNT = '笔数'
SUBJECT_COLUMNS = [COL_SUBJECT, COL_COUNT, '支出合计(元)', '收入合计(元)', '金额合计(元)']

# 「数据透视表」sheet 里的金额列名（pandas 读到的名字形如「求和项:金额(元)」）
AMOUNT_HEAD_HINTS = ('求和项:金额', '求和项：金额', '金额')


# ================= 1. 输入识别 =================

def xlsx_sheets(path):
    """返回 xlsx 的 sheet 名列表；不是 xlsx 或读不了则返回 []。"""
    if not path or os.path.splitext(str(path))[1].lower() not in ('.xlsx', '.xlsm'):
        return []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return []
    try:
        return list(wb.sheetnames)
    except Exception:
        return []
    finally:
        try:
            wb.close()
        except Exception:
            pass


def detect_source(path):
    """识别输入文件。返回 {'kind', 'sheets', 'detail_path', 'note'}。

    kind=detail  ：逐笔明细（流水总表 xlsx / 统一流水 tsv / 合并流水 tsv / v1.3.2 的 4 列 xls）
    kind=pivot   ：功能 C 的汇总产物（数据透视表 / 人物关系图 sheet）
    kind=unknown ：认不出来（交给下游按默认列位尝试）
    """
    path = str(path or '')
    if not path:
        return {'kind': KIND_UNKNOWN, 'sheets': [], 'detail_path': '', 'note': '尚未选择文件'}
    if not os.path.isfile(path):
        return {'kind': KIND_UNKNOWN, 'sheets': [], 'detail_path': '',
                'note': '文件不存在：{}'.format(path)}

    base = os.path.basename(path)
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm', '.xls'):
        sheets = xlsx_sheets(path)
        if SHEET_RELATION in sheets or SHEET_PIVOT in sheets or SHEET_SUMMARY in sheets:
            return {'kind': KIND_PIVOT, 'sheets': sheets, 'detail_path': path,
                    'note': '汇总透视表（「{}」sheet：用户方 ↔ 客户方 的合计金额与笔数，可直接分析；'
                            '不含逐笔时间）'.format(SHEET_PIVOT)}
        if SHEET_DETAIL in sheets:
            return {'kind': KIND_DETAIL, 'sheets': sheets, 'detail_path': path,
                    'note': '逐笔明细（sheet「{}」，含交易时间）'.format(SHEET_DETAIL)}
        joined = '、'.join(sheets[:4]) if sheets else '（读不到 sheet 名）'
        return {'kind': KIND_DETAIL, 'sheets': sheets, 'detail_path': path,
                'note': 'Excel 表格（sheet：{}）'.format(joined)}

    if ext in ('.tsv', '.txt', '.csv'):
        kind = KIND_DETAIL
        if base == paths.P_MERGE_TSV:
            note = '合并流水明细'
        elif base == paths.P_IMPORT_TSV or base == paths.P_IMPORT_TSV_ALIPAY:
            note = '统一流水明细'
        else:
            note = '文本表格（{}）'.format(ext.lstrip('.'))
        return {'kind': kind, 'sheets': [], 'detail_path': path, 'note': note}

    return {'kind': KIND_UNKNOWN, 'sheets': [], 'detail_path': '',
            'note': '未识别的文件类型：{}'.format(ext or '无扩展名')}


def resolve_input(path):
    """整理输入：返回 (kind, 实际分析用的文件, 说明文字)。

    汇总透视表（`03_数据透视表.xlsx`）**不需要**改选明细：它的「数据透视表」sheet 就是
    「用户方 ↔ 客户方 + 方向 + 合计金额」，由 `load_pivot_detail()` 投影成流向分析要的 4 列。
    """
    info = detect_source(path)
    return info['kind'], path, info['note']


def resolve_detail(path):
    """兼容旧调用：返回 (文件路径, 说明文字)。现在不再替换文件。"""
    kind, target, note = resolve_input(path)
    return target, note


def pivot_amount_column(raw):
    """从透视表 sheet 的列名里挑金额列（形如「求和项:金额(元)」）。找不到返回 None。"""
    if raw is None:
        return None
    names = [str(c) for c in raw.columns]
    for hint in AMOUNT_HEAD_HINTS:
        for name in names:
            if hint in name and '笔数' not in name:
                return name
    for name in names:
        if '金额' in name:
            return name
    return None


def load_pivot_detail(path, sheet=SHEET_PIVOT):
    """把 C 的汇总透视表投影成流向分析要的 4 列：用户方 | 支出/收入 | 客户方 | 金额。

    列按名称识别（不依赖列位）：用户方 / 支出·收入方向 / 客户方 / 求和项:金额。角色列是
    Excel 公式（pandas 读到的是公式文本），这里直接丢弃；笔数列也丢弃。
    缺关键列时抛 ValueError，并把实际列名写进信息里。
    """
    raw = pd.read_excel(path, sheet_name=sheet)
    if raw is None or raw.empty:
        raise ValueError('「{}」sheet「{}」里没有数据。'.format(os.path.basename(str(path)), sheet))
    cols = dict((str(c).strip(), c) for c in raw.columns)
    user_col = cols.get(COL_USER)
    cp_col = cols.get(COL_CP)
    dir_col = cols.get(COL_DIR)
    if dir_col is None:
        for name, real in cols.items():
            if '支出' in name or '收入' in name:
                dir_col = real
                break
    amt_col = pivot_amount_column(raw)
    missing = [label for label, real in ((COL_USER, user_col), (COL_DIR, dir_col),
                                         (COL_CP, cp_col), ('金额', amt_col)) if real is None]
    if missing:
        raise ValueError('透视表缺少必要列 {}（实际列名：{}）。'.format(
            '、'.join(missing), '、'.join(str(c) for c in raw.columns)))
    out = pd.DataFrame({
        COL_USER: raw[user_col].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_DIR: raw[dir_col].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_CP: raw[cp_col].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_AMOUNT: pd.to_numeric(raw[amt_col], errors='coerce'),
    })
    return out.reset_index(drop=True)


def detail_sheet_names():
    """返回 xlsx 里可能存明细的 sheet 名（供调试与自检）。"""
    return [SHEET_DETAIL]


# ================= 2. 主体与角色 =================

def subject_table(df, top=None):
    """列出流水里出现过的全部主体：笔数、支出合计、收入合计、金额合计。按笔数降序。

    主体 = 作为「用户方」或「客户方」出现过的名字。
    """
    if df is None or df.empty:
        return pd.DataFrame([], columns=SUBJECT_COLUMNS)
    users = df[COL_USER].astype(str).map(clean_text_value)
    cps = df[COL_CP].astype(str).map(clean_text_value)
    amounts = pd.to_numeric(df[COL_AMOUNT], errors='coerce').fillna(0.0)

    base = pd.DataFrame({'主体': pd.concat([users, cps], ignore_index=True),
                         '金额': pd.concat([amounts, amounts], ignore_index=True)})
    g = base.groupby('主体', as_index=False).agg(**{COL_COUNT: ('金额', 'size'),
                                                    '金额合计(元)': ('金额', 'sum')})
    spend = df[df[COL_DIR] == DIR_SPEND]
    spend_g = (pd.DataFrame({'主体': spend[COL_USER].astype(str).map(clean_text_value),
                             '金额': pd.to_numeric(spend[COL_AMOUNT], errors='coerce').fillna(0.0)})
               .groupby('主体')['金额'].sum())
    recv = df[df[COL_DIR] != DIR_SPEND]
    recv_g = (pd.DataFrame({'主体': recv[COL_USER].astype(str).map(clean_text_value),
                            '金额': pd.to_numeric(recv[COL_AMOUNT], errors='coerce').fillna(0.0)})
              .groupby('主体')['金额'].sum())
    g['支出合计(元)'] = g['主体'].map(spend_g).fillna(0.0).round(2)
    g['收入合计(元)'] = g['主体'].map(recv_g).fillna(0.0).round(2)
    g['金额合计(元)'] = g['金额合计(元)'].round(2)
    g = g[SUBJECT_COLUMNS].sort_values([COL_COUNT, '金额合计(元)'], ascending=False)
    g = g.reset_index(drop=True)
    if top:
        g = g.head(int(top)).reset_index(drop=True)
    return g


def _num_col(raw, name):
    """取一列数值（缺列补 0），用于读「按主体汇总」。"""
    if name not in raw.columns:
        return pd.Series([0.0] * len(raw), index=raw.index)
    return pd.to_numeric(raw[name], errors='coerce').fillna(0.0)


def load_summary(path):
    """读功能 C 产物 `03_数据透视表.xlsx` 的「按主体汇总」sheet → 主体表。

    这样「选关注的人」时不必再读 7 万行明细：同一个文件里就有姓名、笔数与收支合计。
    没有该 sheet 时返回空 DataFrame。
    """
    empty = pd.DataFrame([], columns=SUBJECT_COLUMNS)
    if not path or not os.path.isfile(path) or SHEET_SUMMARY not in xlsx_sheets(path):
        return empty
    try:
        raw = pd.read_excel(path, sheet_name=SHEET_SUMMARY)
    except Exception:
        return empty
    if raw is None or raw.empty or COL_SUBJECT not in raw.columns:
        return empty
    spend, recv = _num_col(raw, '支出合计(元)'), _num_col(raw, '收入合计(元)')
    spend_n, recv_n = _num_col(raw, '支出笔数'), _num_col(raw, '收入笔数')
    out = pd.DataFrame({
        COL_SUBJECT: raw[COL_SUBJECT].map(lambda v: clean_text_value(v)).astype(str),
        COL_COUNT: (spend_n + recv_n).astype(int),
        '支出合计(元)': spend.round(2),
        '收入合计(元)': recv.round(2),
        '金额合计(元)': (spend + recv).round(2),
    })
    out = out[out[COL_SUBJECT].str.strip() != '']
    return out.sort_values([COL_COUNT, '金额合计(元)'], ascending=False).reset_index(drop=True)


def load_relation(path):
    """读功能 C 产物 `03_数据透视表.xlsx` 的「人物关系图」sheet → DataFrame[姓名, 角色]。

    文件不存在、没有该 sheet、或人工还没填内容时返回空 DataFrame（不报错）。
    """
    empty = pd.DataFrame([], columns=['姓名', '角色'])
    if not path or not os.path.isfile(path):
        return empty
    if SHEET_RELATION not in xlsx_sheets(path):
        return empty
    try:
        raw = pd.read_excel(path, sheet_name=SHEET_RELATION, dtype=str)
    except Exception:
        return empty
    if raw is None or raw.empty:
        return empty
    cols = {str(c).strip(): c for c in raw.columns}
    name_col = cols.get('姓名') or (list(raw.columns)[1] if raw.shape[1] > 1 else None)
    role_col = cols.get('角色') or (list(raw.columns)[2] if raw.shape[1] > 2 else None)
    if name_col is None:
        return empty
    out = pd.DataFrame({
        '姓名': raw[name_col].map(lambda v: clean_text_value(v) if pd.notna(v) else '').astype(str),
        '角色': raw[role_col].map(lambda v: clean_text_value(v) if pd.notna(v) else '').astype(str)
        if role_col is not None else '',
    })
    out = out[out['姓名'].str.strip() != '']
    return out.reset_index(drop=True)


def relation_map(relation):
    """角色 → 姓名列表（保留填写顺序）。"""
    groups = {}
    if relation is None or relation.empty:
        return groups
    for r in relation.itertuples(index=False):
        name = str(r[0]).strip()
        role = str(r[1]).strip() or '(未填角色)'
        if name:
            groups.setdefault(role, []).append(name)
    return groups


def relation_names(relation, roles=None):
    """「人物关系图」里的姓名列表（按填写顺序；roles 非空时只取这些角色）。"""
    if relation is None or relation.empty:
        return []
    want = set(str(r).strip() for r in (roles or []) if str(r).strip())
    out = []
    for r in relation.itertuples(index=False):
        name = str(r[0]).strip()
        role = str(r[1]).strip() if len(r) > 1 else ''
        if not name:
            continue
        if want and role not in want:
            continue
        if name not in out:
            out.append(name)
    return out


# ================= 3. 筛选 =================

def apply_filters(df, names=None, min_amount=None, time_from=None, time_to=None, mode=MODE_STRICT):
    """按关注对象/金额/时间裁剪。返回 (新 DataFrame, info dict)。

    names       ：关注的人（默认只有**双方都在名单里**的记录才保留，即 MODE_STRICT）
    mode        ：MODE_STRICT 只看选中人物彼此之间的往来；MODE_EXPAND 保留选中人物的全部往来（含名单外对手）
    min_amount  ：金额下限（含）
    time_from/time_to：日期（YYYY-MM-DD 或更长的日期时间串，按前 10 位比较）；输入无时间列时自动跳过
    """
    info = {'rows_before': 0 if df is None else len(df), 'rows_after': 0 if df is None else len(df),
            'names': 0, 'mode': mode, 'min_amount': 0.0, 'time_from': '', 'time_to': '',
            'dropped_names': 0, 'dropped_amount': 0, 'dropped_time': 0, 'no_time_column': False}
    if df is None or df.empty:
        return df, info

    picked = set()
    for n in (names or []):
        s = str(n).strip()
        if s:
            picked.add(s)
    out = df
    if picked:
        info['names'] = len(picked)
        u = out[COL_USER].astype(str)
        c = out[COL_CP].astype(str)
        if mode == MODE_STRICT:
            mask = u.isin(picked) & c.isin(picked)
        else:
            mask = u.isin(picked) | c.isin(picked)
        info['dropped_names'] = int((~mask).sum())
        out = out[mask]

    try:
        lower = float(min_amount or 0)
    except (TypeError, ValueError):
        lower = 0.0
    if lower > 0:
        info['min_amount'] = lower
        amt = pd.to_numeric(out[COL_AMOUNT], errors='coerce')
        keep = amt >= lower
        info['dropped_amount'] = int((~keep).sum())
        out = out[keep]

    tf, tt = str(time_from or '').strip()[:10], str(time_to or '').strip()[:10]
    if tf or tt:
        info['time_from'], info['time_to'] = tf, tt
        if COL_TIME not in out.columns:
            info['no_time_column'] = True
        else:
            t = out[COL_TIME].astype(str).str.slice(0, 10)
            keep = pd.Series(True, index=out.index)
            if tf:
                keep = keep & (t >= tf)
            if tt:
                keep = keep & (t <= tt)
            info['dropped_time'] = int((~keep).sum())
            out = out[keep]

    out = out.reset_index(drop=True)
    info['rows_after'] = len(out)
    info['subjects_after'] = len(set(out[COL_USER].astype(str)) | set(out[COL_CP].astype(str))) \
        if not out.empty else 0
    return out, info


def is_active(info):
    """这次筛选是否真的起作用（用于决定要不要在界面/报告里标注）。"""
    if not info:
        return False
    return bool(info.get('names') or info.get('min_amount') or info.get('time_from') or info.get('time_to'))


def summary(info):
    """把筛选结果拼成一句话（写日志、session.json 与报告页抬头）。"""
    if not info or not is_active(info):
        return '未启用筛选（全部 {} 行）'.format(info.get('rows_before', 0) if info else 0)
    parts = []
    if info.get('names'):
        mode_label = dict(MODE_LABELS).get(info.get('mode'), '')
        parts.append('关注人物 {} 个（{}）'.format(info['names'], mode_label))
    if info.get('min_amount'):
        parts.append('金额 ≥ {:g} 元'.format(info['min_amount']))
    if info.get('time_from') or info.get('time_to'):
        parts.append('时间 {} ~ {}'.format(info.get('time_from') or '起', info.get('time_to') or '止'))
    parts.append('{} 行 → {} 行'.format(info.get('rows_before', 0), info.get('rows_after', 0)))
    if info.get('subjects_after') is not None:
        parts.append('主体 {} 个'.format(info.get('subjects_after')))
    return '；'.join(parts)
