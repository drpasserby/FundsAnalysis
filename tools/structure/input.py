# -*- coding: utf-8 -*-
"""试验性模块的输入适配：统一流水、微信原始总表、汇总透视表。"""
import os

import pandas as pd

from core import log as logger
from core import schema
from core.schema import COL_AMOUNT, COL_CP, COL_DIR, COL_USER, clean_text_value

KIND_CANONICAL = '统一格式明细'
KIND_RAW_WECHAT = '微信原始总表'
KIND_PIVOT = '汇总透视表'


def _headers(raw):
    return dict((schema._norm_col_name(c), c) for c in raw.columns)


def _pick(cols, names):
    for name in names:
        if name in cols:
            return cols[name]
    return None


def _amount_column(cols):
    for name, real in cols.items():
        if '金额' in name and '笔数' not in name:
            return real, ('fen' if '(分)' in name or '（分）' in name else 'yuan')
    return None, None


def _mapping(raw):
    """返回 (输入类型, 列映射, 金额单位)；无法识别返回 None。"""
    cols = _headers(raw)
    user = _pick(cols, (COL_USER, '用户侧账号名称'))
    cp = _pick(cols, (COL_CP, '对手侧账户名称'))
    direction = _pick(cols, (COL_DIR, '借贷类型'))
    if not user or not cp or not direction:
        return None
    amount, unit = _amount_column(cols)
    if amount is None:
        return None
    amount_name = schema._norm_col_name(amount)
    if (all(c in cols for c in (COL_USER, COL_DIR, COL_CP))
            and amount_name in (COL_AMOUNT, '金额（元）', '金额(元)')):
        kind = KIND_CANONICAL
    elif '交易金额(分)' in cols or '交易金额（分）' in cols:
        kind = KIND_RAW_WECHAT
    else:
        kind = KIND_PIVOT
    return kind, {COL_USER: user, COL_CP: cp, COL_DIR: direction, COL_AMOUNT: amount}, unit


def _project(raw, found, sheet):
    kind, cols, unit = found
    out = pd.DataFrame({
        COL_USER: raw[cols[COL_USER]].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_DIR: raw[cols[COL_DIR]].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_CP: raw[cols[COL_CP]].map(lambda v: clean_text_value(v) if pd.notna(v) else ''),
        COL_AMOUNT: schema._clean_amount(raw[cols[COL_AMOUNT]]),
    })
    out[COL_DIR] = out[COL_DIR].replace({'出': schema.DIR_SPEND, '入': schema.DIR_RECV})
    if unit == 'fen':
        out[COL_AMOUNT] = out[COL_AMOUNT] / 100.0
    info = {'kind': kind, 'sheet': sheet, 'unit': '分（已换算为元）' if unit == 'fen' else '元'}
    return out.reset_index(drop=True), info


def read_input(path):
    """读取受支持输入并投影为 用户方 | 支出/收入 | 客户方 | 金额（元）。

    Excel 会逐 sheet 识别表头，优先使用逐笔明细；没有明细时才使用汇总透视表。
    原始微信总表的「交易金额(分)」在此显式除以 100，不经过统一格式归一。
    """
    if not path or not os.path.isfile(path):
        raise schema.SchemaError('文件不存在：{}'.format(path))
    ext = os.path.splitext(path)[1].lower()
    candidates = []
    if ext in schema.EXCEL_EXTS:
        try:
            with pd.ExcelFile(path) as book:
                for sheet in book.sheet_names:
                    head = pd.read_excel(book, sheet_name=sheet, nrows=0)
                    found = _mapping(head)
                    if found:
                        priority = {KIND_CANONICAL: 0, KIND_RAW_WECHAT: 1, KIND_PIVOT: 2}[found[0]]
                        candidates.append((priority, sheet, found))
                if not candidates:
                    raise schema.SchemaError(
                        '没有找到可识别的流水 sheet。支持统一流水、微信原始总表、'
                        '以及含「用户方/客户方/方向/金额」的汇总透视表。')
                candidates.sort(key=lambda item: (item[0], book.sheet_names.index(item[1])))
                priority, sheet, found = candidates[0]
                raw = pd.read_excel(book, sheet_name=sheet)
        except schema.SchemaError:
            raise
        except Exception as e:
            raise schema.SchemaError('读取 Excel 失败：{}\n{}'.format(path, e))
    elif ext in schema.TSV_EXTS + schema.CSV_EXTS:
        raw = schema._read_text(path, ',' if ext == '.csv' else '\t')
        found = _mapping(raw)
        if not found:
            raise schema.SchemaError(
                '无法识别流水表头：{}。需要用户方/客户方/方向/金额，'
                '或微信原始总表的用户侧账号名称/对手侧账户名称/借贷类型/交易金额(分)。'.format(
                    '、'.join(str(c) for c in raw.columns)))
        sheet = ''
    else:
        raise schema.SchemaError('不支持的文件类型：{}（支持 {}）'.format(
            ext or '无扩展名', '、'.join(schema.SUPPORTED_EXTS)))
    df, info = _project(raw, found, sheet)
    logger.log('[INFO] 试验性模块输入：{}；sheet={}；金额单位={}；{} 行'.format(
        info['kind'], info['sheet'] or '(文本)', info['unit'], len(df)))
    return df, info
