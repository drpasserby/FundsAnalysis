# -*- coding: utf-8 -*-
"""跨来源通用清洗：去重（规则 R5）与按时间排序。

放在 core 是因为功能 A（导入后）与功能 B（合并后）都要用同一套判定，
避免两处逻辑漂移。
"""
import pandas as pd

from core.schema import COL_AMOUNT, COL_CP, COL_DIR, COL_ORDER, COL_TIME, COL_USER


def combo_key(df):
    """R5 组合键：交易时间|金额|用户方|客户方|方向。"""
    return (df[COL_TIME].astype(str) + '|' + df[COL_AMOUNT].astype(str) + '|'
            + df[COL_USER].astype(str) + '|' + df[COL_CP].astype(str) + '|'
            + df[COL_DIR].astype(str))


def dedupe_frame(df):
    """去重：订单号（非空时）优先，其次组合键。返回 (df, 删除行数)。"""
    if df is None or df.empty:
        return df, 0
    before = len(df)
    order = df[COL_ORDER].astype(str).str.strip().replace('', pd.NA)
    # 订单号只在「同一主体、同一方向」内判重：同一笔交易的支出方与收入方记录订单号相同，
    # 但用户方/方向不同，属于两条都应保留的对手记录。
    order_key = order + '|' + df[COL_USER].astype(str) + '|' + df[COL_DIR].astype(str)
    dup = order_key.duplicated(keep='first') & order.notna()
    dup |= combo_key(df).duplicated(keep='first')
    out = df[~dup].reset_index(drop=True)
    return out, before - len(out)


def sort_by_time(df):
    """按交易时间稳定排序；时间为空的记录排在最后。"""
    if df is None or df.empty:
        return df
    key = df[COL_TIME].astype(str).str.strip()
    key = key.where(key != '', '9999-99-99 99:99:99')
    order = key.sort_values(kind='mergesort').index
    return df.loc[order].reset_index(drop=True)
