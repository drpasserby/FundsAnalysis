# -*- coding: utf-8 -*-
"""统一流水格式（canonical schema）—— 套件内部唯一的接口契约。

规范：**前 4 列严格固定**，与主工具 v1.3.2「按列位读取、不依赖表头文字」的逻辑完全兼容；
扩展列一律放在后面，旧流程忽略它们也不会出错。

    用户方 | 支出/收入 | 客户方 | 金额
    用户ID | 对手方ID | 交易时间 | 交易类型 | 交易业务类型 | 交易状态 | 说明 | 订单号 |
    对手方接收金额 | 对手侧银行名称 | 来源

其中「用户方 / 客户方」是**姓名**（支付宝从 `用户信息` / `交易对方信息` 里截取），
「用户ID / 对手方ID」是**原始账号标识**（微信取 `用户ID` / `对手方ID`，支付宝取
`用户信息` / `交易对方信息` 原文，形如 `2088122934548576(许文仙)`）——
两个来源因此能在同一组列上对齐，`来源` 列区分数据出处。
"""
import os
import re

import pandas as pd

# ---- 必需列（前 4 列，顺序不可变）----
COL_USER = '用户方'
COL_DIR = '支出/收入'
COL_CP = '客户方'
COL_AMOUNT = '金额'
CORE_COLUMNS = [COL_USER, COL_DIR, COL_CP, COL_AMOUNT]

# ---- 扩展列（可选，统一放在必需列之后）----
COL_USER_ID = '用户ID'            # 原始账号标识：微信「用户ID」/ 支付宝「用户信息」原文
COL_CP_ID = '对手方ID'            # 原始账号标识：微信「对手方ID」/ 支付宝「交易对方信息」原文
COL_TIME = '交易时间'
COL_TYPE = '交易类型'            # 语义类型：转账 / 提现 / 消费 / 退款 / 充值 / 手续费 …
COL_BIZ_TYPE = '交易业务类型'    # 支付渠道：快捷支付 / 余额支付 / 其它支付（财付通账单里 100% 有值）
COL_STATUS = '交易状态'
COL_MEMO = '说明'
COL_ORDER = '订单号'
COL_RECV_AMOUNT = '对手方接收金额'  # 元；来自账单「对手方接收金额(分)」，只有「出」方向记录带此字段
COL_CP_BANK = '对手侧银行名称'
COL_SOURCE = '来源'
EXT_COLUMNS = [COL_USER_ID, COL_CP_ID, COL_TIME, COL_TYPE, COL_BIZ_TYPE, COL_STATUS, COL_MEMO,
               COL_ORDER, COL_RECV_AMOUNT, COL_CP_BANK, COL_SOURCE]

ALL_COLUMNS = CORE_COLUMNS + EXT_COLUMNS

# 账号标识列：微信/支付宝导入时**强制写入**（界面不可取消），保证总表始终带原始账号、可跨来源对齐
ID_COLUMNS = [COL_USER_ID, COL_CP_ID]
# 导入时固定保留的列（必需 4 列 + 账号标识 2 列）；空值不参与「整批为空自动省略」
LOCKED_COLUMNS = CORE_COLUMNS + ID_COLUMNS

# 数值型扩展列（其余扩展列都是文本）
NUMERIC_EXT_COLUMNS = (COL_RECV_AMOUNT,)

DIR_SPEND = '支出'
DIR_RECV = '收入'
DIRS = (DIR_SPEND, DIR_RECV)

ENCODING = 'utf-8-sig'
SHEET_MAIN = '流水总表'

# 不可见字符：零宽空格/连接符、双向控制符、BOM、不换行空格。
# 实测：财付通账单「备注2」每一行都以 \u200c\u200e\u202a\u202c 开头，去掉这些字符后
# 171214 行里有 119160 行其实是空值 —— 不清掉会污染「说明」列，也会让空值判断
# （去重、来源列、合计）失效。
INVISIBLE_RE = r'[\u200b-\u200f\u202a-\u202e\u2060-\u2064\u2066-\u2069\ufeff\u00a0]'


def clean_text_value(value):
    """单值文本清洗：去掉不可见字符后去首尾空白（解析层逐字段使用）。"""
    if value is None:
        return ''
    return re.sub(INVISIBLE_RE, '', str(value)).strip()

EXCEL_EXTS = ('.xlsx', '.xlsm', '.xltx', '.xls')
TSV_EXTS = ('.tsv', '.txt')
CSV_EXTS = ('.csv',)
SUPPORTED_EXTS = EXCEL_EXTS + TSV_EXTS + CSV_EXTS

DIR_DESC = '用户方 | 支出/收入 | 客户方 | 金额（元）'


class SchemaError(ValueError):
    """统一格式校验失败（供 UI 直接展示中文原因）。"""


def empty_frame():
    """空的标准流水表（含全部列）。"""
    return pd.DataFrame(columns=ALL_COLUMNS)


def _clean_text(series):
    """NaN → 空串，去不可见字符与首尾空白（避免 astype(str) 造出 'nan'）。"""
    s = series.where(series.notna(), '').astype(str)
    s = s.str.replace(INVISIBLE_RE, '', regex=True)
    return s.str.strip()


def _clean_amount(series):
    """金额清洗：去不可见字符、千分位与货币符号后转数值，无法解析的记为 NaN。"""
    s = series.where(series.notna(), '').astype(str).str.strip()
    s = s.str.replace(INVISIBLE_RE, '', regex=True)
    for ch in (',', '"', "'", '¥', '￥', ' '):
        s = s.str.replace(ch, '')
    return pd.to_numeric(s, errors='coerce')


def normalize(df, keep_extra=False):
    """把任意来源的 DataFrame 归一到统一格式。

    - 列名已含 4 个必需列名 → 视为统一格式，直接使用；
    - 否则**按列位**取前 4 列命名（兼容 v1.3.2 的历史文件）；
    - 补齐缺失的扩展列，丢弃 ALL_COLUMNS 之外的列（keep_extra 时保留）；
    - 文本列 strip、金额列转数值。
    """
    if df is None:
        raise SchemaError('数据为空。')
    df = df.copy()
    if df.shape[1] < len(CORE_COLUMNS):
        raise SchemaError('数据列数不足：当前共 {} 列，至少需要 4 列。\n请确认数据格式为：{}'.format(
            df.shape[1], DIR_DESC))
    cols = list(df.columns)
    if not all(c in cols for c in CORE_COLUMNS):
        rename = {}
        for i, c in enumerate(cols[:len(CORE_COLUMNS)]):
            rename[c] = CORE_COLUMNS[i]
        df = df.rename(columns=rename)

    extras = []
    if keep_extra:
        extras = [c for c in df.columns if c not in ALL_COLUMNS]
    for c in ALL_COLUMNS:
        if c not in df.columns:
            df[c] = ''
    df = df[ALL_COLUMNS + extras].copy()

    for c in [COL_USER, COL_DIR, COL_CP] + [x for x in EXT_COLUMNS if x not in NUMERIC_EXT_COLUMNS]:
        df[c] = _clean_text(df[c])
    df[COL_AMOUNT] = _clean_amount(df[COL_AMOUNT])
    for c in NUMERIC_EXT_COLUMNS:
        df[c] = _clean_amount(df[c])
    return df.reset_index(drop=True)


def read_table(path):
    """读取受支持的流水文件（xlsx/xls/tsv/txt/csv）并归一到统一格式。"""
    if not path or not os.path.exists(path):
        raise SchemaError('文件不存在：{}'.format(path))
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXTS:
        try:
            raw = pd.read_excel(path, header=0)
        except Exception as e:
            raise SchemaError('读取 Excel 失败：{}\n{}'.format(path, e))
    elif ext in CSV_EXTS:
        raw = _read_text(path, ',')
    elif ext in TSV_EXTS:
        raw = _read_text(path, '\t')
    else:
        raise SchemaError('不支持的文件类型：{}（支持 {}）'.format(
            ext or '无扩展名', '、'.join(SUPPORTED_EXTS)))
    return normalize(raw)


def _read_text(path, sep):
    """文本表格读取，编码按 utf-8-sig → gbk → utf-8 回退。"""
    last_err = None
    for enc in (ENCODING, 'gbk', 'utf-8'):
        try:
            return pd.read_csv(path, sep=sep, header=0, dtype=str, encoding=enc)
        except UnicodeDecodeError as e:
            last_err = e
        except Exception as e:
            raise SchemaError('解析失败：{}\n{}'.format(path, e))
    raise SchemaError('编码无法识别（已尝试 utf-8-sig / gbk / utf-8）：{}'.format(last_err))


def _ensure_dir(path):
    d = os.path.dirname(os.path.abspath(path))
    if d and not os.path.isdir(d):
        os.makedirs(d, exist_ok=True)


def write_tsv(df, path):
    """写 TSV（utf-8-sig，Excel 可直接双击打开）。"""
    _ensure_dir(path)
    df.to_csv(path, sep='\t', index=False, encoding=ENCODING, lineterminator='\n')
    return path


def write_csv(df, path):
    """写 CSV（utf-8-sig）。"""
    _ensure_dir(path)
    df.to_csv(path, index=False, encoding=ENCODING)
    return path


def write_xlsx(df, path, sheet_name=SHEET_MAIN, style=True):
    """写 xlsx（openpyxl），可选统一表头样式。"""
    _ensure_dir(path)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        if style:
            from core.excel_style import style_worksheet
            style_worksheet(writer.sheets[sheet_name])
    return path


def fen_to_yuan(series, digits=2):
    """金额单位换算：分 → 元。"""
    return (pd.to_numeric(series, errors='coerce') / 100.0).round(digits)
