# -*- coding: utf-8 -*-
"""统一流水格式（canonical schema）—— 套件内部唯一的接口契约。

**列顺序固定**（产物、合并结果、总表一律按此顺序；未勾选导入、来源没有、整批为空的列
直接不写出，其余列的相对顺序不变）：

    用户ID | 用户方 | 支出/收入 | 对手方ID | 客户方 | 金额（元） | 交易时间 | 订单号 | 来源
    交易类型 | 交易业务类型 | 交易状态 | 说明 | 对手方接收金额 | 对手侧银行名称

其中「用户方 / 客户方」是**姓名**（支付宝从 `用户信息` / `交易对方信息` 里截取），
「用户ID / 对手方ID」是**原始账号标识**（微信取 `用户ID` / `对手方ID`，支付宝取
`用户信息` / `交易对方信息` 原文，形如 `2088122934548576(许文仙)`）——两个来源因此占同一组列、
由 `来源` 列区分；两套编号体系不同，跨来源同一人仍按姓名对齐。

读取一律**按列名**（`COL_*` 常量）认列，不依赖列位；只有 v1.3.2 的历史文件
（前 4 列 = 用户方 | 支出/收入 | 客户方 | 金额）才由 `normalize()` 的列位兜底兼容。
"""
import io
import os
import re

import pandas as pd

from core import log as _log

# ---- 必需列（语义必需；读取按列名识别，列顺序见 ORDER_MAIN）----
COL_USER = '用户方'
COL_DIR = '支出/收入'
COL_CP = '客户方'
COL_AMOUNT = '金额'
CORE_COLUMNS = [COL_USER, COL_DIR, COL_CP, COL_AMOUNT]

# ---- 账号标识列：微信/支付宝导入时**强制写入**（界面不可取消），保证总表始终带原始账号 ----
COL_USER_ID = '用户ID'            # 原始账号标识：微信「用户ID」/ 支付宝「用户信息」原文
COL_CP_ID = '对手方ID'            # 原始账号标识：微信「对手方ID」/ 支付宝「交易对方信息」原文
ID_COLUMNS = [COL_USER_ID, COL_CP_ID]

# ---- 其余扩展列 ----
COL_TIME = '交易时间'
COL_TYPE = '交易类型'            # 语义类型：转账 / 提现 / 消费 / 退款 / 充值 / 手续费 …
COL_BIZ_TYPE = '交易业务类型'    # 支付渠道：快捷支付 / 余额支付 / 其它支付（财付通账单里 100% 有值）
COL_STATUS = '交易状态'
COL_MEMO = '说明'
COL_ORDER = '订单号'
COL_RECV_AMOUNT = '对手方接收金额'  # 元；来自账单「对手方接收金额(分)」，只有「出」方向记录带此字段
COL_CP_BANK = '对手侧银行名称'
COL_SOURCE = '来源'

# ---- 列顺序（用户指定，m04268）：用户ID 用户方 借贷类型 对手方ID 客户方 交易金额
#      交易时间 订单号 来源 + 其他信息（其余扩展列，保持相对顺序）----
ORDER_MAIN = [COL_USER_ID, COL_USER, COL_DIR, COL_CP_ID, COL_CP, COL_AMOUNT, COL_TIME, COL_ORDER,
              COL_SOURCE]
ORDER_REST = [COL_TYPE, COL_BIZ_TYPE, COL_STATUS, COL_MEMO, COL_RECV_AMOUNT, COL_CP_BANK]
ALL_COLUMNS = ORDER_MAIN + ORDER_REST
EXT_COLUMNS = [c for c in ALL_COLUMNS if c not in CORE_COLUMNS]

# 导入时固定保留的列（必需 4 列 + 账号标识 2 列，按产物顺序）；空值不参与「整批为空自动省略」
LOCKED_COLUMNS = [c for c in ALL_COLUMNS if c in set(CORE_COLUMNS) | set(ID_COLUMNS)]

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


def _norm_col_name(name):
    """列名规范化：去掉 BOM/零宽等不可见字符、全角空格转半角、去首尾空白。"""
    s = re.sub(INVISIBLE_RE, '', str(name)).replace('\u3000', ' ')
    return s.strip()


def clean_text_value(value):
    """单值文本清洗：去掉不可见字符后去首尾空白（解析层逐字段使用）。"""
    if value is None:
        return ''
    return re.sub(INVISIBLE_RE, '', str(value)).strip()

EXCEL_EXTS = ('.xlsx', '.xlsm', '.xltx', '.xls')
TSV_EXTS = ('.tsv', '.txt')
CSV_EXTS = ('.csv',)
SUPPORTED_EXTS = EXCEL_EXTS + TSV_EXTS + CSV_EXTS

DIR_DESC = '用户方 | 支出/收入 | 客户方 | 金额（元）（4 个必需列，按列名识别）'


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


# 用户最容易「以为勾了、其实导入时没写出来」的列：整列为空时给一句提示
HINT_COLUMNS = (COL_TIME, COL_ORDER, COL_SOURCE, COL_TYPE, COL_MEMO)


def blank_columns(df, columns=None):
    """返回整列空白（含缺列）的列名。

    用途：导入时少勾了几列，`normalize()` 会把缺失列补成空串，后续步骤看到的
    就是「列头在、值全空」（合并流水/总表里那一列整列空白）。这里把情况说清楚，
    提示用户回功能 A 重新导入并勾选这些列。
    """
    cols = list(columns) if columns else list(df.columns)
    out = []
    for c in cols:
        if c not in df.columns:
            out.append(c)
            continue
        col = df[c]
        if col.isna().all():
            out.append(c)
            continue
        # read_csv 会把空字段读成 NaN，所以「空」要同时认 NaN 与空串
        s = col.astype(str).str.strip()
        if s.isin(['', 'nan', 'NaN', 'None']).all():
            out.append(c)
    return out


# ---- 文本编码探测 ----------------------------------------------------------
# 严格解码：第一个成功的编码就采用，绝不用 errors='replace' 把真实编码错误吞掉
# （否则 ANSI/GBK 文件会被当成"看起来成功的乱码 UTF-8"读进来，表头与数据全错位）。
TEXT_ENCODINGS = ('utf-8-sig', 'gb18030', 'utf-16')
ENCODING_FALLBACK = 'utf-8-replace'


def _looks_utf16(raw):
    """无 BOM 的 UTF-16：ASCII 文本里每隔一个字节就是 0x00（纯中文段更少，阈值放宽）。"""
    head = raw[:400]
    if not head:
        return False
    return head.count(0) >= max(2, len(head) // 20)


def decode_text(raw, path=''):
    """把字节串解码成文本，返回 (text, 编码名)。

    顺序：BOM 判定（utf-8-sig / utf-16）→ 无 BOM 的 UTF-16 启发 → 严格尝试
    utf-8-sig → gb18030 → utf-16；全部失败才用 utf-8 + errors='replace' 兜底并告警。
    """
    if raw.startswith(b'\xef\xbb\xbf'):
        return raw.decode('utf-8-sig'), 'utf-8-sig'
    if raw.startswith(b'\xff\xfe') or raw.startswith(b'\xfe\xff'):
        return raw.decode('utf-16'), 'utf-16'
    if _looks_utf16(raw):
        for enc in ('utf-16-le', 'utf-16-be'):
            try:
                text = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            if text.count('\x00') <= max(1, len(text) // 100):
                return text, enc
    last_err = None
    for enc in TEXT_ENCODINGS:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError as e:
            last_err = e
    _log.log('[WARN] 编码无法识别，已用 utf-8 + errors=replace 兜底（文本可能乱码）：{}'
             .format(path or '(内存数据)'))
    return raw.decode('utf-8', errors='replace'), ENCODING_FALLBACK


def normalize(df, keep_extra=False):
    """把任意来源的 DataFrame 归一到统一格式。

    - 列名已含 4 个必需列名 → 视为统一格式，按 `ALL_COLUMNS` 顺序重排；
    - 否则**按列位**取前 4 列命名（仅用于 v1.3.2 的历史文件：
      前 4 列 = 用户方 | 支出/收入 | 客户方 | 金额）；
    - 补齐缺失的扩展列，丢弃 ALL_COLUMNS 之外的列（keep_extra 时保留）；
    - 文本列 strip、金额列转数值。
    """
    if df is None:
        raise SchemaError('数据为空。')
    df = df.copy()
    # 防御：重复列名会让 df[col] 返回 DataFrame，pandas 分组即报
    # InvalidIndexError: Reindexing only valid with uniquely valued Index objects
    if not df.columns.is_unique:
        dups = sorted(set(str(c) for c in df.columns if list(df.columns).count(c) > 1))
        _log.log('[WARN] 检测到重复列名，已按首次出现保留：{}'.format('、'.join(dups)))
        df = df.loc[:, ~df.columns.duplicated()].copy()
    if not df.index.is_unique:
        _log.log('[WARN] 检测到重复索引，已重置')
        df = df.reset_index(drop=True)
    if df.shape[1] < len(CORE_COLUMNS):
        raise SchemaError('数据列数不足：当前共 {} 列，至少需要 4 列。\n请确认数据格式为：{}'.format(
            df.shape[1], DIR_DESC))
    cols = list(df.columns)
    # 列名归一：先去掉 BOM/零宽字符/首尾空白/全角空格再和标准列名比对。
    # 否则「交易时间 」（尾部多一个空格）这种列名会被当成未知列直接丢掉，
    # 而标准列又被补成空列 —— 表现就是「源文件明明有时间列，总表里却是空白」。
    rename = {}
    for c in df.columns:
        std = _norm_col_name(c)
        if std != c and std in set(ALL_COLUMNS) and std not in df.columns:
            rename[c] = std
    if rename:
        _log.log('[WARN] 列名含不可见字符或多余空白，已按标准列名归一：{}'.format(
            '、'.join('{} → {}'.format(k, v) for k, v in rename.items())))
        df = df.rename(columns=rename)
        cols = list(df.columns)
    if not all(c in cols for c in CORE_COLUMNS):
        rename = {}
        for i, c in enumerate(cols[:len(CORE_COLUMNS)]):
            rename[c] = CORE_COLUMNS[i]
        df = df.rename(columns=rename)

    extras = []
    if keep_extra:
        extras = [c for c in df.columns if c not in ALL_COLUMNS]
    else:
        unknown = [str(c) for c in df.columns if c not in ALL_COLUMNS]
        if unknown:
            _log.log('[WARN] 丢弃无法识别的列（列名与标准列名不一致）：{}'.format(
                '、'.join(unknown)))
    for c in ALL_COLUMNS:
        if c not in df.columns:
            df[c] = ''
    df = df[ALL_COLUMNS + extras].copy()

    for c in [COL_USER, COL_DIR, COL_CP] + [x for x in EXT_COLUMNS if x not in NUMERIC_EXT_COLUMNS]:
        df[c] = _clean_text(df[c])
    df[COL_AMOUNT] = _clean_amount(df[COL_AMOUNT])
    for c in NUMERIC_EXT_COLUMNS:
        df[c] = _clean_amount(df[c])
    bad = [str(c) for c in df.columns if '\ufffd' in str(c) or '\ufeff' in str(c)]
    if bad:
        _log.log('[WARN] 列名里发现乱码或 BOM 残留字符（常见于文件编码不对）：{}'.format('、'.join(bad)))
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
        _log.log('[INFO] 读取 {} (Excel, 行数: {})'.format(path, len(raw) + 1))
    elif ext in CSV_EXTS:
        raw = _read_text(path, ',')
    elif ext in TSV_EXTS:
        raw = _read_text(path, '\t')
    else:
        raise SchemaError('不支持的文件类型：{}（支持 {}）'.format(
            ext or '无扩展名', '、'.join(SUPPORTED_EXTS)))
    _log.log('[INFO] {} → shape={} 列名={}'.format(
        os.path.basename(path), tuple(raw.shape), list(raw.columns)))
    return normalize(raw)


def _read_text(path, sep):
    """文本表格读取：先做编码探测（严格解码）再把文本交给 pandas 解析。

    编码一律按 utf-8-sig → gb18030 → utf-16 严格尝试，失败才 replace 兜底并告警；
    不再让 pandas 各编码试错，避免 errors='replace' 把编码错误掩盖成"成功读取"。
    """
    with open(path, 'rb') as f:
        raw_bytes = f.read()
    text, enc = decode_text(raw_bytes, path)
    line_count = len(text.splitlines())
    _log.log('[INFO] 读取 {} (编码: {}, 行数: {})'.format(path, enc, line_count))
    try:
        return pd.read_csv(io.StringIO(text), sep=sep, header=0, dtype=str)
    except Exception as e:
        raise SchemaError('解析失败：{}\n{}'.format(path, e))


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
