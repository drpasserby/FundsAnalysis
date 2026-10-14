# -*- coding: utf-8 -*-
"""支付宝账户明细解析与清洗（功能 A 的支付宝解析层）。

输入：支付宝「账户明细」导出的 csv（**GBK 编码、逗号分隔、17 列固定表头**，
      文件名形如 `741719073_账户明细d2_20251223183605_part1.csv`）。
输出：统一格式 DataFrame（core/schema.py 契约）+ 清洗报告数据。

与微信侧**共用同一套契约与规则编号**（方案 §4），保证两种来源的产物可以直接合并：
  R1 自转自（用户方 == 客户方）去除
  R2 「收/支」不是「支出/收入」的记录去除
     （实测「其它」= 余额宝-自动转入/转出，对手方即本人，属自己账户间划转）
  R4 只保留转账类（可选，默认关）。支付宝没有财付通的「交易用途类型」列，
     本层按「消费名称」判定：含「转账」视为转账类，第三档再排掉余额宝/余利宝/小荷包等理财类
  R5 去重（可选，默认关）：订单号 + 用户方 + 方向，否则 时间+金额+双方+方向
  R6 单位统一：账单「金额（元）」已是元，本层统一四舍五入到 2 位小数（附对账）
  R7 空值 / 非正金额去除

姓名提取：账单里「用户信息」「交易对方信息」形如 `2088122934548576(许文仙)`（带前导制表符），
统一取末尾括号内的名称；没有括号的（如 `花呗`、`陈慧娴`）保留原值。
来源列固定写「支付宝」（COL_SOURCE）。
"""
import csv
import io
import os
import re

import pandas as pd

from core.schema import (ALL_COLUMNS, COL_AMOUNT, COL_BIZ_TYPE, COL_CP, COL_DIR, COL_MEMO, COL_ORDER,
                         COL_SOURCE, COL_STATUS, COL_TIME, COL_TYPE, COL_USER, DIR_RECV, DIR_SPEND,
                         EXCEL_EXTS, EXT_COLUMNS, clean_text_value)

SOURCE_NAME = '支付宝'

KEY_ORDER = '交易号'
KEY_MERCHANT_ORDER = '商户订单号'
KEY_TIME = '交易创建时间'
KEY_PAY_TIME = '付款时间'
KEY_MODIFY_TIME = '最近修改时间'
KEY_AREA = '交易来源地'
KEY_TYPE = '类型'
KEY_USER = '用户信息'
KEY_COUNTERPARTY = '交易对方信息'
KEY_PRODUCT = '消费名称'
KEY_AMOUNT = '金额（元）'
KEY_DIR = '收/支'
KEY_STATUS = '交易状态'
KEY_PAY_METHOD = '支付方式'
KEY_CLEARING = '清算流水号'
KEY_REMARK = '备注'
KEY_CHECK_ID = '对应的协查数据'

# 列名候选（不同导出版本的括号全半角、列名略有差异，按顺序取第一个存在的）
ORDER_KEYS = (KEY_ORDER, '交易单号')
USER_KEYS = (KEY_USER, '用户信息（账号）')
COUNTERPARTY_KEYS = (KEY_COUNTERPARTY, '交易对方')
PRODUCT_KEYS = (KEY_PRODUCT, '商品名称', '交易内容')
AMOUNT_KEYS = (KEY_AMOUNT, '金额(元)', '金额')
DIR_KEYS = (KEY_DIR, '收支')
TIME_KEYS = (KEY_TIME, '交易时间')
TYPE_KEYS = (KEY_TYPE, '交易类型')
STATUS_KEYS = (KEY_STATUS,)
PAY_METHOD_KEYS = (KEY_PAY_METHOD,)
REMARK_KEYS = (KEY_REMARK, '备注1')
AREA_KEYS = (KEY_AREA,)
CHECK_ID_KEYS = (KEY_CHECK_ID, '协查数据')

NEED_KEYS = (KEY_ORDER, KEY_USER, KEY_COUNTERPARTY, KEY_PRODUCT, KEY_AMOUNT, KEY_DIR)

# R4 三档：不过滤 / 只保留「转账」/ 在上一档基础上再排除余额宝·余利宝·小荷包等理财类
R4_OFF = 'off'
R4_TRANSFER = 'transfer'
R4_P2P = 'p2p'
R4_CHOICES = (R4_OFF, R4_TRANSFER, R4_P2P)
R4_LABELS = {
    R4_OFF: '不过滤（保留全部记录）',
    R4_TRANSFER: '只保留转账类（消费名称含「转账」）',
    R4_P2P: '只保留对外转账（再去掉余额宝/余利宝/小荷包等理财类）',
}
TRANSFER_KEYWORD = '转账'
WEALTH_KEYWORDS = ('余额宝', '余利宝', '小荷包', '自动转入', '自动攒')

# 「收/支」→ 支出/收入（实测取值：支出/收入/其它）
DIR_MAP = {
    '支出': DIR_SPEND, '出': DIR_SPEND, '付款': DIR_SPEND,
    '收入': DIR_RECV, '入': DIR_RECV, '收款': DIR_RECV,
}

HEADER_SCAN_BYTES = 4000
# `2088122934548576(许文仙)` → 取末尾括号里的名称
NAME_RE = re.compile(r'[（(]\s*([^（()）]*?)\s*[)）]\s*$')


class ParseError(ValueError):
    """账户明细无法解析。"""


def _decode(path):
    """按 utf-8-sig → gbk → utf-8(replace) 顺序解码（严格解码，避免乱码掩盖真实编码）。"""
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in ('utf-8-sig', 'gbk'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('utf-8', errors='replace')


def _looks_like_header(cells):
    """一行是否是表头（同一文件若拼接了多段明细，每段自带表头行）。"""
    s = set(c.strip() for c in cells)
    return KEY_ORDER in s and (KEY_USER in s or KEY_AMOUNT in s)


def _name_from_cell(value):
    """`2088122934548576(许文仙)` → `许文仙`；`花呗` / `陈慧娴` → 原值。"""
    s = clean_text_value(value)
    if not s:
        return ''
    m = NAME_RE.search(s)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return s


def parse_bill_file(path):
    """解析单个账户明细 csv。

    返回 (header, rows, skipped_header_rows)；无法识别表头时返回 (None, [], 0)。
    """
    if os.path.splitext(path)[1].lower() in EXCEL_EXTS:
        raise ParseError('支付宝账户明细应为 csv 文件，当前为 Excel：{}'.format(os.path.basename(path)))
    try:
        text = _decode(path)
    except Exception as e:
        raise ParseError('读取失败：{}（{}）'.format(path, e))

    header = None
    rows = []
    skipped_header_rows = 0
    for parts in csv.reader(io.StringIO(text)):
        cells = [clean_text_value(p) for p in parts]
        if not any(cells):
            continue
        if _looks_like_header(cells):
            if header is None:
                header = cells
            else:
                skipped_header_rows += 1
            continue
        if header is None:
            continue
        n = len(header)
        if len(cells) < n:
            cells = cells + [''] * (n - len(cells))
        elif len(cells) > n:
            cells = cells[:n]
        rows.append(cells)
    if header is None:
        return None, [], 0
    return header, rows, skipped_header_rows


def find_bill_files(folder):
    """递归查找文件夹内的支付宝账户明细（按文件名或表头特征识别）。"""
    found = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for fn in filenames:
            if not fn.lower().endswith('.csv'):
                continue
            p = os.path.join(dirpath, fn)
            if '账户明细' in fn:
                found.append(p)
                continue
            try:
                head = _decode(p)[:HEADER_SCAN_BYTES]
            except Exception:
                continue
            if KEY_ORDER in head and KEY_USER in head:
                found.append(p)
    return sorted(set(found))


def _pick(df, name):
    """取列（缺失时给空列），统一转字符串、去不可见字符与首尾空白。"""
    if name in df.columns:
        s = df[name]
    else:
        s = pd.Series([''] * len(df), index=df.index)
    return s.where(s.notna(), '').astype(str).map(clean_text_value)


def _pick_first(df, names):
    """按候选列名顺序取第一个存在的列（都没有时给空列）。"""
    for n in names:
        if n in df.columns:
            return _pick(df, n)
    return pd.Series([''] * len(df), index=df.index)


def _top_values(series, n=8):
    """取值分布 top-n（空值记为「(空)」），供报告展示。"""
    v = series.replace('', '(空)')
    return dict((str(k), int(c)) for k, c in v.value_counts().head(n).items())


def _to_amount(series):
    """金额（元）→ 数值；去掉千分位与货币符号，无法解析的记为 NaN。"""
    s = series.str.replace(',', '', regex=False)
    for ch in ('¥', '￥', ' '):
        s = s.str.replace(ch, '', regex=False)
    return pd.to_numeric(s, errors='coerce')


def _direction(value):
    """「收/支」→ 支出/收入（其它取值返回空串，由 R2 过滤）。"""
    v = str(value).strip()
    if not v:
        return ''
    if v in DIR_MAP:
        return DIR_MAP[v]
    if '支出' in v or '付款' in v:
        return DIR_SPEND
    if '收入' in v or '收款' in v:
        return DIR_RECV
    return ''


def empty_stats():
    return {'files_ok': 0, 'files_failed': [], 'rows_raw': 0, 'rows_kept': 0,
            'dropped_direction': 0, 'dropped_self': 0, 'dropped_amount': 0,
            'dropped_empty': 0, 'dropped_dup': 0, 'dropped_nontransfer': 0,
            'skipped_header_rows': 0, 'direction_values': {}, 'other_values': {},
            'other_self_rows': 0, 'type_values': {}, 'pay_method_values': {},
            'status_values': {}, 'product_values': {}, 'area_values': {}, 'user_values': {},
            'r4_mode': None, 'r4_type_column': '',
            'amount_total': 0.0, 'amount_kept_total': 0.0, 'amount_final_total': 0.0,
            'spend_total': 0.0, 'recv_total': 0.0,
            'columns_kept': [], 'empty_columns': [], 'selected_columns': [],
            'subjects': 0, 'missing_columns': [], 'warnings': [],
            'id_rows': 0, 'no_order_rows': 0, 'pay_time_rows': 0, 'dup_order_rows': 0,
            'failed_status_rows': 0}


def build_frame(files, progress=None, drop_self=True, dedupe=False, r4=R4_OFF):
    """把多个账户明细文件解析并归一到统一格式。

    r4：非转账类处理档位（R4_OFF / R4_TRANSFER / R4_P2P）。
    返回 (DataFrame, stats)。DataFrame 列严格为 core.schema.ALL_COLUMNS。
    """
    stats = empty_stats()
    frames = []
    for i, p in enumerate(files, 1):
        if progress is not None:
            progress.update(i - 1, len(files),
                            '解析账户明细 ({}/{})：{}'.format(i, len(files), os.path.basename(p)))
        try:
            header, rows, skipped = parse_bill_file(p)
        except ParseError as e:
            stats['files_failed'].append('{}（{}）'.format(os.path.basename(p), e))
            continue
        if header is None:
            stats['files_failed'].append('{}（未找到表头，已跳过）'.format(os.path.basename(p)))
            continue
        if not rows:
            stats['files_failed'].append('{}（无数据行，已跳过）'.format(os.path.basename(p)))
            continue
        frames.append(pd.DataFrame(rows, columns=header))
        stats['files_ok'] += 1
        stats['skipped_header_rows'] += skipped
    if progress is not None:
        progress.update(len(files), len(files), '解析完成')

    if not frames:
        return pd.DataFrame(columns=ALL_COLUMNS), stats

    raw = pd.concat(frames, ignore_index=True, sort=False)
    stats['rows_raw'] = len(raw)

    need = [k for k in NEED_KEYS if k not in raw.columns]
    if need:
        stats['missing_columns'] = need
        raise ParseError('账户明细缺少必要列：{}\n请确认这是支付宝「账户明细」导出的 csv 文件。'
                         .format('、'.join(need)))

    user = _pick_first(raw, USER_KEYS).map(_name_from_cell)
    counterparty = _pick_first(raw, COUNTERPARTY_KEYS).map(_name_from_cell)
    product = _pick_first(raw, PRODUCT_KEYS)
    direction_raw = _pick_first(raw, DIR_KEYS)
    direction = direction_raw.map(_direction)
    amount = _to_amount(_pick_first(raw, AMOUNT_KEYS)).round(2)
    type_all = _pick_first(raw, TYPE_KEYS)
    pay_method_all = _pick_first(raw, PAY_METHOD_KEYS)
    status_all = _pick_first(raw, STATUS_KEYS)
    area_all = _pick_first(raw, AREA_KEYS)
    remark_all = _pick_first(raw, REMARK_KEYS)
    order_all = _pick_first(raw, ORDER_KEYS)
    check_id_all = _pick_first(raw, CHECK_ID_KEYS)
    pay_time_all = _pick_first(raw, (KEY_PAY_TIME,))

    # 「说明」= 消费名称；备注与它不同且非空时附在后面（支付宝账单里备注常与消费名称重复）
    extra = remark_all.where((remark_all != '') & (remark_all != product), '')
    memo_text = (product + ('；' + extra).where(extra != '', '')).where(product != '', extra)

    keep = pd.Series([True] * len(raw), index=raw.index)

    mask_dir = direction.isin([DIR_SPEND, DIR_RECV])
    stats['dropped_direction'] = int((~mask_dir).sum())
    if stats['dropped_direction']:
        bad = direction_raw[~mask_dir].replace('', '(空)')
        stats['direction_values'] = dict(
            (str(k), int(v)) for k, v in bad.value_counts().head(8).items())
        stats['other_values'] = _top_values(product[~mask_dir], 8)
        stats['other_self_rows'] = int((user[~mask_dir] == counterparty[~mask_dir]).sum())
    keep &= mask_dir

    mask_amount = amount.notna() & (amount > 0)
    stats['dropped_amount'] = int((keep & ~mask_amount).sum())
    keep &= mask_amount

    mask_empty = (user != '') & (counterparty != '')
    stats['dropped_empty'] = int((keep & ~mask_empty).sum())
    keep &= mask_empty

    if drop_self:
        mask_self = user != counterparty
        stats['dropped_self'] = int((keep & ~mask_self).sum())
        keep &= mask_self

    # R4：只保留转账类（默认关）。支付宝没有「交易用途类型」列，按「消费名称」判定：
    # 含「转账」视为转账类；第三档再排掉「转账收款到余额宝」这类理财划转。
    r4 = r4 or R4_OFF
    stats['r4_mode'] = r4
    if r4 != R4_OFF:
        stats['r4_type_column'] = KEY_PRODUCT
        mask_r4 = product.str.contains(TRANSFER_KEYWORD, na=False)
        if r4 == R4_P2P:
            wealth = product.str.contains('|'.join(WEALTH_KEYWORDS), na=False)
            mask_r4 &= ~wealth
        stats['dropped_nontransfer'] = int((keep & ~mask_r4).sum())
        keep &= mask_r4

    sel = raw[keep].copy()
    sel[COL_USER] = user[keep]
    sel[COL_DIR] = direction[keep]
    sel[COL_CP] = counterparty[keep]
    sel[COL_AMOUNT] = amount[keep]
    sel[COL_TIME] = _pick_first(raw, TIME_KEYS)[keep]
    sel[COL_TYPE] = type_all[keep]
    sel[COL_BIZ_TYPE] = pay_method_all[keep]
    sel[COL_STATUS] = status_all[keep]
    sel[COL_MEMO] = memo_text[keep]
    sel[COL_ORDER] = order_all[keep]
    sel[COL_SOURCE] = SOURCE_NAME
    # 支付宝账单没有「对手方接收金额 / 对手侧银行名称」两个扩展列：留空，
    # 由工具层按「整列为空则自动省略」处理（与微信侧不同来源的产物保持同一契约）。
    for c in ALL_COLUMNS:
        if c not in sel.columns:
            sel[c] = ''
    out = sel[ALL_COLUMNS].reset_index(drop=True)

    if dedupe:
        from core.clean import dedupe_frame
        out, stats['dropped_dup'] = dedupe_frame(out)

    # 关键字段取值分布（按清洗后、去重前的记录统计，便于人工判断口径）
    if (type_all[keep] != '').any():
        stats['type_values'] = _top_values(type_all[keep])
    if (pay_method_all[keep] != '').any():
        stats['pay_method_values'] = _top_values(pay_method_all[keep])
    if (status_all[keep] != '').any():
        stats['status_values'] = _top_values(status_all[keep])
    if (product[keep] != '').any():
        stats['product_values'] = _top_values(product[keep])
    if (area_all[keep] != '').any():
        stats['area_values'] = _top_values(area_all[keep])
    stats['user_values'] = _top_values(user[keep], 10)
    stats['id_rows'] = int((check_id_all[keep] != '').sum())
    stats['pay_time_rows'] = int((pay_time_all[keep] != '').sum())
    stats['no_order_rows'] = int((order_all[keep] == '').sum())
    stats['failed_status_rows'] = int(status_all[keep].str.contains('失败', na=False).sum())

    # 交易号重复：同一笔转账的付款方与收款方各有一条明细，用户方/方向不同，属正常
    dup_no = order_all[order_all != '']
    stats['dup_order_rows'] = int(len(dup_no) - dup_no.nunique())

    stats['rows_kept'] = len(out)
    stats['amount_total'] = round(float(amount.fillna(0).sum()), 2)
    stats['amount_kept_total'] = round(float(amount[keep].fillna(0).sum()), 2)
    stats['amount_final_total'] = round(float(out[COL_AMOUNT].sum()), 2)
    stats['spend_total'] = round(float(out.loc[out[COL_DIR] == DIR_SPEND, COL_AMOUNT].sum()), 2)
    stats['recv_total'] = round(float(out.loc[out[COL_DIR] == DIR_RECV, COL_AMOUNT].sum()), 2)
    stats['subjects'] = len(set(out[COL_USER]) | set(out[COL_CP]))

    users = [u for u in stats['user_values'] if u != '(空)']
    if len(users) > 1:
        stats['warnings'].append(
            '本次账单含 {} 个账号主体：{}；「用户方」按账单「用户信息」逐行取值，'
            '多账号导出时同一人可能对应多个账号 ID。'.format(
                len(users), '、'.join('{}（{} 行）'.format(k, v) for k, v in stats['user_values'].items())))
    if stats['dup_order_rows']:
        stats['warnings'].append(
            '有 {} 行「交易号」与其他行重复：同一笔转账在付款方与收款方明细中各出现一次，'
            '用户方/方向不同，R5 去重不会误删；若两行完全相同则会被 R5 合并。'.format(stats['dup_order_rows']))
    if stats['id_rows']:
        stats['warnings'].append(
            '账单「对应的协查数据」列有 {} 行带身份证号（属敏感信息，本工具不入表，'
            '仅保留原始 csv 供溯源）。'.format(stats['id_rows']))
    if stats['failed_status_rows']:
        stats['warnings'].append(
            '有 {} 行「交易状态」含「失败」（如花呗自动还款失败）：账单会保留未成功的交易流水，'
            '它们不产生实际资金变动，本工具默认保留，需要时可据「交易状态」列自行剔除。'
            .format(stats['failed_status_rows']))
    if stats['files_ok'] > 1:
        stats['warnings'].append('本次共解析 {} 个文件：若其中既有汇总文件又有单次明细，'
                                 '同一笔交易可能重复计入；建议勾选「去除重复记录（R5）」。'
                                 .format(stats['files_ok']))
    return out, stats


def format_report(stats, files, out_tsv=None):
    """把清洗统计整理成文本报告（写入 01_清洗报告_支付宝.txt）。"""
    import time
    lines = []
    lines.append('资金流水工具套件 · 导入与清洗报告（支付宝）')
    lines.append('生成时间：' + time.strftime('%Y-%m-%d %H:%M:%S'))
    lines.append('来源：{}（账户明细 csv）'.format(SOURCE_NAME))
    lines.append('')
    lines.append('一、输入')
    lines.append('  文件数：{}'.format(len(files)))
    lines.append('  成功解析：{}'.format(stats['files_ok']))
    if stats['skipped_header_rows']:
        lines.append('  跳过重复表头行：{} 行（同一文件拼接了多段明细）'.format(stats['skipped_header_rows']))
    if stats['files_failed']:
        lines.append('  跳过/失败：{}'.format(len(stats['files_failed'])))
        for f in stats['files_failed']:
            lines.append('    - ' + f)
    lines.append('')
    lines.append('二、清洗结果（规则 R1/R2/R4/R5/R6/R7）')
    lines.append('  原始数据行：{}'.format(stats['rows_raw']))
    line = '  去除「方向不是支出/收入」：{} 行（R2）'.format(stats['dropped_direction'])
    if stats.get('direction_values'):
        line += '；实际取到的「{}」取值：'.format(KEY_DIR) + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['direction_values'].items())
    lines.append(line)
    lines.append('  去除「金额缺失或非正」：{} 行（R7；实测 0 元行多为「资金授权冻结/解冻」'
                 '与多阶段支付的占位记录，不产生资金变动）'.format(stats['dropped_amount']))
    lines.append('  去除「用户方或客户方为空」：{} 行（R7，'
                 '常见于提现/充值/手续费等无对手方记录）'.format(stats['dropped_empty']))
    lines.append('  去除「自转自（用户方=客户方）」：{} 行（R1）'.format(stats['dropped_self']))
    r4_mode = stats.get('r4_mode')
    if r4_mode:
        if r4_mode == R4_OFF:
            lines.append('  非转账类处理（R4）：未启用（保留全部记录）')
        else:
            lines.append('  去除「非转账类」：{} 行（R4，档位「{}」；判定列：{}）'.format(
                stats['dropped_nontransfer'], R4_LABELS.get(r4_mode, r4_mode),
                stats.get('r4_type_column') or '—'))
    lines.append('  去除「重复记录」：{} 行（R5）'.format(stats['dropped_dup']))
    lines.append('  有效数据行：{}'.format(stats['rows_kept']))
    lines.append('  涉及主体：{} 个'.format(stats['subjects']))
    lines.append('')
    lines.append('三、金额与单位（R6）')
    lines.append('  本账单「金额（元）」已是元，本工具不再换算，仅四舍五入到 2 位小数。')
    lines.append('  原始金额合计（全部行）：{:.2f} 元'.format(stats['amount_total']))
    lines.append('  清洗后（去重前）金额合计：{:.2f} 元'.format(stats['amount_kept_total']))
    lines.append('  入表金额合计：{:.2f} 元（支出 {:.2f} / 收入 {:.2f}）'.format(
        stats['amount_final_total'], stats['spend_total'], stats['recv_total']))
    if abs(stats['amount_kept_total'] - stats['amount_final_total']) > 0.01:
        lines.append('  （入表合计小于清洗后合计，差额来自 R5 去重：{} 行）'.format(stats['dropped_dup']))
    lines.append('')
    lines.append('四、关键字段取值分布（按清洗后、去重前的记录统计，用于判断口径）')
    if stats.get('product_values'):
        lines.append('  「消费名称」：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['product_values'].items()))
    if stats.get('type_values'):
        lines.append('  「类型」：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['type_values'].items()))
    if stats.get('pay_method_values'):
        lines.append('  「支付方式」（渠道）：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['pay_method_values'].items()))
    if stats.get('status_values'):
        lines.append('  「交易状态」：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['status_values'].items()))
    if stats.get('area_values'):
        lines.append('  「交易来源地」：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['area_values'].items()))
    if stats.get('user_values'):
        lines.append('  「用户信息」（账号主体）：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['user_values'].items()))
    lines.append('')
    lines.append('五、被 R2 去掉的「其它」方向记录构成')
    if not stats['dropped_direction']:
        lines.append('  本次没有方向为「其它」的记录。')
    else:
        lines.append('  「其它」记录：{} 行，其中「用户方=客户方」（自己账户间划转）：{} 行'.format(
            stats['dropped_direction'], stats.get('other_self_rows', 0)))
        if stats.get('other_values'):
            lines.append('  业务构成：' + '、'.join(
                '{} {} 行'.format(k, v) for k, v in stats['other_values'].items()))
        lines.append('  说明：支付宝「收/支=其它」实测均为余额宝-自动转入/转出等自账户划转，'
                     '对手方即本人，既非对外支出也非对外收入，故不入表（本工具一律以「收/支」为准）。')
    lines.append('')
    lines.append('六、写入字段与格式说明')
    kept = stats.get('columns_kept') or []
    if kept:
        lines.append('  本次写入列（{} 列）：{}'.format(len(kept), '、'.join(kept)))
    if stats.get('empty_columns'):
        lines.append('  整批为空、已自动省略的扩展列：{}'.format('、'.join(stats['empty_columns'])))
    if stats.get('selected_columns'):
        lines.append('  人工选择的列：{}'.format('、'.join(stats['selected_columns'])))
    lines.append('  前 4 列为固定契约：用户方 | 支出/收入 | 客户方 | 金额（元）')
    lines.append('  扩展列：' + '、'.join(EXT_COLUMNS))
    lines.append('  与微信侧的对应关系：交易创建时间→交易时间；类型→交易类型；'
                 '支付方式→交易业务类型；交易状态→交易状态；消费名称（+备注）→说明；交易号→订单号；'
                 '来源固定为「支付宝」')
    lines.append('  「用户信息 / 交易对方信息」形如 `2088…(许文仙)`，统一取括号内名称；'
                 '无括号的（如「花呗」）保留原值')
    lines.append('  文本字段已清除不可见字符（零宽字符、双向控制符等）与制表符前缀')
    lines.append('  R5 去重判定：订单号（+用户方+方向）优先；订单号缺失时用'
                 '「交易时间+金额+用户方+客户方+方向」组合键，同一秒内同额同对手的'
                 '不同订单号交易会被视为重复（勾选去重时才生效）')
    if out_tsv:
        lines.append('  产物：{}'.format(os.path.basename(out_tsv)))
    lines.append('')
    lines.append('七、说明与局限')
    lines.append('  本报告由程序按固定规则自动生成，仅用于数据整理与后续分析，'
                 '不构成审计结论、鉴定意见或法律意见。')
    for w in stats['warnings']:
        lines.append('  注意：' + w)
    return '\n'.join(lines) + '\n'
