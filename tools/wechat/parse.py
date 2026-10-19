# -*- coding: utf-8 -*-
"""微信账单解析与清洗（功能 A 的解析层）。

输入：微信/财付通「账单流水」导出的 txt（制表符分隔，带若干行抬头/说明；
      同一文件里可能拼接了多份账单，每份都带自己的表头）。
输出：统一格式 DataFrame（schema.py 契约）+ 清洗报告数据。

清洗规则（方案 §4，P0 实现 R1/R2/R5/R6/R7，P1 实现 R4）：
  R1 自转自（用户方 == 客户方）去除
  R2 借贷类型不是「支出/收入」的记录去除（实测取值：出/入、支出/收入）
  R4 只保留转账类（可选，默认关，三档：不过滤 / 只保留转账类 / 只保留对外转账）
  R5 去重（可选，默认关）：订单号 + 用户方 + 方向，否则 时间+金额+双方+方向
  R6 单位统一：「交易金额(分)」→ 元（附对账）
  R7 空值 / 非正金额去除
"""
import io
import os

import pandas as pd

from core import log as logger
from core.schema import (ALL_COLUMNS, COL_AMOUNT, COL_BIZ_TYPE, COL_CP, COL_CP_BANK, COL_CP_ID, COL_DIR,
                         COL_MEMO, COL_ORDER, COL_RECV_AMOUNT, COL_SOURCE, COL_STATUS, COL_TIME, COL_TYPE,
                         COL_USER, COL_USER_ID, DIR_RECV, DIR_SPEND, EXCEL_EXTS, EXT_COLUMNS, ID_COLUMNS,
                         clean_text_value, decode_text, fen_to_yuan)

TARGET_FILENAME = 'TenpayTrades.txt'
SOURCE_NAME = '微信'

KEY_ACCOUNT = '用户侧账号名称'
KEY_USER_ID = '用户ID'
KEY_CP_ID = '对手方ID'
KEY_LOAN_TYPE = '借贷类型'
KEY_COUNTERPARTY = '对手侧账户名称'
KEY_AMOUNT_FEN = '交易金额(分)'
KEY_TIME = '交易时间'
KEY_TRADE_TYPE = '交易类型'
KEY_STATUS = '当前状态'
KEY_ORDER = '交易单号'
KEY_REMARK = '交易备注'
KEY_PRODUCT = '商品'
KEY_USAGE_TYPE = '交易用途类型'
KEY_BIZ_TYPE = '交易业务类型'
KEY_CP_BANK = '对手侧银行名称'
KEY_RECV_FEN = '对手方接收金额(分)'

NEED_KEYS = (KEY_ACCOUNT, KEY_LOAN_TYPE, KEY_COUNTERPARTY, KEY_AMOUNT_FEN)

# 扩展列的候选列名（不同导出版本的列名不同，按顺序取第一个存在的）。
# 实测：财付通「账单流水」用 交易用途类型/备注1/备注2，微信支付账单明细用 交易类型/交易备注/商品。
TYPE_KEYS = (KEY_TRADE_TYPE, KEY_USAGE_TYPE, KEY_BIZ_TYPE)
MEMO_KEYS = (KEY_REMARK, '备注1')
MEMO2_KEYS = (KEY_PRODUCT, '备注2')
# R4 判定用的「业务类型」列（财付通账单为 交易用途类型：转账/提现/消费/退款/充值/手续费/信用卡还款）
R4_TYPE_KEYS = (KEY_USAGE_TYPE, KEY_TRADE_TYPE)

# 原始账号标识列的候选列名（微信侧账单里叫「用户ID / 对手方ID」，不同导出版本可能带后缀）
USER_ID_KEYS = (KEY_USER_ID, '用户侧账号ID')
CP_ID_KEYS = (KEY_CP_ID, '对手侧账号ID')

# R4 三档：不过滤 / 只保留「转账」/ 在上一档基础上再排除零钱通申购·赎回等理财类
R4_OFF = 'off'
R4_TRANSFER = 'transfer'
R4_P2P = 'p2p'
R4_CHOICES = (R4_OFF, R4_TRANSFER, R4_P2P)
R4_LABELS = {
    R4_OFF: '不过滤（保留全部记录）',
    R4_TRANSFER: '只保留转账类（交易用途类型 = 转账）',
    R4_P2P: '只保留对外转账（再去掉零钱通申购/赎回等理财类）',
}
TRANSFER_TYPES = ('转账',)
WEALTH_KEYWORDS = ('申购', '赎回')

# 借贷类型 → 支出/收入。实测：财付通「账单流水」(TenpayTrades.txt) 用单字「出/入」，
# 微信支付账单明细用「支出/收入」，此处统一映射。
DIR_MAP = {
    '支出': DIR_SPEND, '出': DIR_SPEND, '支': DIR_SPEND, '借': DIR_SPEND,
    '收入': DIR_RECV, '入': DIR_RECV, '收': DIR_RECV, '贷': DIR_RECV,
}

SKIP_PREFIXES = ('---', '注', '说明')
HEADER_SCAN_LINES = 100


class ParseError(ValueError):
    """账单无法解析。"""


def _decode(path):
    """读取并解码账单，返回 (文本, 编码名)。

    编码按 utf-8-sig → gb18030 → utf-16 严格尝试，第一个成功的就用；绝不用
    errors='replace' 静默兜底——那会把 ANSI/GBK 文件当成"看似成功的乱码 UTF-8"
    读进来，导致表头字段名与数据错位，进而在 pandas 分组阶段报 InvalidIndexError。
    """
    with open(path, 'rb') as f:
        raw = f.read()
    return decode_text(raw, path)


def _looks_like_header(parts):
    """一行是否是表头（同一文件里可能拼接多份账单，每份自带表头行）。"""
    s = set(p.strip() for p in parts)
    return KEY_LOAN_TYPE in s and (KEY_ACCOUNT in s or KEY_AMOUNT_FEN in s)


def parse_bill_file(path):
    """解析单个账单文件。

    返回 (header, rows, skipped_header_rows)；无法识别表头时返回 (None, [], 0)。
    """
    if os.path.splitext(path)[1].lower() in EXCEL_EXTS:
        raise ParseError('微信账单应为文本文件（.txt），当前为 Excel：{}'.format(os.path.basename(path)))
    try:
        text, enc = _decode(path)
    except Exception as e:
        raise ParseError('读取失败（{}）：{}'.format(type(e).__name__, e))
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    logger.log('[INFO] 读取 {} (编码: {}, 行数: {})'.format(path, enc, len(lines)))

    # 表头定位健壮化：先去掉 BOM / 零宽字符与首尾空白再匹配，
    # 避免编码转换残留的 \ufeff/\u200b 让表头识别失败。
    scan = [clean_text_value(ln) for ln in lines[:HEADER_SCAN_LINES]]
    header_idx = None
    for i, ln in enumerate(scan):
        if KEY_ACCOUNT in ln and '\t' in ln:
            header_idx = i
            break
    if header_idx is None:
        for i, ln in enumerate(scan):
            if ln.count('\t') >= 2:
                header_idx = i
                break
    if header_idx is None:
        raise ParseError('未找到表头行（编码={}，已扫描前 {} 行）：文件可能不是微信「账单流水」，'
                         '或编码/分隔符不对'.format(enc, HEADER_SCAN_LINES))

    header = [clean_text_value(h) for h in lines[header_idx].split('\t')]
    n = len(header)
    rows = []
    skipped_header_rows = 0
    for ln in lines[header_idx + 1:]:
        s = ln.strip()
        if not s or s.startswith(SKIP_PREFIXES):
            continue
        parts = ln.split('\t')
        if len(parts) < 2:
            continue
        if _looks_like_header(parts):
            skipped_header_rows += 1
            continue
        if len(parts) < n:
            parts = parts + [''] * (n - len(parts))
        elif len(parts) > n:
            parts = parts[:n]
        rows.append(parts)
    return header, rows, skipped_header_rows


def find_bill_files(folder):
    """递归查找文件夹内的账单文件（按文件名或表头特征识别）。"""
    found = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for fn in filenames:
            if not fn.lower().endswith('.txt'):
                continue
            p = os.path.join(dirpath, fn)
            if fn == TARGET_FILENAME:
                found.append(p)
                continue
            try:
                text, _enc = _decode(p)
            except Exception as e:
                logger.log('[WARN] 无法解码，已跳过该文件：{}（{}）'.format(p, e))
                continue
            head = text[:4000]
            if KEY_ACCOUNT in head and '\t' in head:
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


def _direction(value):
    """借贷类型 → 支出/收入/其他（支持「出/入」「支出/收入」「支/收」「借/贷」）。"""
    v = str(value).strip()
    if not v:
        return ''
    if v in DIR_MAP:
        return DIR_MAP[v]
    if DIR_SPEND in v:
        return DIR_SPEND
    if DIR_RECV in v:
        return DIR_RECV
    return ''


SAMPLE_LIMIT = 5


def _sample_text(user, direction, counterparty, yuan, time_all, mask, limit=SAMPLE_LIMIT):
    """把被某条规则删掉的前 limit 行拼成单行描述，写进报告供人工复核。"""
    out = []
    for i in user.index[mask][:limit]:
        amt = yuan[i]
        amt_txt = '（金额空）' if pd.isna(amt) else '{:.2f} 元'.format(float(amt))
        out.append('{} | {} → {} | {} | {}'.format(
            time_all[i] or '(无时间)', user[i] or '(空)', counterparty[i] or '(空)',
            direction[i] or '(空)', amt_txt))
    return out


def empty_stats():
    return {'files_ok': 0, 'files_failed': [], 'rows_raw': 0, 'rows_kept': 0,
            'dropped_direction': 0, 'dropped_self': 0, 'dropped_amount': 0,
            'dropped_empty': 0, 'dropped_dup': 0, 'dropped_nontransfer': 0,
            'skipped_header_rows': 0, 'direction_values': {},
            'usage_values': {}, 'biz_type_values': {},
            'r4_mode': None, 'r4_type_column': '',
            'recv_rows': 0, 'recv_diff_rows': 0, 'recv_diff_values': {},
            'columns_kept': [], 'empty_columns': [], 'selected_columns': [],
            'samples': [], 'errors': [],
            'fen_total': 0.0, 'fen_kept_total': 0.0, 'yuan_kept_total': 0.0, 'yuan_total': 0.0, 'spend_total': 0.0, 'recv_total': 0.0,
            'subjects': 0, 'missing_columns': [], 'warnings': []}


def build_frame(files, progress=None, drop_self=True, dedupe=False, r4=R4_OFF):
    """把多个账单文件解析并归一到统一格式。

    r4：非转账类处理档位（R4_OFF / R4_TRANSFER / R4_P2P）。
    返回 (DataFrame, stats)。DataFrame 列严格为 core.schema.ALL_COLUMNS。
    """
    stats = empty_stats()
    frames = []
    for i, p in enumerate(files, 1):
        if progress is not None:
            progress.update(i - 1, len(files), '解析账单 ({}/{})：{}'.format(i, len(files), os.path.basename(p)))
        try:
            header, rows, skipped = parse_bill_file(p)
        except ParseError as e:
            stats['files_failed'].append('{}（{}）'.format(os.path.basename(p), e))
            stats['errors'].append({'file': p, 'error': '{}'.format(e)})
            logger.log('[ERROR] 解析失败：{} —— {}'.format(p, e))
            continue
        if not rows:
            stats['files_failed'].append('{}（无数据行，已跳过）'.format(os.path.basename(p)))
            stats['errors'].append({'file': p, 'error': '文件里只有表头、没有数据行'})
            logger.log('[WARN] 只有表头、没有数据行，已跳过：{}'.format(p))
            continue
        f = pd.DataFrame(rows, columns=header)
        frames.append(f)
        stats['files_ok'] += 1
        stats['skipped_header_rows'] += skipped
    if progress is not None:
        progress.update(len(files), len(files), '解析完成')

    if not frames:
        return pd.DataFrame(columns=ALL_COLUMNS), stats

    raw = pd.concat(frames, ignore_index=True, sort=False)
    if not raw.columns.is_unique:
        logger.log('[WARN] 合并后出现重复列名，已按首次出现保留（各文件表头可能不一致）')
        raw = raw.loc[:, ~raw.columns.duplicated()]
    if not raw.index.is_unique:
        logger.log('[WARN] 检测到重复索引，已重置')
        raw = raw.reset_index(drop=True)
    logger.log('[INFO] 合并结果 shape={} 列名={}'.format(tuple(raw.shape), list(raw.columns)))
    stats['rows_raw'] = len(raw)

    need = [k for k in NEED_KEYS if k not in raw.columns]
    if need:
        stats['missing_columns'] = need
        raise ParseError('账单缺少必要列：{}\n请确认这是微信「账单流水」导出的 txt 文件。'.format('、'.join(need)))

    user = _pick(raw, KEY_ACCOUNT)
    loan_type = _pick(raw, KEY_LOAN_TYPE)
    direction = loan_type.map(_direction)
    counterparty = _pick(raw, KEY_COUNTERPARTY)
    fen = pd.to_numeric(_pick(raw, KEY_AMOUNT_FEN).str.replace(',', '').str.replace('"', ''),
                        errors='coerce')
    yuan = fen_to_yuan(fen)
    type_all = _pick_first(raw, TYPE_KEYS)
    biz_type_all = _pick(raw, KEY_BIZ_TYPE)
    cp_bank_all = _pick(raw, KEY_CP_BANK)
    memo_all = _pick_first(raw, MEMO_KEYS)
    product_all = _pick_first(raw, MEMO2_KEYS)
    # 「说明」= 备注1；备注2（转账说明/商户名/祝福语）。备注2 里全是不可见字符开头的串，
    # 已在 _pick 里清掉，因此真正的空备注不会污染这一列。
    memo_text = (memo_all + ('；' + product_all).where(product_all != '', '')).where(
        memo_all != '', product_all)
    recv_yuan = fen_to_yuan(pd.to_numeric(
        _pick(raw, KEY_RECV_FEN).str.replace(',', '').str.replace('"', ''), errors='coerce'))
    r4_type_all = _pick_first(raw, R4_TYPE_KEYS)
    time_all = _pick(raw, KEY_TIME)

    keep = pd.Series([True] * len(raw), index=raw.index)
    mask_dir = direction.isin([DIR_SPEND, DIR_RECV])
    stats['dropped_direction'] = int((~mask_dir).sum())
    if stats['dropped_direction']:
        bad = loan_type[~mask_dir].replace('', '(空)')
        stats['direction_values'] = dict(
            (str(k), int(v)) for k, v in bad.value_counts().head(8).items())
        stats['samples'].append(('方向不是支出/收入（R2）', _sample_text(
            user, direction, counterparty, yuan, time_all, ~mask_dir)))
    keep &= mask_dir

    mask_amount = yuan.notna() & (yuan > 0)
    drop_amount = keep & ~mask_amount
    stats['dropped_amount'] = int(drop_amount.sum())
    if stats['dropped_amount']:
        stats['samples'].append(('金额缺失或非正（R7）', _sample_text(
            user, direction, counterparty, yuan, time_all, drop_amount)))
    keep &= mask_amount

    mask_empty = (user != '') & (counterparty != '')
    drop_empty = keep & ~mask_empty
    stats['dropped_empty'] = int(drop_empty.sum())
    if stats['dropped_empty']:
        stats['samples'].append(('用户方或客户方为空（R7）', _sample_text(
            user, direction, counterparty, yuan, time_all, drop_empty)))
    keep &= mask_empty

    if drop_self:
        mask_self = user != counterparty
        drop_self_rows = keep & ~mask_self
        stats['dropped_self'] = int(drop_self_rows.sum())
        if stats['dropped_self']:
            stats['samples'].append(('自转自：用户方 = 客户方（R1）', _sample_text(
                user, direction, counterparty, yuan, time_all, drop_self_rows)))
        keep &= mask_self

    # R4：只保留转账类（默认关）。财付通账单的「交易用途类型」= 转账/提现/消费/退款/充值/手续费/
    # 信用卡还款；其中「转账」里还混着零钱通申购·赎回（理财），故第三档再排掉这些。
    r4 = r4 or R4_OFF
    stats['r4_mode'] = r4
    if r4 != R4_OFF:
        if KEY_USAGE_TYPE in raw.columns:
            stats['r4_type_column'] = KEY_USAGE_TYPE
        elif KEY_TRADE_TYPE in raw.columns:
            stats['r4_type_column'] = KEY_TRADE_TYPE
        else:
            stats['warnings'].append('R4 未生效：账单里没有可识别的「交易用途类型 / 交易类型」列。')
        if stats['r4_type_column']:
            mask_r4 = r4_type_all.isin(TRANSFER_TYPES)
            if r4 == R4_P2P:
                mask_r4 &= ~memo_text.str.contains('|'.join(WEALTH_KEYWORDS), na=False)
            drop_r4 = keep & ~mask_r4
            stats['dropped_nontransfer'] = int(drop_r4.sum())
            if stats['dropped_nontransfer']:
                stats['samples'].append(('非转账类（R4）', _sample_text(
                    user, direction, counterparty, yuan, time_all, drop_r4)))
            keep &= mask_r4

    sel = raw[keep].copy()
    sel[COL_USER] = user[keep]
    sel[COL_DIR] = direction[keep]
    sel[COL_CP] = counterparty[keep]
    sel[COL_AMOUNT] = yuan[keep].round(2)
    # 原始账号标识（强制列：保留原样，便于与支付宝侧「用户信息 / 交易对方信息」对齐、也便于回查账单）
    sel[COL_USER_ID] = _pick_first(raw, USER_ID_KEYS)[keep]
    sel[COL_CP_ID] = _pick_first(raw, CP_ID_KEYS)[keep]
    for _col, _label in ((COL_USER_ID, '用户ID / 用户侧账号ID'), (COL_CP_ID, '对手方ID / 对手侧账号ID')):
        _s = sel[_col]
        if _col not in raw.columns or bool((_s.isna() | (_s.astype(str) == '')).all()):
            stats['warnings'].append(
                '账单里没有可用的「{}」列：「{}」会整列为空（该列是强制保留列，不会自动省略）。'.format(
                    _label, _col))
    sel[COL_TIME] = time_all[keep]
    sel[COL_TYPE] = type_all[keep]
    sel[COL_BIZ_TYPE] = biz_type_all[keep]
    sel[COL_STATUS] = _pick(raw, KEY_STATUS)[keep]
    sel[COL_MEMO] = memo_text[keep]
    sel[COL_ORDER] = _pick(raw, KEY_ORDER)[keep]
    sel[COL_RECV_AMOUNT] = recv_yuan[keep]
    sel[COL_CP_BANK] = cp_bank_all[keep]
    sel[COL_SOURCE] = SOURCE_NAME
    out = sel[ALL_COLUMNS].reset_index(drop=True)

    if dedupe:
        from core.clean import dedupe_frame
        out, stats['dropped_dup'] = dedupe_frame(out)

    # 关键字段取值分布（按清洗后、去重前的记录统计，便于人工判断口径）
    if (r4_type_all[keep] != '').any():
        stats['usage_values'] = _top_values(r4_type_all[keep])
    if (biz_type_all[keep] != '').any():
        stats['biz_type_values'] = _top_values(biz_type_all[keep])

    # 双向记账线索：账单「对手方接收金额(分)」只有「出」方向记录带，可与交易金额对账
    recv_kept = recv_yuan[keep]
    has_recv = recv_kept.notna() & (recv_kept > 0)
    stats['recv_rows'] = int(has_recv.sum())
    diff_recv = has_recv & ((recv_kept - yuan[keep]).abs() > 0.01)
    stats['recv_diff_rows'] = int(diff_recv.sum())
    if stats['recv_diff_rows']:
        combo = (r4_type_all[keep][diff_recv] + '|' + memo_all[keep][diff_recv])
        stats['recv_diff_values'] = _top_values(combo, 5)
        stats['warnings'].append('有 {} 行「交易金额」与账单里的「对手方接收金额」不一致'
                                 '（红包为发放额与实际领取额之差、零钱通为代销渠道记账），'
                                 '本工具一律以「交易金额」入表，构成见清洗报告第五节。'
                                 .format(stats['recv_diff_rows']))

    stats['rows_kept'] = len(out)
    stats['fen_total'] = round(float(fen[keep].fillna(0).sum()), 2)
    # 清洗后（去重前）的换算对账：只有这里 分/100 与元不一致才是真正的单位问题；
    # 去重会再减少行数，因此入表合计与它不同属正常。
    stats['fen_kept_total'] = stats['fen_total']
    stats['yuan_kept_total'] = round(float(yuan[keep].fillna(0).sum()), 2)
    stats['yuan_total'] = round(float(out[COL_AMOUNT].sum()), 2)
    stats['spend_total'] = round(float(out.loc[out[COL_DIR] == DIR_SPEND, COL_AMOUNT].sum()), 2)
    stats['recv_total'] = round(float(out.loc[out[COL_DIR] == DIR_RECV, COL_AMOUNT].sum()), 2)
    subjects = set(out[COL_USER]) | set(out[COL_CP])
    stats['subjects'] = len(subjects)
    if abs(stats['fen_kept_total'] / 100.0 - stats['yuan_kept_total']) > 0.01:
        stats['warnings'].append('单位换算对账不一致：清洗后金额(分)合计 {} → {:.2f} 元，'
                                 '实际 {:.2f} 元'.format(stats['fen_kept_total'],
                                                     stats['fen_kept_total'] / 100.0,
                                                     stats['yuan_kept_total']))
    if stats['files_ok'] > 1:
        stats['warnings'].append('本次共解析 {} 个文件：若其中既有汇总文件又有单次账单，'
                                 '同一笔交易可能重复计入；建议勾选「去除重复记录（R5）」。'
                                 .format(stats['files_ok']))
    return out, stats


def format_report(stats, files, out_tsv=None):
    """把清洗统计整理成文本报告（写入 01_清洗报告.txt）。"""
    import time
    lines = []
    lines.append('资金流水工具套件 · 导入与清洗报告')
    lines.append('生成时间：' + time.strftime('%Y-%m-%d %H:%M:%S'))
    lines.append('来源：{}（原始账单）'.format(SOURCE_NAME))
    lines.append('')
    lines.append('一、输入')
    lines.append('  文件数：{}'.format(len(files)))
    lines.append('  成功解析：{}'.format(stats['files_ok']))
    if stats['skipped_header_rows']:
        lines.append('  跳过重复表头行：{} 行（同一文件拼接了多段账单）'.format(stats['skipped_header_rows']))
    if stats['files_failed']:
        lines.append('  跳过/失败：{}'.format(len(stats['files_failed'])))
        for f in stats['files_failed']:
            lines.append('    - ' + f)
    lines.append('')
    lines.append('二、清洗结果（规则 R1/R2/R4/R5/R6/R7）')
    lines.append('  原始数据行：{}'.format(stats['rows_raw']))
    line = '  去除「方向不是支出/收入」：{} 行（R2）'.format(stats['dropped_direction'])
    if stats.get('direction_values'):
        line += '；实际取到的「{}」取值：'.format(KEY_LOAN_TYPE) + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['direction_values'].items())
    lines.append(line)
    lines.append('  去除「金额缺失或非正」：{} 行（R7）'.format(stats['dropped_amount']))
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
    samples = [s for s in (stats.get('samples') or []) if s[1]]
    if samples:
        lines.append('')
        lines.append('  被删样例（各规则前 {} 条，便于人工复核；不代表全部被删记录）：'
                     .format(SAMPLE_LIMIT))
        for label, rows in samples:
            lines.append('    · {}：'.format(label))
            for r in rows:
                lines.append('        ' + r)
    lines.append('')
    lines.append('三、金额与单位换算对账（R6）')
    lines.append('  原始「交易金额(分)」合计：{:.2f} 分'.format(stats['fen_total']))
    lines.append('  换算为元（÷100）：{:.2f} 元'.format(stats['fen_total'] / 100.0))
    lines.append('  清洗后（去重前）金额合计：{:.2f} 元（与上一行 ÷100 的结果应一致）'.format(
        stats['yuan_kept_total']))
    lines.append('  入表金额合计：{:.2f} 元（支出 {:.2f} / 收入 {:.2f}）'.format(
        stats['yuan_total'], stats['spend_total'], stats['recv_total']))
    if abs(stats['yuan_kept_total'] - stats['yuan_total']) > 0.01:
        lines.append('  （入表合计小于清洗后合计，差额来自 R5 去重：{} 行）'.format(stats['dropped_dup']))
    lines.append('')
    lines.append('四、关键字段取值分布（按清洗后、去重前的记录统计，用于判断口径）')
    if stats.get('usage_values'):
        lines.append('  「交易用途类型 / 交易类型」：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['usage_values'].items()))
    if stats.get('biz_type_values'):
        lines.append('  「交易业务类型」（支付渠道）：' + '、'.join(
            '{} {} 行'.format(k, v) for k, v in stats['biz_type_values'].items()))
    if not stats.get('usage_values') and not stats.get('biz_type_values'):
        lines.append('  账单中没有「交易用途类型 / 交易业务类型」列。')
    lines.append('')
    lines.append('五、对手方接收金额对账（双向记账线索）')
    lines.append('  带「对手方接收金额」的记录：{} 行（账单里只有「出」方向记录带此字段）'
                 .format(stats['recv_rows']))
    if stats['recv_rows']:
        same = stats['recv_rows'] - stats['recv_diff_rows']
        lines.append('  与「交易金额」一致：{} 行；不一致：{} 行'.format(same, stats['recv_diff_rows']))
        if stats.get('recv_diff_values'):
            lines.append('  不一致记录的业务构成：' + '、'.join(
                '{} {} 行'.format(k, v) for k, v in stats['recv_diff_values'].items()))
        lines.append('  说明：不一致主要出现在微信红包（发放额与实际领取额不同）与零钱通支付/转账'
                     '（代销渠道记账口径不同）；本工具一律以「交易金额」入表。')
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
    lines.append('  强制保留列：{}（账号原始标识，固定写入、界面不可取消；'.format('、'.join(ID_COLUMNS))
                 + '即使整批为空也不省略，保证与支付宝流水在同一组列上对齐）')
    lines.append('  扩展列：' + '、'.join(EXT_COLUMNS))
    lines.append('  「借贷类型」取值统一映射：出/支出/支/借 → 支出；入/收入/收/贷 → 收入')
    lines.append('  文本字段已清除不可见字符（零宽字符、双向控制符等）')
    if out_tsv:
        lines.append('  产物：{}'.format(os.path.basename(out_tsv)))
    lines.append('')
    lines.append('七、说明与局限')
    lines.append('  本报告由程序按固定规则自动生成，仅用于数据整理与后续分析，'
                 '不构成审计结论、鉴定意见或法律意见。')
    for w in stats['warnings']:
        lines.append('  注意：' + w)
    return '\n'.join(lines) + '\n'
