# -*- coding: utf-8 -*-
"""输出目录、产物命名与会话接力（session.json）。

约定（方案 §11.1）：
- 每次运行创建 `输出_YYYYMMDD_HHMM/`（重名自动加 _2、_3…）；
- 全部产物与 `运行日志.txt`、`session.json` 都放在该目录内；
- 下游功能通过 session.json 找到上游产物，实现「互不干扰但可接力」。
"""
import json
import os
import shutil
import time

OUTPUT_DIR_PREFIX = '输出_'
SESSION_NAME = 'session.json'
LOG_NAME = '运行日志.txt'

# ---- 产物文件名（方案附录 B）----
P_IMPORT_TSV = '01_统一流水.tsv'
P_IMPORT_REPORT = '01_清洗报告.txt'
# 支付宝分支另存一份，避免与微信产物同名互相覆盖（两者契约一致，可一起合并）
P_IMPORT_TSV_ALIPAY = '01_统一流水_支付宝.tsv'
P_IMPORT_REPORT_ALIPAY = '01_清洗报告_支付宝.txt'
P_MERGE_TSV = '02_合并流水.tsv'
P_MERGE_REPORT = '02_合并说明.txt'
P_PIVOT_XLSX = '03_流水总表.xlsx'
P_PIVOT_CSV = '03_流水总表.csv'
P_PIVOT_PIVOT = '03_数据透视表.xlsx'
P_FLOW_HTML = '04_资金流向图与流水分析报告.html'


def timestamp():
    return time.strftime('%Y%m%d_%H%M')


def new_output_dir(base_dir=None, stamp=None):
    """创建并返回新的输出目录。"""
    base_dir = base_dir or os.getcwd()
    if not os.path.isdir(base_dir):
        os.makedirs(base_dir, exist_ok=True)
    name = OUTPUT_DIR_PREFIX + (stamp or timestamp())
    path = os.path.join(base_dir, name)
    i = 2
    while os.path.exists(path):
        path = os.path.join(base_dir, '{}_{}'.format(name, i))
        i += 1
    os.makedirs(path)
    return path


def product(out_dir, filename):
    return os.path.join(out_dir, filename)


def backup_if_exists(path):
    """产物已存在时先改名为 .bak（避免静默覆盖）。返回备份路径或 None。"""
    if not os.path.exists(path):
        return None
    bak = path + '.bak'
    try:
        if os.path.exists(bak):
            os.remove(bak)
        shutil.move(path, bak)
        return bak
    except Exception:
        return None


# ---- 会话接力 ----
def read_session(out_dir):
    if not out_dir:
        return None
    p = os.path.join(out_dir, SESSION_NAME)
    if not os.path.exists(p):
        return None
    try:
        with open(p, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def write_session(out_dir, data):
    p = os.path.join(out_dir, SESSION_NAME)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return p


def update_session(out_dir, step_id, status, outputs=None, args=None, note=''):
    """记录某一步的执行结果（status: done/failed/skipped）。"""
    from core import __version__
    data = read_session(out_dir) or {}
    data['app'] = data.get('app') or '资金流水工具套件'
    data['version'] = __version__
    data.setdefault('created', time.strftime('%Y-%m-%d %H:%M:%S'))
    data['updated'] = time.strftime('%Y-%m-%d %H:%M:%S')
    steps = data.setdefault('steps', {})
    item = steps.get(step_id) or {}
    item['status'] = status
    item['time'] = time.strftime('%Y-%m-%d %H:%M:%S')
    if outputs is not None:
        item['outputs'] = outputs
    if args is not None:
        item['args'] = args
    if note:
        item['note'] = note
    steps[step_id] = item
    write_session(out_dir, data)
    return data


def step_output(out_dir, step_id, key=None):
    """取某一步的产物路径（接力用）。"""
    data = read_session(out_dir) or {}
    item = (data.get('steps') or {}).get(step_id) or {}
    outs = item.get('outputs') or {}
    if key:
        return outs.get(key)
    for name in sorted(outs.keys()):
        if outs[name]:
            return outs[name]
    return None


def find_recent_output_dir(base_dir):
    """在 base_dir 下找最近修改过的 输出_* 目录（自动接力）。"""
    if not base_dir or not os.path.isdir(base_dir):
        return None
    cands = []
    for n in os.listdir(base_dir):
        p = os.path.join(base_dir, n)
        if n.startswith(OUTPUT_DIR_PREFIX) and os.path.isdir(p):
            cands.append(p)
    if not cands:
        return None
    cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return cands[0]


def default_output_base(input_path=None):
    """默认输出根目录：输入文件所在目录；无输入则当前工作目录。"""
    if input_path:
        d = os.path.dirname(os.path.abspath(input_path))
        if os.path.isdir(d):
            return d
    return os.getcwd()
