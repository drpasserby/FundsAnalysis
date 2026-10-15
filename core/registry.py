# -*- coding: utf-8 -*-
"""功能注册表：主界面卡片读这里。

新增来源（如支付宝）时只改本文件 + 新建 tools/<name>/ 子包，launcher 无需改动。
"""


class ToolSpec(object):
    """一个功能卡片。module 指向实现了 open_window(parent, state) 的模块。"""

    def __init__(self, tid, title, subtitle, module, enabled=True, note=''):
        self.id = tid
        self.title = title
        self.subtitle = subtitle
        self.module = module
        self.enabled = enabled
        self.note = note


TOOLS = [
    ToolSpec('import', '微信流水导入',
             '微信支付账单 → 统一格式流水（01_统一流水.tsv）',
             'tools.wechat.tool'),
    ToolSpec('import_alipay', '支付宝流水导入',
             '支付宝账户明细 csv → 统一格式流水（01_统一流水_支付宝.tsv，来源标「支付宝」）',
             'tools.alipay.tool'),
    ToolSpec('merge', '多来源合并',
             '多份统一流水（微信 / 支付宝 / 历史产物）→ 合并、去重、按时间排序',
             'tools.merge.tool'),
    ToolSpec('pivot', '流水总表与数据透视表生成',
             '统一流水 → 流水总表(xlsx/csv) + 数据透视表',
             'tools.pivot.tool'),
    ToolSpec('flow', '资金流向分析图和报告生成',
             '流水总表 → 资金流向图与流水分析报告(HTML)',
             'tools.flow.tool'),
]

BY_ID = dict((t.id, t) for t in TOOLS)


def get(tid):
    return BY_ID.get(tid)


def load_tool(tid):
    """动态导入功能模块：单个功能导入失败不影响其它卡片。"""
    spec = BY_ID[tid]
    return __import__(spec.module, fromlist=['open_window'])
