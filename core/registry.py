# -*- coding: utf-8 -*-
"""功能注册表：主界面卡片与「一键跑全程」都读这里。

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
    ToolSpec('import', '导入与清洗',
             '微信原始账单 → 统一格式流水（支付宝入口预留，待样本）',
             'tools.wechat.tool'),
    ToolSpec('merge', '多来源合并',
             '多份统一流水 → 合并、去重、按时间排序',
             'tools.merge.tool'),
    ToolSpec('pivot', '总表与透视表',
             '统一流水 → 流水总表(xlsx/csv) + 数据透视表',
             'tools.pivot.tool'),
    ToolSpec('flow', '资金流向分析',
             '流水总表 → 资金流向图与流水分析报告(HTML)',
             'tools.flow.tool'),
]

BY_ID = dict((t.id, t) for t in TOOLS)

# 「一键跑全程」的步骤顺序（功能 D 需要人工确认输入，故默认不纳入）
PIPELINE = ['import', 'merge', 'pivot']


def get(tid):
    return BY_ID.get(tid)


def load_tool(tid):
    """动态导入功能模块：单个功能导入失败不影响其它卡片。"""
    spec = BY_ID[tid]
    return __import__(spec.module, fromlist=['open_window'])
