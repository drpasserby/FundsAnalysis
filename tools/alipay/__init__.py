# -*- coding: utf-8 -*-
"""功能 A 的支付宝分支：支付宝「账户明细」csv → 统一格式流水。

- `parse.py`：解析 + 清洗（GBK/UTF-8 自动识别、逗号分隔、17 列表头族、姓名提取）
- `tool.py`：功能窗口与 headless `run_import`，产物 `01_统一流水_支付宝.tsv` + `01_清洗报告_支付宝.txt`

产物与微信侧同一契约（前 4 列固定 + 扩展列一致 + 来源列写「支付宝」），
可直接交给「多来源合并」一起处理。
"""

STATUS = '已启用'
NOTE = '支持支付宝「账户明细」导出的 csv（GBK/UTF-8、逗号分隔、17 列表头）。'


def available():
    return True
