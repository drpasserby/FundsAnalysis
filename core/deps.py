# -*- coding: utf-8 -*-
"""依赖自检：确认必需三方库可用；缺依赖时给出可直接照做的处理办法。

背景：本套件运行时依赖 pandas / networkx / openpyxl。若用没装依赖的解释器启动
（例如本机的 Python 3.8 线），主界面能正常打开，但一点功能卡片就会报
`No module named 'pandas'`，容易被误判成程序坏了。启动时先自检一次即可避免。
"""
import importlib
import os
import sys

# （模块名，用途）
REQUIRED = (
    ('pandas', '数据处理与读写 Excel/TSV'),
    ('networkx', '资金流向图分析'),
    ('openpyxl', 'xlsx 读写与样式'),
)
OPTIONAL = (
    ('xlrd', '读取旧版 .xls（只影响 .xls，.xlsx 不受影响）'),
)

# Python 3.8（Win7 线）需使用的兼容版本组合；3.9+ 无此限制
PINS_PY38 = 'pandas==2.0.3 networkx==3.1 openpyxl==3.1.2 xlrd==2.0.1'
MIN_PY = (3, 8)


def _importable(name):
    try:
        importlib.import_module(name)
        return True
    except Exception:
        return False


def missing():
    """返回缺失的必需依赖模块名列表。"""
    return [name for name, _desc in REQUIRED if not _importable(name)]


def missing_optional():
    return [name for name, _desc in OPTIONAL if not _importable(name)]


def ok():
    return not missing()


def python_version():
    return '{}.{}.{}'.format(*sys.version_info[:3])


def interpreter_info():
    return '当前解释器：{}\nPython 版本：{}'.format(sys.executable or '(未知)', python_version())


def is_py38_line():
    return sys.version_info[:2] == (3, 8)


def advice(extra=''):
    """缺依赖时的完整提示文本（弹窗用）。"""
    miss = missing()
    lines = ['缺少运行依赖：{}'.format('、'.join(miss) if miss else '（无）')]
    for name, desc in REQUIRED:
        if name in miss:
            lines.append('  · {} —— {}'.format(name, desc))
    lines.append('')
    lines.append(interpreter_info())
    lines.append('')
    lines.append('处理办法（任选其一）：')
    lines.append('  1) 用已经装好依赖的解释器运行本程序（例如本机 Python 3.13 线）；')
    lines.append('  2) 给当前解释器安装依赖：')
    if is_py38_line():
        lines.append('     "{}" -m pip install {}'.format(sys.executable, PINS_PY38))
        lines.append('     （Python 3.8 请勿安装 pandas 3.x / networkx 3.6，它们不支持 3.8）')
    else:
        lines.append('     "{}" -m pip install -r requirements.txt'.format(sys.executable))
    lines.append('')
    if sys.version_info[:2] < MIN_PY:
        lines.append('注意：Python {} 低于最低要求 {}。'.format(
            python_version(), '.'.join(str(v) for v in MIN_PY)))
    opt = missing_optional()
    if opt:
        lines.append('可选依赖缺失：{}（仅影响 .xls 旧格式读取）'.format('、'.join(opt)))
    if extra:
        lines.append('')
        lines.append(extra)
    return '\n'.join(lines)


def _version(name):
    try:
        mod = importlib.import_module(name)
        return getattr(mod, '__version__', '')
    except Exception:
        return ''


def diagnostic_text():
    """一键自检报告：解释器、工作目录、依赖与版本、结论。用于排障时复制给维护者。"""
    miss = missing()
    opt = missing_optional()
    lines = ['【运行环境自检】', interpreter_info(), '工作目录：{}'.format(os.getcwd()), '',
             '必需依赖：']
    for name, desc in REQUIRED:
        ver = _version(name)
        lines.append('  {} {}{} —— {}'.format('正常' if _importable(name) else '缺失', name,
                                            (' ' + ver) if ver else '', desc))
    lines.append('可选依赖：')
    for name, desc in OPTIONAL:
        ver = _version(name)
        lines.append('  {} {}{} —— {}'.format('正常' if _importable(name) else '缺失', name,
                                            (' ' + ver) if ver else '', desc))
    lines.append('')
    lines.append('结论：{}'.format('依赖齐全，可以正常使用。' if not miss
                              else '缺少 {}，请按上方提示安装后再试。'.format('、'.join(miss))))
    if opt:
        lines.append('提示：可选依赖 {} 缺失不影响 .xlsx/.tsv/.csv 的使用。'.format('、'.join(opt)))
    return '\n'.join(lines)
