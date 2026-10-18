# -*- coding: utf-8 -*-
"""统一日志：终端 + 运行日志.txt 双写。

打包为 `--noconsole` 时终端输出全部丢失，日志文件是排障的唯一线索。
"""
import os
import sys
import time

_state = {'path': None, 'fh': None}
LOG_NAME = '运行日志.txt'
ERROR_NAME = 'error_log.txt'


def init_log(out_dir=None, log_name=LOG_NAME):
    """设置日志文件（追加模式）并写入会话头。返回日志文件路径或 None。"""
    close_log()
    if out_dir:
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            out_dir = None
    if out_dir:
        path = os.path.join(out_dir, log_name)
        try:
            _state['fh'] = open(path, 'a', encoding='utf-8')
            _state['path'] = path
        except Exception:
            _state['fh'] = None
            _state['path'] = None
    log('-' * 56)
    return _state['path']


def log_path():
    return _state['path']


def log(*args):
    """带时间戳输出到终端与日志文件；两者都不可用时静默。"""
    msg = '[{}] {}'.format(time.strftime('%H:%M:%S'), ' '.join(str(a) for a in args))
    if sys.stdout is not None:
        try:
            sys.stdout.write(msg + '\n')
            sys.stdout.flush()
        except Exception:
            pass
    fh = _state['fh']
    if fh:
        try:
            fh.write('{} {}\n'.format(time.strftime('%Y-%m-%d'), msg))
            fh.flush()
        except Exception:
            pass


def close_log():
    fh = _state['fh']
    if fh:
        try:
            fh.close()
        except Exception:
            pass
    _state['fh'] = None


def error_log_path(out_dir):
    return os.path.join(out_dir, ERROR_NAME) if out_dir else None


def write_error_log(out_dir, entries, title='读取/解析失败明细'):
    """把失败明细写成 error_log.txt（UTF-8），返回路径；没有失败项时返回 None。

    entries 元素可以是 {'file': ..., 'error': ...} 字典，也可以是普通字符串。
    读取/解析失败的文件不再静默跳过：明细落盘 + 终端日志 + 完成弹窗提示。
    """
    if not out_dir or not entries:
        return None
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        return None
    path = os.path.join(out_dir, ERROR_NAME)
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write('{}（共 {} 项）\n'.format(title, len(entries)))
            f.write('生成时间：{}\n'.format(time.strftime('%Y-%m-%d %H:%M:%S')))
            f.write('-' * 56 + '\n')
            for i, e in enumerate(entries, 1):
                if isinstance(e, dict):
                    f.write('{}. 文件：{}\n   原因：{}\n\n'.format(
                        i, e.get('file', ''), e.get('error', '')))
                else:
                    f.write('{}. {}\n'.format(i, e))
        return path
    except Exception:
        return None
