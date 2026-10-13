# -*- coding: utf-8 -*-
"""一键跑全程：导入与清洗 → 多来源合并 → 流水总表与透视表。

与逐个功能窗口的差别：全程共用一个输出目录（不会生成多个 输出_ 目录），
三步的产物与状态都记在同一个 session.json 里，便于功能 D 接力。
"""
import os

from core import log as logger
from core import paths
from core.progress import NullProgress


def run_all(files=None, folders=None, out_dir=None, progress=None,
            drop_self=True, dedupe=True, do_merge=True, do_pivot=True):
    """依次执行 A（导入清洗）→ B（合并）→ C（总表与透视表），返回产物清单。"""
    progress = progress or NullProgress(logger.log)
    from tools.merge.tool import run_merge
    from tools.pivot.tool import run_pivot
    from tools.wechat.tool import collect_files, run_import

    targets = collect_files(files, folders)
    if not targets:
        raise ValueError('没有找到微信账单文件。\n请添加包含 TenpayTrades.txt 的文件夹，'
                         '或直接选择账单 txt 文件。')
    if not out_dir:
        out_dir = paths.new_output_dir(paths.default_output_base(targets[0]))
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    logger.init_log(out_dir)
    logger.log('一键跑全程：{} 个输入文件 → {}'.format(len(targets), out_dir))

    results = {}
    progress.set_indeterminate('步骤 1/3：导入与清洗…')
    results['import'] = run_import(files=files, folders=folders, out_dir=out_dir,
                                   progress=progress, drop_self=drop_self, dedupe=False)
    current = results['import']['tsv']
    products = [('① 统一流水', current), ('① 清洗报告', results['import']['report'])]

    if do_merge:
        progress.set_indeterminate('步骤 2/3：多来源合并…')
        results['merge'] = run_merge(inputs=[current], out_dir=out_dir, progress=progress,
                                     dedupe=dedupe)
        current = results['merge']['tsv']
        products.append(('② 合并流水', current))

    if do_pivot:
        progress.set_indeterminate('步骤 3/3：流水总表与透视表…')
        results['pivot'] = run_pivot(current, out_dir=out_dir, progress=progress)
        products.append(('③ 流水总表 (xlsx)', results['pivot']['xlsx']))
        products.append(('③ 流水总表 (csv)', results['pivot']['csv']))
        products.append(('③ 数据透视表', results['pivot']['pivot_xlsx']))
    progress.update(1, 1, '完成')

    imp = results['import']['stats']
    lines = ['微信账单 {} 个文件 → 有效记录 {} 行 / {} 个主体。'.format(
        imp['files_ok'], imp['rows_kept'], imp['subjects'])]
    if 'merge' in results:
        m = results['merge']['stats']
        lines.append('合并去重 {} 行 → 保留 {} 行。'.format(m['dropped_dup'], m['rows_kept']))
    if 'pivot' in results:
        p = results['pivot']['stats']
        lines.append('总表 {} 行；主体 {} 个；通道 {} 条。'.format(p['rows'], p['subjects'], p['channels']))
    lines.append('全部产物在同一目录：{}'.format(out_dir))
    summary = '\n'.join(lines)
    logger.log('一键跑全程完成：{}'.format(out_dir))
    return {'out_dir': out_dir, 'final': current, 'products': products, 'summary': summary,
            'results': results, 'files': targets, 'log': logger.log_path()}
