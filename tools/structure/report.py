# -*- coding: utf-8 -*-
"""试验性模块的 Markdown 报告（Mermaid 图 + 纯文本树）。"""
import html
import os
import time

from tools.structure.analyze import LEVELS, LEVEL_LABELS, TOL

MAX_GROUPS = 8
MAX_TREE_NODES = 30
MAX_FLOW_EDGES = 40


def _money(value):
    return '{:,.2f}'.format(float(value or 0.0))


def _pct(part, total):
    return '{:.1f}%'.format(float(part) * 100.0 / total) if total else '0.0%'


def _cell(value):
    return str(value).replace('|', '\\|').replace('\r', ' ').replace('\n', ' ')


def _code(value):
    return str(value).replace('\r', ' ').replace('\n', ' ')


def _plain(value):
    return _code(value).replace('`', "'")


def _mermaid(value):
    return html.escape(str(value).replace('\r', ' ').replace('\n', ' '), quote=True)


def _ids(model):
    return dict((node['name'], 'n{}'.format(i)) for i, node in enumerate(model['nodes']))


def _pair_direction(pair, parent, child):
    if parent == pair['a']:
        forward, reverse = pair['ab'], pair['ba']
    else:
        forward, reverse = pair['ba'], pair['ab']
    net = forward - reverse
    if abs(net) > TOL:
        if net > 0:
            return parent, child, abs(net), False, forward, reverse
        return child, parent, abs(net), False, forward, reverse
    return parent, child, min(forward, reverse), True, forward, reverse


def _node_line(model, ids, name):
    node = next(n for n in model['nodes'] if n['name'] == name)
    return '{}["{}"]'.format(ids[name], _mermaid('{}｜{}'.format(node['level'], name)))


def _architecture_mermaid(model, group, ids):
    members = group['members'][:MAX_TREE_NODES]
    included = set(members)
    lines = ['```mermaid', 'flowchart TD']
    for name in members:
        lines.append('    ' + _node_line(model, ids, name))
    for name in members:
        parent = group['parent'].get(name)
        if not parent or parent not in included:
            continue
        pair = model['pairs'].get(tuple(sorted((parent, name))))
        if not pair:
            continue
        source, target, amount, both, forward, reverse = _pair_direction(pair, parent, name)
        if both:
            lines.append('    {} -->|"{} 元"| {}'.format(
                ids[parent], _money(forward), ids[name]))
            lines.append('    {} -->|"{} 元"| {}'.format(
                ids[name], _money(reverse), ids[parent]))
        else:
            lines.append('    {} -->|"净额 {} 元"| {}'.format(
                ids[source], _money(amount), ids[target]))
    for level in LEVELS:
        level_nodes = [ids[n] for n in members if group['levels'][n] == level]
        if level_nodes:
            lines.append('    class {} {};'.format(','.join(level_nodes), level.lower()))
    lines.extend([
        '    classDef p0 fill:#9c2f2f,color:#fff;',
        '    classDef p1 fill:#c77c16,color:#fff;',
        '    classDef p2 fill:#3b6ea5,color:#fff;',
        '    classDef p3 fill:#666,color:#fff;',
        '```',
    ])
    return lines


def _architecture_text(model, group):
    children = dict((name, []) for name in group['members'])
    for name in group['members']:
        parent = group['parent'].get(name)
        if parent in children:
            children[parent].append(name)
    lines = []
    omitted = max(0, len(group['members']) - MAX_TREE_NODES)

    def visit(name, depth):
        if len(lines) >= MAX_TREE_NODES:
            return
        prefix = '  ' * depth
        parent = group['parent'].get(name)
        relation = ''
        if parent:
            pair = model['pairs'].get(tuple(sorted((parent, name))))
            if pair:
                source, target, amount, both, forward, reverse = _pair_direction(pair, parent, name)
                if both:
                    relation = '（{} ─{} 元→ {}；{} ─{} 元→ {}）'.format(
                        _code(parent), _money(forward), _code(name),
                        _code(name), _money(reverse), _code(parent))
                else:
                    relation = '（{} ─净额 {} 元→ {}）'.format(
                        _code(source), _money(amount), _code(target))
        lines.append('{}[{}] {}{}'.format(prefix, group['levels'][name], _code(name), relation))
        for child in children.get(name, []):
            visit(child, depth + 1)

    visit(group['root'], 0)
    if omitted:
        lines.append('……本团体另有 {} 个外围主体未展开'.format(omitted))
    return ['```text'] + lines + ['```']


def _flow_mermaid(model, ids):
    edges = model['net_edges'][:MAX_FLOW_EDGES]
    lines = ['```mermaid', 'flowchart LR']
    used = set()
    for edge in edges:
        used.add(edge['source'])
        used.add(edge['target'])
    for name in sorted(used):
        lines.append('    {}["{}"]'.format(ids[name], _mermaid(name)))
    for edge in edges:
        lines.append('    {} -->|"净额 {} 元"| {}'.format(
            ids[edge['source']], _money(edge['amount']), ids[edge['target']]))
    lines.append('```')
    return lines


def _summary_people(nodes):
    return '；'.join('{}（{} 元）'.format(_plain(n['name']), _money(abs(n['net']))) for n in nodes) or '无'


def render_markdown(model, source_file='', input_info=None):
    stats = model['stats']
    info = input_info or {}
    ids = _ids(model)
    sources = sorted((n for n in model['nodes'] if n['net'] < -TOL),
                     key=lambda n: (n['net'], n['name']))
    sinks = sorted((n for n in model['nodes'] if n['net'] > TOL),
                   key=lambda n: (-n['net'], n['name']))
    hubs = sorted(model['nodes'], key=lambda n: (-n['strength'], n['name']))
    h = []
    h.append('# 资金流水走向大致分析与人物关系预估（试验性）')
    h.append('')
    h.append('> 本报告仅根据流水金额与连接关系自动估算，用于没有人物关系图时的初步摸排。'
             'P0–P3 是图结构层级，不是真实身份、组织职级或违法责任认定。')
    h.append('')
    meta = ['生成时间：' + time.strftime('%Y-%m-%d %H:%M:%S')]
    if source_file:
        meta.append('数据文件：' + os.path.basename(source_file))
    if info:
        meta.append('输入类型：' + str(info.get('kind', '')))
        if info.get('sheet'):
            meta.append('工作表：' + str(info['sheet']))
        meta.append('金额单位：' + str(info.get('unit', '元')))
    meta.append('聚合口径：支出方记录优先，无支出记录时使用收入记录')
    h.append('；'.join(meta) + '。')
    h.append('')

    h.append('## 一、结论摘要')
    h.append('')
    h.append('- 有效流水 **{}** 行，涉及 **{}** 个主体、**{}** 条有向通道，通道金额合计 **{}** 元。'.format(
        stats['rows_valid'], stats['subjects_scope'], stats['channels_scope'], _money(stats['amount_scope'])))
    h.append('- 报告重点展示 **{}** 个核心主体；其相关通道金额占当前分析范围 **{}**，'
             '核心主体之间的通道金额占 **{}**。'.format(
                 stats['subjects_core'], _pct(stats['amount_core_related'], stats['amount_scope']),
                 _pct(stats['core_internal_amount'], stats['amount_scope'])))
    h.append('- 估算得到 **{}** 个资金团体。P0 为各团体内“往来强度 × 关联广度”最高的枢纽候选，'
             'P1/P2/P3 依次为距 P0 一级、二级、三级及更远的主体。'.format(stats['groups']))
    h.append('- 资金来源候选：' + _summary_people(sources[:3]) + '。')
    h.append('- 资金受益端候选：' + _summary_people(sinks[:3]) + '。')
    h.append('- 高流量枢纽：' + '；'.join('{}（流量 {} 元）'.format(
        _plain(n['name']), _money(n['strength'])) for n in hubs[:3]) + '。')
    if stats['subjects_omitted']:
        h.append('- 为避免图形拥挤，另有 **{}** 个低流量主体未进入核心图；他们不是被认定为无关人员。'.format(
            stats['subjects_omitted']))
    h.append('')

    h.append('## 二、数据范围与质量提示')
    h.append('')
    h.append('| 项目 | 数量/金额 |')
    h.append('| --- | ---: |')
    h.append('| 原始记录 | {} 行 |'.format(stats['rows_raw']))
    h.append('| 剔除：主体为空 / 方向无效 / 金额无效 / 自环 | {} / {} / {} / {} 行 |'.format(
        stats['dropped_empty'], stats['dropped_direction'], stats['dropped_amount'], stats['dropped_self']))
    h.append('| 当前金额下限 | {} 元 |'.format(_money(stats['min_amount'])))
    h.append('| 下限前通道 / 金额 | {} 条 / {} 元 |'.format(
        stats['channels_all'], _money(stats['amount_all'])))
    h.append('| 同时有支出与收入记录的同向通道 | {} 条 |'.format(stats['dual_channels']))
    h.append('| 上述通道双方记录差额合计 | {} 元 |'.format(_money(stats['dual_diff_total'])))
    h.append('')
    h.append('同一方向同时存在双方账单时只采用支出方金额，避免把一笔交易重复相加；'
             '双方差额只作数据质量提示，不参与层级判定。')
    h.append('')

    h.append('## 三、核心人员分层表')
    h.append('')
    h.append('| 团体 | 层级 | 姓名 | 资金倾向 | 收入(元) | 支出(元) | 净额(元) | 关联人数 |')
    h.append('| --- | --- | --- | --- | ---: | ---: | ---: | ---: |')
    for node in model['nodes']:
        h.append('| {} | {} | {} | {} | {} | {} | {} | {} |'.format(
            node['group'], node['level'], _cell(node['name']), _cell(node['tendency']),
            _money(node['in_amount']), _money(node['out_amount']), _money(node['net']), node['degree']))
    h.append('')
    h.append('层级含义：' + '；'.join('{}={}'.format(level, LEVEL_LABELS[level]) for level in LEVELS) + '。')
    h.append('净额 = 收入 − 支出；流入/流出倾向使用核心主体在当前筛选范围内的全部通道计算。')
    h.append('')

    h.append('## 四、资金团体与人员架构图')
    h.append('')
    h.append('每个团体单独选 P0，树中的箭头表示两节点之间的净流向；近乎对抵时同时画出双向总额。'
             'Mermaid 不可用时可直接阅读其后的纯文本树。')
    h.append('')
    for group in model['groups'][:MAX_GROUPS]:
        h.append('### {}：P0 = {}（{} 个核心主体）'.format(
            group['id'], _plain(group['root']), len(group['members'])))
        h.append('')
        h.append('- 层级分布：' + '，'.join('{} {} 人'.format(
            level, group['level_counts'][level]) for level in LEVELS if group['level_counts'][level]))
        h.append('- 团体内部双向通道总额：{} 元'.format(_money(group['internal_amount'])))
        h.append('')
        h.extend(_architecture_mermaid(model, group, ids))
        h.append('')
        h.extend(_architecture_text(model, group))
        h.append('')
    if len(model['groups']) > MAX_GROUPS:
        h.append('其余 {} 个较小团体未展开；可缩小输入范围或调整核心主体上限后复核。'.format(
            len(model['groups']) - MAX_GROUPS))
        h.append('')

    h.append('## 五、核心资金走向图')
    h.append('')
    if model['net_edges']:
        h.append('下图按绝对净额从大到小展示前 {} 条通道，箭头由净支出方指向净收入方。'.format(
            min(MAX_FLOW_EDGES, len(model['net_edges']))))
        h.append('')
        h.extend(_flow_mermaid(model, ids))
        h.append('')
        h.append('| 净支出方 | 净收入方 | 净额(元) | 正向总额(元) | 反向总额(元) |')
        h.append('| --- | --- | ---: | ---: | ---: |')
        for edge in model['net_edges'][:20]:
            h.append('| {} | {} | {} | {} | {} |'.format(
                _cell(edge['source']), _cell(edge['target']), _money(edge['amount']),
                _money(edge['forward']), _money(edge['reverse'])))
    else:
        h.append('核心主体之间没有可辨识的单向净额，可能以双向对转为主，请结合下节核对。')
    h.append('')

    h.append('## 六、候选资金脉络')
    h.append('')
    if model['chains']:
        h.append('以下链路从主要净流入主体反向选择最大的净流入边，属于贪心回溯结果，不代表唯一或完整路径。')
        h.append('')
        for chain in model['chains']:
            text = _plain(chain['names'][0])
            for name, amount in zip(chain['names'][1:], chain['amounts']):
                text += ' ─{} 元→ {}'.format(_money(amount), _plain(name))
            h.append('- ' + text)
    else:
        h.append('未形成可回溯的净流向链路。')
    h.append('')

    h.append('## 七、双向对转与回流提示')
    h.append('')
    if model['reciprocal']:
        h.append('| 主体 A | 主体 B | A→B(元) | B→A(元) |')
        h.append('| --- | --- | ---: | ---: |')
        for pair in model['reciprocal'][:15]:
            h.append('| {} | {} | {} | {} |'.format(
                _cell(pair['a']), _cell(pair['b']), _money(pair['ab']), _money(pair['ba'])))
        h.append('')
        h.append('双向流水可能包含正常往来，也可能是对转或回流；净额接近零不等于没有资金活动。')
    else:
        h.append('核心主体之间未发现明显的双向对转通道。')
    h.append('')

    h.append('## 八、使用边界')
    h.append('')
    notes = [
        'P0–P3 只表示自动选出的团体枢纽及其图距离，不能据此认定谁是组织者、受益者或涉案人员。',
        '团体按核心图的连通关系估算；同名主体未做身份合并，不同账号也可能被视作不同主体。',
        '净流向会把双向流水对抵，适合找大致方向，不适合替代逐笔核对；对转、退款和账务冲正需回原始流水验证。',
        '核心图只展示高流量主体，省略的主体会影响路径完整性；报告中的人名、金额和箭头都应结合原始凭证复核。',
        '本报告不构成审计结论、鉴定意见或法律意见。',
    ]
    for note in notes:
        h.append('- ' + note)
    h.append('')
    h.append('---')
    h.append('由 资金流水工具套件 · 试验性模块生成')
    return '\n'.join(h) + '\n'
