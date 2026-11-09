# -*- coding: utf-8 -*-
"""仅凭流水估算核心主体、P0-P3 层级、资金团体与净流向。"""
import math
from collections import deque

import networkx as nx
import numpy as np
import pandas as pd

from core.schema import COL_AMOUNT, COL_CP, COL_DIR, COL_USER, DIR_RECV, DIR_SPEND

TOL = 0.005
LEVELS = ('P0', 'P1', 'P2', 'P3')
LEVEL_LABELS = {
    'P0': '团体枢纽候选',
    'P1': '一级直接关联',
    'P2': '二级扩散关联',
    'P3': '三级及更远外围',
}


def _clean_scope(df):
    work = df[[COL_USER, COL_DIR, COL_CP, COL_AMOUNT]].copy()
    stats = {'rows_raw': int(len(work))}
    for col in (COL_USER, COL_DIR, COL_CP):
        work[col] = work[col].map(lambda v: '' if pd.isna(v) else str(v).strip())
    work[COL_AMOUNT] = pd.to_numeric(work[COL_AMOUNT], errors='coerce')

    missing = work[COL_USER].eq('') | work[COL_CP].eq('')
    stats['dropped_empty'] = int(missing.sum())
    work = work.loc[~missing]
    bad_dir = ~work[COL_DIR].isin((DIR_SPEND, DIR_RECV))
    stats['dropped_direction'] = int(bad_dir.sum())
    work = work.loc[~bad_dir]
    bad_amount = ~np.isfinite(work[COL_AMOUNT]) | work[COL_AMOUNT].le(0)
    stats['dropped_amount'] = int(bad_amount.sum())
    work = work.loc[~bad_amount]
    self_loops = work[COL_USER].eq(work[COL_CP])
    stats['dropped_self'] = int(self_loops.sum())
    work = work.loc[~self_loops]
    if work.empty:
        raise ValueError('没有可用于分析的有效流水：主体、方向或金额均不符合要求。')

    work['payer'] = np.where(work[COL_DIR].eq(DIR_SPEND), work[COL_USER], work[COL_CP])
    work['payee'] = np.where(work[COL_DIR].eq(DIR_SPEND), work[COL_CP], work[COL_USER])
    work['spend_amount'] = np.where(work[COL_DIR].eq(DIR_SPEND), work[COL_AMOUNT], 0.0)
    work['recv_amount'] = np.where(work[COL_DIR].eq(DIR_RECV), work[COL_AMOUNT], 0.0)
    work['spend_count'] = work[COL_DIR].eq(DIR_SPEND).astype(int)
    work['recv_count'] = work[COL_DIR].eq(DIR_RECV).astype(int)
    stats['rows_valid'] = int(len(work))
    return work, stats


def _aggregate(work):
    edges = (work.groupby(['payer', 'payee'], as_index=False, sort=False)
             .agg(spend_amount=('spend_amount', 'sum'), recv_amount=('recv_amount', 'sum'),
                  spend_count=('spend_count', 'sum'), recv_count=('recv_count', 'sum')))
    edges['amount'] = edges['spend_amount'].where(edges['spend_amount'].gt(0),
                                                edges['recv_amount'])
    edges['records'] = edges['spend_count'].where(edges['spend_count'].gt(0),
                                                edges['recv_count'])
    edges['dual'] = edges['spend_amount'].gt(0) & edges['recv_amount'].gt(0)
    edges['dual_diff'] = (edges['spend_amount'] - edges['recv_amount']).abs()
    return edges


def _tendency(node):
    strength = node['strength']
    if strength <= 0:
        return '无明显方向'
    ratio = node['net'] / strength
    if ratio >= 0.20:
        return '净流入（受益端候选）'
    if ratio <= -0.20:
        return '净流出（来源端候选）'
    if node['in_amount'] > 0 and node['out_amount'] > 0:
        return '双向过手（中转候选）'
    return '单向往来'


def _node_metrics(edges):
    nodes = {}
    parties = {}
    for row in edges.itertuples(index=False):
        for name in (row.payer, row.payee):
            if name not in nodes:
                nodes[name] = {'name': name, 'in_amount': 0.0, 'out_amount': 0.0,
                               'strength': 0.0, 'degree': 0}
                parties[name] = set()
        nodes[row.payer]['out_amount'] += row.amount
        nodes[row.payee]['in_amount'] += row.amount
        parties[row.payer].add(row.payee)
        parties[row.payee].add(row.payer)
    for name, node in nodes.items():
        node['degree'] = len(parties[name])
        node['strength'] = node['in_amount'] + node['out_amount']
        node['net'] = node['in_amount'] - node['out_amount']
        node['score'] = node['strength'] * (1.0 + math.log1p(node['degree']))
        node['tendency'] = _tendency(node)
    return nodes


def _pair_map(edges):
    pairs = {}
    for row in edges.itertuples(index=False):
        a, b = sorted((row.payer, row.payee))
        key = (a, b)
        if key not in pairs:
            pairs[key] = {'a': a, 'b': b, 'ab': 0.0, 'ba': 0.0, 'weight': 0.0}
        if row.payer == a:
            pairs[key]['ab'] += row.amount
        else:
            pairs[key]['ba'] += row.amount
        pairs[key]['weight'] += row.amount
    return pairs


def _tree(members, graph, nodes):
    root = sorted(members, key=lambda n: (-nodes[n]['score'], n))[0]
    tree_graph = graph.subgraph(members)
    parent = {root: None}
    level = {root: 'P0'}
    order = [root]
    queue = deque([root])
    while queue:
        current = queue.popleft()
        neighbors = [n for n in tree_graph.neighbors(current) if n not in parent]
        neighbors.sort(key=lambda n: (-tree_graph[current][n]['weight'], n))
        for name in neighbors:
            parent[name] = current
            distance = 0
            cursor = name
            while parent[cursor] is not None:
                distance += 1
                cursor = parent[cursor]
            level[name] = LEVELS[min(distance, 3)]
            order.append(name)
            queue.append(name)
    for name in members:
        if name not in parent:
            parent[name] = None
            level[name] = 'P3'
            order.append(name)
    return root, order, parent, level


def _backtrack_chains(net_edges, nodes, limit=5):
    sinks = sorted((n for n in nodes.values() if n['net'] > TOL),
                   key=lambda n: (-n['net'], n['name']))
    chains = []
    seen = set()
    for sink in sinks:
        path = [sink['name']]
        amounts = []
        visited = {sink['name']}
        current = sink['name']
        for _ in range(6):
            incoming = [e for e in net_edges if e['target'] == current and e['source'] not in visited]
            if not incoming:
                break
            edge = sorted(incoming, key=lambda e: (-e['amount'], e['source']))[0]
            path.append(edge['source'])
            amounts.append(edge['amount'])
            visited.add(edge['source'])
            current = edge['source']
        path.reverse()
        amounts.reverse()
        key = tuple(path)
        if len(path) > 1 and key not in seen:
            seen.add(key)
            chains.append({'names': path, 'amounts': amounts})
        if len(chains) >= limit:
            break
    return chains


def build_model(df, min_amount=0.0, max_subjects=60):
    """返回报告所需的完整估算模型；所有金额单位均为元。"""
    min_amount = float(min_amount or 0.0)
    max_subjects = int(max_subjects)
    if min_amount < 0:
        raise ValueError('最小通道金额不能为负数。')
    if max_subjects < 2:
        raise ValueError('核心主体上限至少为 2。')
    work, stats = _clean_scope(df)
    aggregated = _aggregate(work)
    stats['dual_channels'] = int(aggregated['dual'].sum())
    stats['dual_diff_total'] = float(aggregated.loc[aggregated['dual'], 'dual_diff'].sum())
    stats['channels_all'] = int(len(aggregated))
    stats['amount_all'] = float(aggregated['amount'].sum())
    edges_df = aggregated.loc[aggregated['amount'].ge(min_amount)].copy()
    if edges_df.empty:
        raise ValueError('按当前最小通道金额筛选后没有可分析的通道，请降低金额下限。')
    stats['min_amount'] = min_amount
    stats['channels_scope'] = int(len(edges_df))
    stats['amount_scope'] = float(edges_df['amount'].sum())
    stats['subjects_scope'] = int(len(set(edges_df['payer']) | set(edges_df['payee'])))

    nodes = _node_metrics(edges_df)
    ordered = sorted(nodes.values(), key=lambda n: (-n['strength'], -n['degree'], n['name']))
    core_names = set(n['name'] for n in ordered[:max_subjects])
    stats['subjects_core'] = len(core_names)
    stats['subjects_omitted'] = max(0, len(nodes) - len(core_names))
    stats['amount_core_related'] = float(edges_df.loc[
        edges_df['payer'].isin(core_names) | edges_df['payee'].isin(core_names), 'amount'].sum())
    core_df = edges_df.loc[edges_df['payer'].isin(core_names) & edges_df['payee'].isin(core_names)]
    pairs = _pair_map(core_df)

    graph = nx.Graph()
    graph.add_nodes_from(sorted(core_names))
    for pair in pairs.values():
        graph.add_edge(pair['a'], pair['b'], weight=pair['weight'])
    communities = nx.connected_components(graph)
    groups = []
    node_group = {}
    node_level = {}
    for members in communities:
        root, order, parent, levels = _tree(members, graph, nodes)
        internal = sum(graph[u][v]['weight'] for u, v in graph.subgraph(members).edges())
        groups.append({'members': order, 'root': root, 'parent': parent, 'levels': levels,
                       'internal_amount': internal})
    groups.sort(key=lambda g: (-g['internal_amount'], g['root']))
    for index, group in enumerate(groups, 1):
        group['id'] = 'G{}'.format(index)
        for name in group['members']:
            node_group[name] = group['id']
            node_level[name] = group['levels'][name]
        group['level_counts'] = dict((level, sum(1 for n in group['members']
                                                if group['levels'][n] == level)) for level in LEVELS)
    core_nodes = []
    for name in core_names:
        node = dict(nodes[name])
        node['group'] = node_group.get(name, '')
        node['level'] = node_level.get(name, 'P3')
        core_nodes.append(node)
    level_rank = dict((level, i) for i, level in enumerate(LEVELS))
    core_nodes.sort(key=lambda n: (int(n['group'][1:] or 9999), level_rank[n['level']],
                                  -n['strength'], n['name']))

    net_edges = []
    reciprocal = []
    for pair in pairs.values():
        net = pair['ab'] - pair['ba']
        if abs(net) > TOL:
            source, target = (pair['a'], pair['b']) if net > 0 else (pair['b'], pair['a'])
            net_edges.append({'source': source, 'target': target, 'amount': abs(net),
                              'forward': max(pair['ab'], pair['ba']),
                              'reverse': min(pair['ab'], pair['ba'])})
        if pair['ab'] > TOL and pair['ba'] > TOL:
            reciprocal.append({'a': pair['a'], 'b': pair['b'], 'ab': pair['ab'], 'ba': pair['ba'],
                               'smaller': min(pair['ab'], pair['ba'])})
    net_edges.sort(key=lambda e: (-e['amount'], e['source'], e['target']))
    reciprocal.sort(key=lambda e: (-e['smaller'], e['a'], e['b']))
    stats['groups'] = len(groups)
    stats['core_internal_amount'] = float(sum(p['weight'] for p in pairs.values()))
    stats['reciprocal_pairs'] = len(reciprocal)
    stats['net_channels'] = len(net_edges)
    return {'stats': stats, 'nodes': core_nodes, 'groups': groups, 'pairs': pairs,
            'net_edges': net_edges, 'reciprocal': reciprocal,
            'chains': _backtrack_chains(net_edges, dict((n['name'], n) for n in core_nodes))}
