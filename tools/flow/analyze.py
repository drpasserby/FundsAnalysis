# -*- coding: utf-8 -*-
"""资金流向分析（功能 D 的计算层）。

由 1.3.2 版 main.py 原样迁移：validate_and_clean / build_graphs /
audit_consistency / build_contract / analyze_flow。
除函数体内必要的等价改写外，行为保持与 1.3.2 完全一致
（验收方式：同一输入产出的 HTML 归一化后逐字节相同）。
"""
import math
from collections import defaultdict

import networkx as nx
import pandas as pd

from core.schema import COL_AMOUNT, COL_CP, COL_DIR, COL_USER, DIR_DESC

# 4 个必需列（按列名认列，不依赖列位；顺序仅用于内部 4 列 DataFrame）
REQUIRED_COLS = [COL_USER, COL_DIR, COL_CP, COL_AMOUNT]

# ================= 聚合口径 =================
# 支出/收入记录不一致时，边金额以谁为准（语义详见 build_graphs 文档字符串）
POLICIES = [
    ('spend', '以支出方为准'),
    ('recv', '以收入方为准'),
    ('max', '取大去重'),
    ('sum', '双向求和'),
]


POLICY_LABELS = dict(POLICIES)


# ================= 数据校验与清洗 =================
def _resolve_columns(raw_df, messages):
    """认列：统一格式**按列名**认（列顺序变了也不受影响）；
    列名不全时回退**按列位**（仅用于 v1.3.2 的历史文件：前 4 列 = 用户方|支出/收入|客户方|金额）。"""
    cols = [str(c) for c in raw_df.columns]
    if all(c in cols for c in REQUIRED_COLS):
        return REQUIRED_COLS
    messages.append('提示：表头与预设不一致（应为 用户方|支出/收入|客户方|金额），'
                    '已按默认列位读取：第 1 列=用户方，第 2 列=支出/收入，第 3 列=客户方，第 4 列=金额')
    return cols[:len(REQUIRED_COLS)]


def validate_and_clean(raw_df):
    """校验数据格式并清洗：返回 (清洗后的DataFrame, 提示信息列表)。
    列数不足时抛 ValueError（由调用方以弹窗呈现，避免打包后错误信息不可见）。"""
    messages = []

    if raw_df is not None:
        if not raw_df.columns.is_unique:
            messages.append('警告：输入里存在重复列名，已按首次出现保留')
            raw_df = raw_df.loc[:, ~raw_df.columns.duplicated()]
        if not raw_df.index.is_unique:
            messages.append('警告：检测到重复索引，已重置')
            raw_df = raw_df.reset_index(drop=True)

    ncol = 0 if raw_df is None else raw_df.shape[1]
    if ncol < len(REQUIRED_COLS):
        raise ValueError(
            '数据列数不足：当前共 {} 列，至少需要 {} 列。\n'
            '请确认数据格式为：{}'.format(ncol, len(REQUIRED_COLS), DIR_DESC))

    use = _resolve_columns(raw_df, messages)

    cleaned = []
    bad_rows = []
    self_loops = 0
    for i, row in raw_df.iterrows():
        row_no = i + 2  # 表头占第 1 行，数据从第 2 行开始
        a, direction, c, amount = row[use[0]], row[use[1]], row[use[2]], row[use[3]]

        if pd.isna(a) or pd.isna(c) or pd.isna(direction):
            bad_rows.append('第 {} 行：用户方/客户方/方向 存在空值'.format(row_no))
            continue

        a = str(a).strip()
        direction = str(direction).strip()
        c = str(c).strip()

        if not a or not c or not direction:
            bad_rows.append('第 {} 行：用户方/客户方/方向 为空白'.format(row_no))
            continue

        if '支出' not in direction and '收入' not in direction:
            bad_rows.append('第 {} 行：方向“{}”不是“支出”或“收入”'.format(row_no, direction))
            continue

        try:
            amount = float(amount)
        except (TypeError, ValueError):
            bad_rows.append('第 {} 行：金额“{}”无法转换为数字'.format(row_no, amount))
            continue

        # 金额必须是有限的非负数（拦截 NaN / ±Inf / 负数，避免污染后续求和）
        if not math.isfinite(amount) or amount < 0:
            bad_rows.append('第 {} 行：金额“{}”不是有效的非负数字'.format(row_no, amount))
            continue

        # 自环（用户方 == 客户方）：此处只计数、仍保留在中间数据里，真正被忽略发生在
        # build_graphs()（`if a == c: continue`）——自环不构成通道，因此不会进图。
        if a == c:
            self_loops += 1

        cleaned.append([a, direction, c, amount])

    if self_loops:
        messages.append('自环记录 {} 条（同一账户的进出：数据保留，构图时忽略、不产生通道）'
                        .format(self_loops))
    if bad_rows:
        messages.append('发现 {} 处问题，已自动跳过：'.format(len(bad_rows)))
        messages.extend(bad_rows[:20])
        if len(bad_rows) > 20:
            messages.append('……共 {} 处，仅显示前 20 处'.format(len(bad_rows)))

    return pd.DataFrame(cleaned, columns=REQUIRED_COLS), messages


# ================= 图构建（一次遍历，按节点对聚合“不重复”的双向流水） =================
def build_graphs(df, policy='spend'):
    """一次遍历，按无序节点对聚合双向流水，并按所选聚合口径取值。
    对每一对 (a, b)（a 为字典序较小端）：
      · spend[(a,b)]：a 记“支出”给 b 的总额（钱 a→b）
      · recv[(a,b)]：b 记“收入”自 a 的总额（钱 a→b，双方视角的同一批交易）
    聚合口径 policy（双方记录不一致时以谁为准）：
      · 'spend'：以支出方为准（默认）；该方向无支出记录时回退用收入记录，避免丢边
      · 'recv' ：以收入方为准；无收入记录时回退用支出记录
      · 'max'  ：取大去重（旧版默认行为）
      · 'sum'  ：双向求和（仅当确认双方是相互独立的记录时使用，可能重复计）
    注意：数据无交易编号时，无法区分“同一笔交易的双方记账差异”与“多笔独立交易”，
    因此聚合口径是研究者的显式选择而非自动猜测；差异明细由 audit_consistency() 输出。
    返回无向图 G，每对节点一条边：
      out=a→b 口径金额、inn=b→a 口径金额，
      sab/rab= a→b 的支出方/收入方合计、sba/rba= b→a 的支出方/收入方合计（供一致性审计）。
    """
    spend = defaultdict(float)  # spend[(付款方, 收款方)] = 付款方记的“支出”总额
    recv = defaultdict(float)   # recv[(付款方, 收款方)] = 收款方记的“收入”总额（钱从付款方流向收款方）

    # 按列名取值（列顺序变化不影响）；列名不全时回退列位
    _cols = REQUIRED_COLS if all(c in df.columns for c in REQUIRED_COLS) else list(df.columns[:4])

    for _, row in df.iterrows():
        a = str(row[_cols[0]]).strip()
        direction = str(row[_cols[1]]).strip()
        c = str(row[_cols[2]]).strip()
        amount = float(row[_cols[3]])

        if a == c:
            continue

        if '支出' in direction:
            spend[(a, c)] += amount
        elif '收入' in direction:
            # “a 收入 c”表示钱从 c 流向 a
            recv[(c, a)] += amount

    G = nx.Graph()
    pairs = {}
    for (a, b) in set(spend) | set(recv):
        pairs[(min(a, b), max(a, b))] = True  # 每对只保留一个规范方向
    for (a, b) in sorted(pairs):
        sa, ra = spend.get((a, b), 0.0), recv.get((a, b), 0.0)  # a→b
        sb, rb = spend.get((b, a), 0.0), recv.get((b, a), 0.0)  # b→a
        if policy == 'spend':
            ab = sa if sa > 0 else ra
            ba = sb if sb > 0 else rb
        elif policy == 'recv':
            ab = ra if ra > 0 else sa
            ba = rb if rb > 0 else sb
        elif policy == 'max':
            ab = max(sa, ra)
            ba = max(sb, rb)
        elif policy == 'sum':
            ab = sa + ra
            ba = sb + rb
        else:
            raise ValueError('未知聚合口径：{!r}（可选 spend/recv/max/sum）'.format(policy))
        G.add_edge(a, b, out=ab, inn=ba, sab=sa, rab=ra, sba=sb, rba=rb)
    return G


def audit_consistency(G, tol=0.01):
    """逐对比较支出方/收入方合计，返回不一致明细列表（按差异降序）。
    每条记录：{a, b, spend, recv, diff, ratio, one_sided}
      · spend/recv：方向 a→b 的支出方/收入方合计（金额四舍五入到分）
      · diff：|spend - recv|（大于 tol 才记录）
      · ratio：diff / max(spend, recv) * 100（%）
      · one_sided：只有单侧记录（spend==0 或 recv==0）
    """
    rows = []
    for u, v in sorted(G.edges()):
        info = G[u][v]
        for a, b, sp, rv in ((u, v, info.get('sab', 0.0), info.get('rab', 0.0)),
                             (v, u, info.get('sba', 0.0), info.get('rba', 0.0))):
            if sp <= 0 and rv <= 0:
                continue
            diff = abs(sp - rv)
            if diff > tol:
                rows.append({
                    'a': a, 'b': b,
                    'spend': round(sp, 2), 'recv': round(rv, 2),
                    'diff': round(diff, 2),
                    'ratio': round(diff / max(sp, rv) * 100, 1) if max(sp, rv) else 0.0,
                    'one_sided': (sp == 0) or (rv == 0),
                })
    rows.sort(key=lambda r: -r['diff'])
    return rows


# ================= 数据契约（供 HTML 渲染使用） =================
# networkx 的 spring_layout 在节点数 >= 500 时改用 scipy 稀疏矩阵（见
# networkx/drawing/layout.py：`if len(G) >= 500: A = nx.to_scipy_sparse_array(...)`），
# 而 scipy 不是本项目的依赖（引入后打包体积会显著增大）。因此超过该规模时
# 改用纯 math 实现的快速近似布局，保证大图也能出图、且不新增依赖。
SPRING_MAX_NODES = 499


def _strength(G):
    """节点流量强度：关联边的 out + inn 之和（用于大图布局时把枢纽放在中心）。"""
    s = {}
    for n in G.nodes():
        total = 0.0
        for _, _, info in G.edges(n, data=True):
            total += float(info.get('out', 0.0)) + float(info.get('inn', 0.0))
        s[n] = total
    return s


def _spiral_layout(G):
    """超大图兜底布局：按流量强度降序做黄金角螺旋（强度大的靠近中心）。

    纯 math 实现（无 scipy / 无稠密矩阵），上万节点也是秒级；
    同一张图结果完全确定，便于复现。
    """
    strength = _strength(G)
    nodes = sorted(G.nodes(), key=lambda n: (-strength[n], str(n)))
    total = len(nodes)
    golden = math.pi * (3.0 - math.sqrt(5.0))
    pos = {}
    for i, node in enumerate(nodes):
        r = math.sqrt((i + 0.5) / total) if total else 0.0
        theta = i * golden
        pos[node] = (r * math.cos(theta), r * math.sin(theta))
    return pos


def layout_graph(G):
    """计算节点坐标，返回 (pos, note)。

    · 节点数 <= SPRING_MAX_NODES：与 1.3.2 完全一致的力导向布局（numpy 路径）；
    · 更大的图：改用 _spiral_layout，避免因缺少 scipy 而崩溃。
    note 为 None 表示走的旧版布局；否则是给用户看的说明文字。
    """
    n = G.number_of_nodes()
    if n == 0:
        return {}, None
    if n <= SPRING_MAX_NODES:
        try:
            return nx.spring_layout(G, k=4, seed=42), None
        except ImportError:
            return _spiral_layout(G), '环境缺少 scipy，已改用快速近似布局'
    return _spiral_layout(G), '主体数 {} 超过 {}，已改用快速近似布局（力导向布局需额外依赖 scipy）'.format(
        n, SPRING_MAX_NODES)


def build_contract(G, tol=0.01):
    """生成 HTML 端使用的数据契约：坐标归一化到 [0,1]，节点大小设上限。
    布局后做重叠消除，配合 HTML 端的半径缩放，任意窗口尺寸下圆圈互不遮挡。
    返回 {'nodes': [...], 'edges': [...]}；每条边 amount=source→target 不重复总额，
    back=target→source 不重复总额。
    """
    all_nodes = sorted(G.nodes())

    pos, layout_note = layout_graph(G)

    degree_dict = {n: G.degree(n) for n in all_nodes}

    # 节点大小：基础值 + 度 * 缩放，设上限防止枢纽节点过大
    base_size = 500
    scale = 300
    max_size = 3000
    size_dict = {n: min(base_size + scale * degree_dict[n], max_size) for n in all_nodes}
    max_s = max(size_dict.values()) if size_dict else 1
    min_s = min(size_dict.values()) if size_dict else 1

    display_s = {}
    for n in all_nodes:
        raw_s = size_dict[n]
        display_s[n] = 15 + (raw_s - min_s) / (max_s - min_s) * 40 if max_s != min_s else 30

    # 坐标归一化到 [0,1]
    xs = [pos[n][0] for n in all_nodes if n in pos]
    ys = [pos[n][1] for n in all_nodes if n in pos]
    if xs and ys:
        minx, maxx = min(xs), max(xs)
        miny, maxy = min(ys), max(ys)
        if maxx - minx == 0:
            maxx += 1
        if maxy - miny == 0:
            maxy += 1
    else:
        minx, maxx, miny, maxy = 0, 1, 0, 1

    pos_norm = {}
    for node in all_nodes:
        if node in pos:
            pos_norm[node] = [float((pos[node][0] - minx) / (maxx - minx)),
                              float((pos[node][1] - miny) / (maxy - miny))]
        else:
            pos_norm[node] = [0.5, 0.5]

    # 密度自适应：圆圈总面积超过画布可用比例时整体缩小，保证任意规模都可不重叠布局
    ref_size = 700.0
    max_frac = 0.35
    total_area = sum(math.pi * (display_s[n] / ref_size) ** 2 for n in all_nodes)
    if total_area > max_frac:
        shrink = (max_frac / total_area) ** 0.5
        for n in all_nodes:
            display_s[n] *= shrink

    # 重叠消除：按参考画布最小边 700px 折算节点半径（与 HTML 端 radiusScale 同源），
    # 迭代把互相遮挡的节点对推开，每轮后整体重新归一化回画布内，直至无重叠。
    # 配合 HTML 端的半径缩放，保证任意窗口/屏幕尺寸下圆圈互不遮挡名字。
    radius_norm = {n: display_s[n] / ref_size for n in all_nodes}
    if len(all_nodes) <= 400:
        def push_pass(margin):
            """一轮推离：距离小于阈值(半径和×margin)的节点对互相推开。返回是否发生移动。"""
            moved = False
            for i in range(len(all_nodes)):
                for j in range(i + 1, len(all_nodes)):
                    ni, nj = all_nodes[i], all_nodes[j]
                    dx = pos_norm[nj][0] - pos_norm[ni][0]
                    dy = pos_norm[nj][1] - pos_norm[ni][1]
                    dist2 = dx * dx + dy * dy
                    min_dist = (radius_norm[ni] + radius_norm[nj]) * margin
                    if dist2 < min_dist * min_dist:
                        dist = dist2 ** 0.5 or 1e-9
                        push = (min_dist - dist) / 2 * 1.2  # 过松弛，加快收敛
                        ux, uy = dx / dist, dy / dist
                        pos_norm[ni][0] -= ux * push
                        pos_norm[ni][1] -= uy * push
                        pos_norm[nj][0] += ux * push
                        pos_norm[nj][1] += uy * push
                        moved = True
            return moved

        # 主循环：推离 + 每轮整体归一化回画布内（目标留有 10% 缓冲，抵消归一化的回挤）
        for _ in range(100):
            if not push_pass(1.02 * 1.10):
                break
            xs = [pos_norm[n][0] for n in all_nodes]
            ys = [pos_norm[n][1] for n in all_nodes]
            cminx, cmaxx = min(xs), max(xs)
            cminy, cmaxy = min(ys), max(ys)
            if cmaxx - cminx == 0:
                cmaxx = cminx + 1
            if cmaxy - cminy == 0:
                cmaxy = cminy + 1
            for n in all_nodes:
                pos_norm[n][0] = 0.02 + 0.96 * (pos_norm[n][0] - cminx) / (cmaxx - cminx)
                pos_norm[n][1] = 0.02 + 0.96 * (pos_norm[n][1] - cminy) / (cmaxy - cminy)
        # 收尾：不再整体归一化，按重叠严重程度优先处理，打破环形相切链
        for _ in range(200):
            pairs = []
            for i in range(len(all_nodes)):
                for j in range(i + 1, len(all_nodes)):
                    ni, nj = all_nodes[i], all_nodes[j]
                    dx = pos_norm[nj][0] - pos_norm[ni][0]
                    dy = pos_norm[nj][1] - pos_norm[ni][1]
                    dist = (dx * dx + dy * dy) ** 0.5
                    min_dist = (radius_norm[ni] + radius_norm[nj]) * 1.02
                    if dist < min_dist:
                        pairs.append((min_dist - dist, ni, nj))
            if not pairs:
                break
            pairs.sort(reverse=True)
            done = 0
            for _, ni, nj in pairs:
                dx = pos_norm[nj][0] - pos_norm[ni][0]
                dy = pos_norm[nj][1] - pos_norm[ni][1]
                dist = (dx * dx + dy * dy) ** 0.5 or 1e-9
                min_dist = (radius_norm[ni] + radius_norm[nj]) * 1.02
                if dist < min_dist:
                    push = (min_dist - dist) / 2 * 1.3
                    ux, uy = dx / dist, dy / dist
                    pos_norm[ni][0] -= ux * push
                    pos_norm[ni][1] -= uy * push
                    pos_norm[nj][0] += ux * push
                    pos_norm[nj][1] += uy * push
                    done += 1
            if done == 0:
                break
        for n in all_nodes:
            pos_norm[n][0] = min(0.98, max(0.02, pos_norm[n][0]))
            pos_norm[n][1] = min(0.98, max(0.02, pos_norm[n][1]))

    nodes_json = []
    for node in all_nodes:
        x, y = pos_norm[node]
        nodes_json.append({'id': node, 'x': x, 'y': y, 'size': display_s[node], 'degree': degree_dict[node]})

    def edge_list(G):
        """每对节点一条边：amount=source→target 不重复总额，back=target→source 不重复总额。
        build_graphs 以规范方向 (min,max) 写入 out=flow(min→max)、inn=flow(max→min)，
        这里统一让 source 为较小端，保证方向语义稳定（G.edges() 迭代顺序与 add_edge 无关）。
        """
        edges = []
        for u, v in G.edges():
            info = G[u][v]
            if u < v:
                source, target = u, v
            else:
                source, target = v, u
            edges.append({
                'source': source,
                'target': target,
                'amount': round(info['out'], 2),
                'back': round(info['inn'], 2),
                'sab': round(info.get('sab', 0.0), 2),
                'rab': round(info.get('rab', 0.0), 2),
                'sba': round(info.get('sba', 0.0), 2),
                'rba': round(info.get('rba', 0.0), 2),
            })
        return edges

    return {
        'nodes': nodes_json,
        'edges': edge_list(G),
        'audit': audit_consistency(G, tol),
        'layout_note': layout_note,
    }


def analyze_flow(edges, top_n=5, tol=0.01):
    """对契约边做初步资金流向分析（仅供研究参考，不构成审计结论）。
    edges：build_contract 产出的边列表（含 amount/back 口径金额）。
    返回 dict：total / nodes / edge_count / net_in / net_out / top_flows /
    concentration / top3_in / beneficiary / chain / cycles。
    """
    nodes = set()
    inflow, outflow = {}, {}
    for e in edges:
        s, t = e['source'], e['target']
        nodes.add(s)
        nodes.add(t)
        inflow.setdefault(s, 0.0)
        inflow.setdefault(t, 0.0)
        outflow.setdefault(s, 0.0)
        outflow.setdefault(t, 0.0)
        outflow[s] += e['amount']
        inflow[t] += e['amount']
        outflow[t] += e['back']
        inflow[s] += e['back']
    nodes = sorted(nodes)
    total = sum(inflow.values())
    sig = max(tol, total * 0.001)  # 显著净额阈值：流转总额的 0.1%

    net = {n: inflow[n] - outflow[n] for n in nodes}
    net_in = sorted([(n, net[n]) for n in nodes if net[n] > sig], key=lambda x: -x[1])
    net_out = sorted([(n, -net[n]) for n in nodes if net[n] < -sig], key=lambda x: -x[1])

    flows = []
    for e in edges:
        if e['amount'] > tol:
            flows.append((e['amount'], e['source'], e['target']))
        if e['back'] > tol:
            flows.append((e['back'], e['target'], e['source']))
    flows.sort(key=lambda x: -x[0])
    top_flows = flows[:top_n]

    # 资金集中度：前 3 大归集主体的流入占全网比例
    conc = 0.0
    top3_in = []
    if total > 0:
        inflow_sorted = sorted(nodes, key=lambda n: -inflow[n])
        top3_in = inflow_sorted[:3]
        conc = sum(inflow[n] for n in top3_in) / total

    # 疑似最终收益方：净流入居前且再流出占比低（资金沉淀率高）
    beneficiary = None
    cands = []
    for n in nodes:
        if inflow[n] > sig:
            retain = 1.0 - outflow[n] / inflow[n] if inflow[n] > 0 else 0.0
            cands.append((n, net[n], inflow[n], outflow[n], retain))
    if cands:
        cands.sort(key=lambda x: (-x[1], -x[4]))
        n, netv, inv, outv, retain = cands[0]
        if netv > sig:
            beneficiary = {'node': n, 'net': netv, 'inflow': inv, 'outflow': outv, 'retain': retain}

    # 资金传导链：最大净流出方 → 最大净流入方 的最短路径
    chain = None
    if net_out and net_in:
        DG = nx.DiGraph()
        for e in edges:
            if e['amount'] > tol:
                DG.add_edge(e['source'], e['target'])
            if e['back'] > tol:
                DG.add_edge(e['target'], e['source'])
        s, tgt = net_out[0][0], net_in[0][0]
        try:
            path = nx.shortest_path(DG, s, tgt)
            if len(path) >= 2:
                chain = path
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            chain = None

    # 资金回流环（有向简单环：仅保留最短的若干个，防性能问题）
    cycles = []
    if len(edges) <= 200:
        DG = nx.DiGraph()
        for e in edges:
            if e['amount'] > tol:
                DG.add_edge(e['source'], e['target'])
            if e['back'] > tol:
                DG.add_edge(e['target'], e['source'])
        try:
            scan = 0
            for c in nx.simple_cycles(DG):
                scan += 1
                if scan > 3000:
                    break
                if len(c) >= 3:
                    cycles.append(c)
                    cycles.sort(key=len)
                    cycles = cycles[:5]
                    if len(cycles) == 5 and len(cycles[-1]) == 3:
                        break
        except Exception:
            cycles = []

    return {
        'total': total,
        'nodes': nodes,
        'edge_count': len(edges),
        'net_in': net_in,
        'net_out': net_out,
        'top_flows': top_flows,
        'concentration': conc,
        'top3_in': top3_in,
        'beneficiary': beneficiary,
        'chain': chain,
        'cycles': cycles,
    }
