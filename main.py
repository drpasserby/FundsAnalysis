"""
资金流水走向分析工具
版本：1.2.0
作者：wulvxinchen
"""

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import sys
import json
import math
import pandas as pd
import networkx as nx
from collections import defaultdict


# ================= 选择数据文件 =================
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

# 使用提示：说明可读文件格式与表格内容格式
messagebox.showinfo(
    '使用提示',
    '可读取的文件格式：Excel（.xlsx / .xls）\n'
    '\n'
    '表格内容格式（第 1 行为表头）：\n'
    '用户方 | 支出/收入 | 客户方 | 金额（元）\n'
    '\n'
    '点击“确定”后选择数据文件。')

file_path = filedialog.askopenfilename(
    title='选择数据文件',
    filetypes=[('Excel 文件', '*.xlsx;*.xls'), ('所有文件', '*.*')])

if not file_path:
    sys.exit()  # 用户取消选择，正常退出

df = pd.read_excel(file_path, sheet_name='Sheet1', header=0)

# ================= 数据校验与清洗 =================
def validate_and_clean(raw_df):
    """校验数据格式并清洗：返回 (清洗后的DataFrame, 提示信息列表)。"""
    messages = []

    if raw_df.shape[1] < 4:
        messagebox.showerror(
            '数据格式错误',
            '数据列数不足：当前共 {} 列，至少需要 4 列。\n'
            '请确认数据格式为：用户方 | 支出/收入 | 客户方 | 金额'.format(raw_df.shape[1]))
        sys.exit(1)

    # 表头软校验（仅提示，不阻断）
    headers = [str(h) for h in raw_df.columns[:4]]
    if not ('收入' in headers[1] or '支出' in headers[1]):
        messages.append('提示：第 2 列表头“{}”未包含“支出/收入”'.format(headers[1]))
    if '金额' not in headers[3]:
        messages.append('提示：第 4 列表头“{}”未包含“金额”'.format(headers[3]))

    cleaned = []
    bad_rows = []
    self_loops = 0
    for i, row in raw_df.iterrows():
        row_no = i + 2  # 表头占第 1 行，数据从第 2 行开始
        a, direction, c, amount = row.iloc[0], row.iloc[1], row.iloc[2], row.iloc[3]

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

        if a == c:
            self_loops += 1

        cleaned.append([a, direction, c, amount])

    if self_loops:
        messages.append('自环记录 {} 条（已跳过）'.format(self_loops))
    if bad_rows:
        messages.append('发现 {} 处问题，已自动跳过：'.format(len(bad_rows)))
        messages.extend(bad_rows[:20])
        if len(bad_rows) > 20:
            messages.append('……共 {} 处，仅显示前 20 处'.format(len(bad_rows)))

    return pd.DataFrame(cleaned, columns=raw_df.columns[:4]), messages


df, data_messages = validate_and_clean(df)

if data_messages:
    messagebox.showwarning('数据校验提示', '\n'.join(data_messages))

if df.empty:
    sys.exit('没有可用的有效数据，程序退出。')

# ================= 图构建（一次遍历，按节点对聚合“不重复”的双向流水） =================
def build_graphs(df):
    """一次遍历，按无序节点对聚合“不重复”的双向流水。
    对每一对 (a, b)：
      a→b 流向 = max(a支出b总额, b收入a总额)  —— “a支出b”与“b收入a”是同一笔资金，只算一次；
      b→a 流向 = max(b支出a总额, a收入b总额)  —— 同理。
    同一对的两笔真实往来分别保留，不做 X-Y 相减。
    返回无向图 G，每对节点一条边，属性 out=a→b 不重复总额、inn=b→a 不重复总额
    （(a, b) 为规范序：a 为字典序较小端，每对只出现一次）。
    """
    spend = defaultdict(float)  # spend[(付款方, 收款方)] = 付款方记的“支出”总额
    recv = defaultdict(float)   # recv[(付款方, 收款方)] = 收款方记的“收入”总额（钱从付款方流向收款方）

    for _, row in df.iterrows():
        a = str(row.iloc[0]).strip()
        direction = str(row.iloc[1]).strip()
        c = str(row.iloc[2]).strip()
        amount = float(row.iloc[3])

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
        pairs[(min(a, b), max(a, b))] = True  # 每对只保留一个规范方向，避免双向流水互相覆盖
    for (a, b) in sorted(pairs):
        ab = max(spend.get((a, b), 0.0), recv.get((a, b), 0.0))  # a→b 不重复总额
        ba = max(spend.get((b, a), 0.0), recv.get((b, a), 0.0))  # b→a 不重复总额
        G.add_edge(a, b, out=ab, inn=ba)
    return G


G = build_graphs(df)

# ================= 数据契约（供 HTML 渲染使用） =================
def build_contract(G):
    """生成 HTML 端使用的数据契约：坐标归一化到 [0,1]，节点大小设上限。
    布局后做重叠消除，配合 HTML 端的半径缩放，任意窗口尺寸下圆圈互不遮挡。
    返回 {'nodes': [...], 'edges': [...]}；每条边 amount=source→target 不重复总额，
    back=target→source 不重复总额。
    """
    all_nodes = sorted(G.nodes())

    pos = nx.spring_layout(G, k=4, seed=42) if len(G.nodes()) > 0 else {}

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
                'amount': info['out'],
                'back': info['inn'],
            })
        return edges

    return {
        'nodes': nodes_json,
        'edges': edge_list(G),
    }


contract = build_contract(G)


# ================= HTML 输出（自包含交互式查看器） =================
def generate_html(contract, title='资金流水分析演示图', show_amount=True, hide_other=True):
    """生成自包含的交互式 HTML：数据内嵌、离线可用（JS 保持 ES5 兼容旧浏览器）。
    内置两套界面风格，可在设置面板切换：
      · 经典（原 1.1.7 风格）——顶栏图例、方形按钮、居中设置弹窗、朴素表格；
      · iOS（Apple iOS 设计语言）——系统字体、毛玻璃、圆角卡片、深浅色自适应、
        弹簧动画、安全区适配、底部弹窗设置、iOS 开关；iOS 风格下箭头收于圆圈下方不遮挡名字。
    其余交互：点击高亮、悬浮提示、滚轮缩放、拖拽平移、触屏手势、刷新视图；
    点击某用户后画布下方以表格展示其交易详情（交易类型/客户方/金额），不点击不显示。
    """
    nodes_json = json.dumps(contract['nodes'], ensure_ascii=False).replace('</', '<\\/')
    edges_json = json.dumps(contract['edges'], ensure_ascii=False).replace('</', '<\\/')
    safe_title = (title.replace('&', '&amp;').replace('<', '&lt;')
                       .replace('>', '&gt;').replace('"', '&quot;'))
    amt_checked = ' checked' if show_amount else ''
    hide_checked = ' checked' if hide_other else ''
    show_js = 'true' if show_amount else 'false'
    hide_js = 'true' if hide_other else 'false'

    return '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">
<title>''' + safe_title + '''</title>
<style>
:root {
  color-scheme: light dark;
  --bg:#f2f2f7; --card-solid:#ffffff; --glass-bg:rgba(255,255,255,0.78);
  --text:#1c1c1e; --text-2:#8e8e93; --separator:rgba(60,60,67,0.29); --fill:rgba(120,120,128,0.12);
  --tint:#007aff; --income:#34c759; --expense:#ff3b30;
  --shadow-sm:0 1px 2px rgba(0,0,0,0.04), 0 2px 8px rgba(0,0,0,0.05);
  --shadow-md:0 2px 6px rgba(0,0,0,0.06), 0 14px 36px rgba(0,0,0,0.10);
  --canvas-bg:#ffffff; --edge:#c7c7cc;
  --node-fill:rgba(10,132,255,0.10); --node-border:#007aff; --node-text:#1c1c1e;
  --active-fill:rgba(255,149,0,0.30); --active-border:#ff9500;
  --conn-fill:rgba(52,199,89,0.20); --conn-border:#34c759;
  --dim-fill:rgba(120,120,128,0.14); --dim-border:#aeaeb2;
  --tooltip-bg:rgba(28,28,30,0.88); --tooltip-text:#ffffff;
  --label-bg:rgba(255,255,255,0.92); --watermark:rgba(0,0,0,0.05);
  --spring:cubic-bezier(0.32,0.72,0,1);
}
@media (prefers-color-scheme: dark) {
  :root { --bg:#000000; --card-solid:#1c1c1e; --glass-bg:rgba(30,30,32,0.78); --text:#ffffff; --text-2:#8e8e93;
    --separator:rgba(84,84,88,0.6); --fill:rgba(120,120,128,0.24); --tint:#0a84ff;
    --shadow-sm:0 1px 2px rgba(0,0,0,0.4), 0 2px 8px rgba(0,0,0,0.5);
    --shadow-md:0 2px 6px rgba(0,0,0,0.5), 0 14px 36px rgba(0,0,0,0.6);
    --canvas-bg:#0c0c0e; --edge:#48484a;
    --node-fill:rgba(10,132,255,0.22); --node-border:#0a84ff; --node-text:#ffffff;
    --active-fill:rgba(255,149,0,0.35); --active-border:#ff9f0a;
    --conn-fill:rgba(48,209,88,0.22); --conn-border:#30d158;
    --dim-fill:rgba(120,120,128,0.26); --dim-border:#636366;
    --tooltip-bg:rgba(44,44,46,0.94); --tooltip-text:#ffffff; --label-bg:rgba(28,28,30,0.88); --watermark:rgba(255,255,255,0.07);
  }
}
/* ================= 经典界面（默认样式，对应 1.1.7） ================= */
html, body { margin:0; padding:0; }
body { display:flex; flex-direction:column; align-items:center;
       font-family:'Microsoft YaHei','SimHei',sans-serif; background:#fafafa; }
#page { width:100%; display:flex; flex-direction:column; align-items:center; }
#titleText { margin:8px 0 2px; font-size:20px; color:#333; }
#legendClassic { font-size:13px; color:#555; margin:12px 0 8px; }
#legendIos { display:none; }
#canvasWrap { position:relative; }
#canvas { border:1px solid #ccc; background:#fff; cursor:grab; }
#zoomBtns { position:absolute; right:10px; bottom:10px; display:flex; flex-direction:column; gap:6px; }
#zoomBtns button { width:36px; height:36px; font-size:24px; line-height:1; cursor:pointer;
                   border:1px solid #bbb; border-radius:4px; background:#fff; color:#333;
                   -webkit-user-select:none; user-select:none; padding:0; }
#zoomBtns button:hover { background:#f0f0f0; }
#zoomBtns button:active { background:#e0e0e0; }
#settingsPanel { display:none; position:absolute; top:0; left:0; right:0; bottom:0;
                 background:rgba(0,0,0,0.25); z-index:10; }
#settingsPanel.show { display:block; }
#settingsBox { position:absolute; top:50%; left:50%; transform:translate(-50%,-50%);
               background:#fff; border:1px solid #ccc; border-radius:6px; padding:16px 20px;
               font-size:14px; color:#333; box-shadow:0 4px 16px rgba(0,0,0,0.2); }
#grabber { display:none; }
#settingsTitle { font-size:16px; font-weight:bold; margin-bottom:8px; }
#settingsBox .row { display:block; margin:8px 0; cursor:pointer; }
#settingsBox #titleRow { margin:8px 0; cursor:default; }
#titleRow .rowLabel::after { content:'：'; }
#settingsBox .switch { display:inline-block; vertical-align:middle; }
#settingsBox .switch input { position:static; opacity:1; width:auto; height:auto; margin:0; }
#settingsBox .switch .track, #settingsBox .switch .thumb { display:none; }
#settingsClose { margin-top:12px; padding:4px 16px; cursor:pointer; }
.t-ios { display:none; }
.t-classic { display:inline; }
.seg { display:inline-flex; vertical-align:middle; }
.seg-btn { border:1px solid #bbb; background:#fff; color:#333; padding:2px 12px; cursor:pointer; font-size:13px; }
.seg-btn.active { background:#007aff; color:#fff; border-color:#007aff; }
#detailPanel { display:none; margin:16px auto 40px; width:80%; max-width:1000px;
               min-width:400px; font-size:14px; color:#333; }
#detailTitle { text-align:center; font-size:17px; margin:0 0 10px; }
#detailTable { border-collapse:collapse; width:100%; background:#fff; }
#detailTable th, #detailTable td { border:1px solid #dcdcdc; padding:7px 14px; text-align:center; }
#detailTable th { background:#f0f4f8; color:#333; font-weight:bold; }
#detailTable tbody tr:nth-child(even) { background:#fafafa; }
#detailTable tbody tr:hover { background:#f0f7ff; }
#detailTable td.empty { color:#999; }
/* ================= iOS 界面覆盖（body.mode-ios） ================= */
body.mode-ios { background:var(--bg); color:var(--text); -webkit-font-smoothing:antialiased;
  -webkit-tap-highlight-color:transparent;
  font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","SF Pro Display","Helvetica Neue","PingFang SC","Hiragino Sans GB","Microsoft YaHei","Microsoft YaHei UI",sans-serif;
  min-height:100vh; padding:16px;
  padding:max(16px,env(safe-area-inset-top)) max(16px,env(safe-area-inset-right)) max(28px,env(safe-area-inset-bottom)) max(16px,env(safe-area-inset-left)); }
body.mode-ios #page { animation:pageIn .5s var(--spring) both; }
@keyframes pageIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:none; } }
body.mode-ios #titleText { font-size:32px; font-size:clamp(26px,4.5vw,34px); font-weight:700;
  letter-spacing:-0.022em; line-height:1.15; margin:10px 0 0; color:var(--text); }
body.mode-ios #legendClassic { display:none; }
body.mode-ios #legendIos { display:flex; gap:22px; justify-content:center; align-items:center;
  font-size:13px; color:var(--text-2); margin:10px 0 22px; letter-spacing:0.01em; }
#legendIos .dot { width:8px; height:8px; border-radius:50%; display:inline-block; margin-right:7px; }
body.mode-ios #canvasWrap { display:inline-block; max-width:100%; border-radius:24px; overflow:hidden;
  background:var(--canvas-bg); box-shadow:var(--shadow-md); }
body.mode-ios #canvas { border:none; display:block; background:transparent; touch-action:none; }
body.mode-ios #zoomBtns { right:14px; bottom:14px; gap:12px; }
body.mode-ios #zoomBtns button { width:46px; height:46px; border-radius:50%; border:0.5px solid var(--separator);
  background:var(--glass-bg); -webkit-backdrop-filter:saturate(180%) blur(20px); backdrop-filter:saturate(180%) blur(20px);
  color:var(--tint); font-size:22px; box-shadow:var(--shadow-sm); touch-action:manipulation;
  transition:transform .18s var(--spring), opacity .18s ease, background .18s ease; }
body.mode-ios #zoomBtns button:hover { background:var(--fill); }
body.mode-ios #zoomBtns button:active { transform:scale(0.90); opacity:0.72; }
body.mode-ios #zoomBtns button:focus-visible, body.mode-ios #settingsClose:focus-visible, body.mode-ios #titleInput:focus-visible { outline:2px solid var(--tint); outline-offset:2px; }
body.mode-ios #settingsPanel { display:flex; position:fixed; left:0; top:0; right:0; bottom:0; z-index:50; background:rgba(0,0,0,0.4);
  align-items:flex-end; justify-content:center; visibility:hidden; opacity:0; pointer-events:none;
  transition:opacity .3s ease, visibility .3s ease; }
body.mode-ios #settingsPanel.show { visibility:visible; opacity:1; pointer-events:auto; }
body.mode-ios #settingsBox { position:static; width:100%; max-width:520px; background:var(--glass-bg);
  -webkit-backdrop-filter:saturate(180%) blur(40px); backdrop-filter:saturate(180%) blur(40px);
  border-radius:24px 24px 0 0; border:0; border-top:0.5px solid var(--separator);
  box-shadow:0 -12px 40px rgba(0,0,0,0.25); color:var(--text); font-size:17px;
  padding:8px 24px 28px; padding:8px 24px calc(28px + env(safe-area-inset-bottom));
  transform:translateY(105%); transition:transform .5s var(--spring); }
body.mode-ios #settingsPanel.show #settingsBox { transform:translateY(0); }
body.mode-ios #grabber { display:block; width:36px; height:5px; border-radius:3px; background:var(--fill); margin:8px auto 12px; }
body.mode-ios #settingsTitle { font-size:20px; font-weight:700; text-align:center; letter-spacing:-0.01em; margin:4px 0 6px; color:var(--text); }
body.mode-ios #settingsBox .row { display:flex; align-items:center; justify-content:space-between; min-height:50px; gap:16px; padding:2px; margin:0; font-size:17px; color:var(--text); }
body.mode-ios #settingsBox .row + .row { border-top:0.5px solid var(--separator); }
body.mode-ios #settingsBox .row .rowLabel { order:-1; -webkit-user-select:none; user-select:none; }
body.mode-ios #titleRow .rowLabel::after { content:none; }
body.mode-ios #settingsBox .switch { position:relative; display:inline-block; width:51px; height:31px; flex:none; }
body.mode-ios #settingsBox .switch input { position:absolute; opacity:0; width:100%; height:100%; margin:0; cursor:pointer; z-index:2; }
body.mode-ios #settingsBox .switch .track { display:block; position:absolute; left:0; top:0; right:0; bottom:0; border-radius:16px; background:rgba(120,120,128,0.32); transition:background .25s ease; }
body.mode-ios #settingsBox .switch .thumb { display:block; position:absolute; top:2px; left:2px; width:27px; height:27px; border-radius:50%; background:#fff;
  box-shadow:0 2px 5px rgba(0,0,0,0.3), 0 0 1px rgba(0,0,0,0.1); transition:transform .3s var(--spring); }
body.mode-ios #settingsBox .switch input:checked + .track { background:var(--income); }
body.mode-ios #settingsBox .switch input:checked + .track + .thumb { transform:translateX(20px); }
body.mode-ios #titleInput { flex:1; max-width:260px; min-width:0; border:none; outline:none; background:var(--fill); color:var(--text);
  border-radius:10px; padding:8px 12px; font-size:17px; text-align:right; }
body.mode-ios #titleInput:focus { box-shadow:0 0 0 3px rgba(10,132,255,0.28); }
body.mode-ios #settingsClose { width:100%; margin:20px 0 4px; padding:14px; border:none; border-radius:14px; background:var(--tint); color:#fff;
  font-size:17px; font-weight:600; transition:opacity .18s ease, transform .18s var(--spring); }
body.mode-ios #settingsClose:active { opacity:0.85; transform:scale(0.98); }
body.mode-ios .t-classic { display:none; }
body.mode-ios .t-ios { display:inline; }
body.mode-ios .seg { background:var(--fill); border-radius:9px; padding:2px; display:inline-flex; }
body.mode-ios .seg-btn { border:none; background:transparent; color:var(--text-2); padding:5px 18px; font-size:13px; font-weight:500; border-radius:7px; }
body.mode-ios .seg-btn.active { background:var(--card-solid); color:var(--text); box-shadow:var(--shadow-sm); }
body.mode-ios #detailPanel { margin:24px auto 48px; width:80%; max-width:100%; font-size:15px; color:var(--text); }
body.mode-ios #detailTitle { font-size:24px; font-weight:700; letter-spacing:-0.02em; margin:0 0 16px; color:var(--text); }
body.mode-ios #detailTable { border-collapse:separate; border-spacing:0; background:var(--card-solid);
  border-radius:18px; overflow:hidden; box-shadow:var(--shadow-sm); }
body.mode-ios #detailTable th, body.mode-ios #detailTable td { border:0; border-bottom:0.5px solid var(--separator); }
body.mode-ios #detailTable th { background:transparent; color:var(--text-2); font-size:13px; font-weight:600; letter-spacing:0.02em; padding:14px 16px 8px; }
body.mode-ios #detailTable td { padding:13px 16px; color:var(--text); }
body.mode-ios #detailTable tbody tr:last-child td { border-bottom:none; }
body.mode-ios #detailTable tbody tr:nth-child(even) { background:transparent; }
body.mode-ios #detailTable td.empty { color:var(--text-2); }
@media (prefers-reduced-motion: reduce) { * { animation:none !important; transition-duration:0.01ms !important; } }
</style>
</head>
<body class="mode-ios">
<div id="page">
<h2 id="titleText">''' + safe_title + '''</h2>
<div id="legendClassic"><span style="color:#c0392b;">——净流入</span> <span style="color:#2e8b57;">——净流出</span></div>
<div id="legendIos">
  <span><span class="dot" style="background:#ff3b30;"></span>净流入</span>
  <span><span class="dot" style="background:#34c759;"></span>净流出</span>
</div>
<div id="canvasWrap">
<canvas id="canvas" role="img" aria-label="资金流向图，点击圆圈查看其交易详情"></canvas>
<div id="zoomBtns">
  <button id="refreshBtn" type="button" title="刷新视图" aria-label="刷新视图">↻</button>
  <button id="settingsBtn" type="button" title="设置" aria-label="设置">⚙</button>
  <button id="zoomIn" type="button" title="放大" aria-label="放大">+</button>
  <button id="zoomOut" type="button" title="缩小" aria-label="缩小">−</button>
</div>
<div id="settingsPanel" role="dialog" aria-modal="true" aria-label="设置" aria-hidden="true">
  <div id="settingsBox">
    <div id="grabber"></div>
    <div id="settingsTitle">设置</div>
    <div class="row" id="uiStyleRow" role="radiogroup" aria-label="界面风格">
      <span class="rowLabel">界面风格</span>
      <span class="seg">
        <button type="button" class="seg-btn" data-mode="classic" aria-label="经典">经典</button>
        <button type="button" class="seg-btn active" data-mode="ios" aria-label="iOS">iOS</button>
      </span>
    </div>
    <label class="row"><span class="switch"><input type="checkbox" id="cbAmount"''' + amt_checked + '''><span class="track"></span><span class="thumb"></span></span><span class="rowLabel">显示金额</span></label>
    <label class="row" title="开启后点击某个圆圈，只显示该用户及其直接关联的圆圈和连线"><span class="switch"><input type="checkbox" id="cbHideOther"''' + hide_checked + '''><span class="track"></span><span class="thumb"></span></span><span class="rowLabel">隐藏其他</span></label>
    <div class="row" id="titleRow"><span class="rowLabel">标题</span><input type="text" id="titleInput" value="''' + safe_title + '''" aria-label="标题"></div>
    <button id="settingsClose" type="button" aria-label="完成并关闭"><span class="t-classic">关闭</span><span class="t-ios">完成</span></button>
  </div>
</div>
</div>
<div id="detailPanel">
  <h3 id="detailTitle">用户与其他人员的交易详情</h3>
  <table id="detailTable">
    <thead><tr><th>交易类型</th><th>客户方</th><th>金额</th></tr></thead>
    <tbody id="detailBody"></tbody>
  </table>
</div>
</div>
<script>
var nodes = ''' + nodes_json + ''';
var edges = ''' + edges_json + ''';
var PADDING = 80;
var radiusScale = 1;  // 节点半径随画布最小边缩放，防止不同窗口尺寸下圆圈互相遮挡
var scale = 1, panX = 0, panY = 0;
var activeNode = null, hoverNode = null;
var showAmount = ''' + show_js + ''';
var hideOthers = ''' + hide_js + ''';
/* 单边视图：每对用户一条线（source→target 为 amount，反向为 back），默认黑色 */
/* 点击某用户后，连线按该用户视角的净流向着色：净流出=绿色，净流入=红色 */
var canvas = document.getElementById('canvas');
var ctx = canvas.getContext('2d');
var W, H;
var mode = 'ios';  // 'ios' | 'classic'

var nodeIndex = Object.create(null);
for (var k = 0; k < nodes.length; k++) { nodeIndex[nodes[k].id] = nodes[k]; }

/* 从 CSS 读取 iOS 主题色；经典模式使用固定色板 CLASSIC */
function cssVar(name, fallback) {
    var v = '';
    try { v = getComputedStyle(document.documentElement).getPropertyValue(name); } catch (e) { v = ''; }
    v = String(v).trim();
    return (v === '') ? fallback : v;
}
function readColors() {
    return {
        income: cssVar('--income', '#34c759'),
        expense: cssVar('--expense', '#ff3b30'),
        tint: cssVar('--tint', '#007aff'),
        edge: cssVar('--edge', '#c7c7cc'),
        nodeFill: cssVar('--node-fill', 'rgba(10,132,255,0.10)'),
        nodeBorder: cssVar('--node-border', '#007aff'),
        nodeText: cssVar('--node-text', '#1c1c1e'),
        activeFill: cssVar('--active-fill', 'rgba(255,149,0,0.30)'),
        activeBorder: cssVar('--active-border', '#ff9500'),
        connFill: cssVar('--conn-fill', 'rgba(52,199,89,0.20)'),
        connBorder: cssVar('--conn-border', '#34c759'),
        dimFill: cssVar('--dim-fill', 'rgba(120,120,128,0.14)'),
        dimBorder: cssVar('--dim-border', '#aeaeb2'),
        tooltipBg: cssVar('--tooltip-bg', 'rgba(28,28,30,0.88)'),
        tooltipText: cssVar('--tooltip-text', '#ffffff'),
        labelBg: cssVar('--label-bg', 'rgba(255,255,255,0.92)'),
        watermark: cssVar('--watermark', 'rgba(0,0,0,0.05)'),
        canvasBg: cssVar('--canvas-bg', '#ffffff'),
        bwNormal: 1.4, bwActive: 3, bwConn: 2, bwDim: 0.6
    };
}
var CLASSIC = {
    income: '#2e8b57', expense: '#c0392b', edge: '#888',
    nodeFill: '#dbe9fb', nodeBorder: '#333', nodeText: '#000',
    activeFill: '#f39c12', activeBorder: '#c0392b',
    connFill: '#a9dfbf', connBorder: '#2e8b57',
    dimFill: '#e5e5e5', dimBorder: '#999',
    tooltipBg: 'rgba(255,255,255,0.95)', tooltipText: '#333',
    labelBg: '#fff', watermark: '#ddd', canvasBg: '#fff',
    bwNormal: 0.8, bwActive: 3, bwConn: 2, bwDim: 0.5
};
var C = CLASSIC;

function applyMode(newMode) {
    mode = newMode;
    document.body.className = (mode === 'ios') ? 'mode-ios' : 'mode-classic';
    C = (mode === 'ios') ? readColors() : CLASSIC;
    var segs = document.querySelectorAll('.seg-btn');
    for (var i = 0; i < segs.length; i++) {
        segs[i].className = (segs[i].getAttribute('data-mode') === mode) ? 'seg-btn active' : 'seg-btn';
    }
    var dp = document.getElementById('detailPanel');
    if (mode === 'ios' && W) { dp.style.width = W + 'px'; dp.style.maxWidth = '100%'; }
    else { dp.style.width = ''; dp.style.maxWidth = ''; }
    draw();
}
(function() {
    if (!window.matchMedia) { return; }
    var mq = window.matchMedia('(prefers-color-scheme: dark)');
    var on = function() { if (mode === 'ios') { C = readColors(); draw(); } };
    if (mq.addEventListener) { mq.addEventListener('change', on); }
    else if (mq.addListener) { mq.addListener(on); }
})();
/* 轻触反馈模拟（支持 vibrate 的浏览器/设备） */
function tap() { try { if (navigator.vibrate) { navigator.vibrate(8); } } catch (e) {} }

function resizeCanvas() {
    var vw = window.innerWidth || 1200;
    var vh = window.innerHeight || 800;
    var w = (vw < 760) ? Math.max(280, vw - 40) : Math.round(vw * 0.8);
    var h = (vh < 520) ? Math.max(360, vh - 48) : Math.round(vh * 0.8);
    canvas.width = Math.max(280, w);
    canvas.height = Math.max(360, h);
    W = canvas.width; H = canvas.height;
    radiusScale = (Math.min(W, H) - 2 * PADDING) / 700;
    var dp = document.getElementById('detailPanel');
    if (mode === 'ios') { dp.style.width = W + 'px'; dp.style.maxWidth = '100%'; }
    draw();
}
C = readColors();
resizeCanvas();

function px(n) { return (PADDING + n.x * (W - 2 * PADDING)) * scale + panX; }
function py(n) { return (PADDING + n.y * (H - 2 * PADDING)) * scale + panY; }

function fmt(n) { return Math.round(n).toLocaleString(); }

function nodeTotals(id) {
    var tin = 0, tout = 0;
    for (var i = 0; i < edges.length; i++) {
        var e = edges[i];
        if (e.source === id) { tout += e.amount; tin += e.back; }
        else if (e.target === id) { tout += e.back; tin += e.amount; }
    }
    return { in: tin, out: tout };
}

function draw() {
    ctx.clearRect(0, 0, W, H);
    drawEdges();
    drawWatermark();
    drawNodes();
    drawTooltip();
    renderDetail();
}

function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function roundRect(x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y, x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x, y + h, r);
    ctx.arcTo(x, y + h, x, y, r);
    ctx.arcTo(x, y, x + w, y, r);
    ctx.closePath();
}

function renderDetail() {
    var panel = document.getElementById('detailPanel');
    if (activeNode === null) { panel.style.display = 'none'; return; }
    var rows = [];
    for (var i = 0; i < edges.length; i++) {
        var e = edges[i];
        if (e.source === activeNode) {
            rows.push({ type: '支出', other: e.target, amount: e.amount });
            rows.push({ type: '收入', other: e.target, amount: e.back });
        } else if (e.target === activeNode) {
            rows.push({ type: '支出', other: e.source, amount: e.back });
            rows.push({ type: '收入', other: e.source, amount: e.amount });
        }
    }
    rows.sort(function(a, b) { return b.amount - a.amount; });
    document.getElementById('detailTitle').textContent = '用户' + activeNode + '与其他人员的交易详情';
    var html = '';
    for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        var color = (r.type === '收入') ? C.income : C.expense;
        html += '<tr><td style="color:' + color + ';">' + esc(r.type) + '</td><td>' + esc(r.other) + '</td><td>' + fmt(r.amount) + '</td></tr>';
    }
    if (rows.length === 0) {
        html += '<tr><td colspan="3" class="empty">暂无交易记录</td></tr>';
    }
    document.getElementById('detailBody').innerHTML = html;
    panel.style.display = 'block';
}

function drawWatermark() {
    if (mode === 'ios') {
        ctx.font = '1.4em -apple-system, "Helvetica Neue", "PingFang SC", "Microsoft YaHei", sans-serif';
    } else {
        ctx.font = '1.5em Microsoft YaHei, SimHei';
    }
    ctx.fillStyle = C.watermark;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'bottom';
    if (mode === 'ios') { ctx.fillText('资金流水走向分析工具 Github@drpasserby(WLXC)', 18, H - 14); }
    else { ctx.fillText('资金流水走向分析工具 Github@drpasserby(WLXC)', 16, H - 12); }
}

function drawArrow(x, y, angle, color, alpha) {
    var len = 12, w = Math.PI / 7;
    ctx.save();
    ctx.globalAlpha = alpha;
    ctx.fillStyle = color;
    ctx.translate(x, y);
    ctx.rotate(angle);
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(-len, -len * Math.tan(w));
    ctx.lineTo(-len, len * Math.tan(w));
    ctx.closePath();
    ctx.fill();
    ctx.restore();
}

function drawEdges() {
    var connected = null;
    if (activeNode !== null) { connected = Object.create(null); connected[activeNode] = true; }
    for (var i = 0; i < edges.length; i++) {
        var e = edges[i];
        if (hideOthers && activeNode !== null && e.source !== activeNode && e.target !== activeNode) { continue; }
        var sn = nodeIndex[e.source];
        var tn = nodeIndex[e.target];
        if (!sn || !tn) { continue; }
        var sx = px(sn), sy = py(sn);
        var tx = px(tn), ty = py(tn);
        var dx = tx - sx, dy = ty - sy;
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) { continue; }

        // 是否以活动节点为端点：是则可按活动节点视角计算净流向
        var isActiveEnd = (activeNode !== null) && (e.source === activeNode || e.target === activeNode);
        var isSrc = (e.source === activeNode);   // 活动节点在边的 source 端
        var out, in_;
        if (isActiveEnd) {
            out = isSrc ? e.amount : e.back;     // 活动节点流向对方
            in_ = isSrc ? e.back : e.amount;     // 对方流向活动节点
        }

        // 颜色：默认黑色（iOS 用主题中性色保证深浅色下可见）；
        // 点击后按活动节点视角——净流出=绿色、净流入=红色、相等仍为中性
        var lineColor = (mode === 'classic') ? '#000000' : C.edge;
        if (isActiveEnd) {
            if (out > in_) { lineColor = C.income; }
            else if (in_ > out) { lineColor = C.expense; }
        }

        var alpha = 0.7, textAlpha = 1.0;
        if (activeNode !== null) {
            if (isActiveEnd) { alpha = 1.0; }
            else if (connected[e.source] && connected[e.target]) { alpha = 0.9; }
            else { alpha = 0.15; textAlpha = 0.3; }
        }

        ctx.beginPath();
        ctx.moveTo(sx, sy);
        ctx.lineTo(tx, ty);
        ctx.strokeStyle = lineColor;
        ctx.lineWidth = 1.6;
        ctx.globalAlpha = alpha;
        ctx.stroke();
        ctx.globalAlpha = 1;

        // 箭头：仅当活动节点为端点且存在净流向。净流出箭头指向对方，净流入箭头指向活动节点。
        // 尖端收在对应圆圈边缘（绘制于圆圈下方，不遮挡名字）。
        if (isActiveEnd && out !== in_) {
            var aX = isSrc ? sx : tx;   // 活动节点位置
            var aY = isSrc ? sy : ty;
            var oX = isSrc ? tx : sx;   // 对方位置
            var oY = isSrc ? ty : sy;
            var odx = oX - aX, ody = oY - aY;
            var oLen = Math.sqrt(odx * odx + ody * ody) || 1;
            var rAct = (isSrc ? sn : tn).size * scale * radiusScale;
            var rOth = (isSrc ? tn : sn).size * scale * radiusScale;
            var ax2, ay2, aAng;
            if (out > in_) {
                // 净流出：活动节点→对方，箭头尖端收在对方圆边
                ax2 = oX - odx / oLen * rOth;
                ay2 = oY - ody / oLen * rOth;
                aAng = Math.atan2(ody, odx);
            } else {
                // 净流入：对方→活动节点，箭头尖端收在活动节点圆边
                ax2 = aX + odx / oLen * rAct;
                ay2 = aY + ody / oLen * rAct;
                aAng = Math.atan2(-ody, -odx);
            }
            drawArrow(ax2, ay2, aAng, lineColor, 1.0);
        }

        if (showAmount) {
            var label;
            if (isActiveEnd) {
                // 从活动节点视角显示“不重复”的支出与收入
                label = '支出 ' + fmt(out) + '  收入 ' + fmt(in_);
            } else {
                label = e.source + '→' + e.target + ':' + fmt(e.amount) + '  ' + e.target + '→' + e.source + ':' + fmt(e.back);
            }
            ctx.font = (mode === 'ios')
                ? '11px -apple-system, "Helvetica Neue", "PingFang SC", "Microsoft YaHei", sans-serif'
                : '11px Microsoft YaHei, SimHei';
            var tw = ctx.measureText(label).width;
            var lx = (sx + tx) / 2;
            var ly = (sy + ty) / 2;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.globalAlpha = textAlpha;
            if (mode === 'ios') {
                ctx.fillStyle = C.labelBg;
                ctx.fillRect(lx - tw / 2 - 5, ly - 10, tw + 10, 20);
                ctx.fillStyle = lineColor;
            } else {
                ctx.fillStyle = '#fff';
                ctx.fillRect(lx - tw / 2 - 4, ly - 9, tw + 8, 18);
                ctx.fillStyle = '#333';
            }
            ctx.fillText(label, lx, ly);
            ctx.globalAlpha = 1;
        }
    }
}

function buildConnected() {
    var conn = Object.create(null);
    if (activeNode === null) { return conn; }
    for (var i = 0; i < edges.length; i++) {
        var e = edges[i];
        if (e.source === activeNode) { conn[e.target] = true; }
        if (e.target === activeNode) { conn[e.source] = true; }
    }
    conn[activeNode] = true;
    return conn;
}

function drawNodes() {
    var connected = null;
    if (activeNode !== null) { connected = buildConnected(); }
    // 第一轮：画所有圆圈
    for (var j = 0; j < nodes.length; j++) {
        var n = nodes[j];
        if (hideOthers && connected !== null && !connected[n.id]) { continue; }
        var cx = px(n), cy = py(n);
        var r = n.size * scale * radiusScale;
        var fill = C.nodeFill, border = C.nodeBorder, bw = C.bwNormal;
        if (activeNode !== null) {
            if (n.id === activeNode) { fill = C.activeFill; border = C.activeBorder; bw = C.bwActive; }
            else if (connected[n.id]) { fill = C.connFill; border = C.connBorder; bw = C.bwConn; }
            else { fill = C.dimFill; border = C.dimBorder; bw = C.bwDim; }
        }
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, 2 * Math.PI);
        // 先垫一层不透明底色，保证连线永远在圆圈下方（半透明填充时也看不穿）
        ctx.fillStyle = C.canvasBg;
        ctx.fill();
        ctx.fillStyle = fill;
        ctx.fill();
        ctx.strokeStyle = border;
        ctx.lineWidth = bw;
        ctx.stroke();
    }
    // 第二轮：所有名字最后画，保证不被任何圆圈遮挡，信息完整可读
    var labelFont = Math.max(9, Math.round(12 * radiusScale)) + 'px ' +
        ((mode === 'ios') ? '-apple-system,"Helvetica Neue","PingFang SC","Microsoft YaHei",sans-serif' : 'Microsoft YaHei, SimHei');
    ctx.font = labelFont;
    ctx.fillStyle = C.nodeText;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (var k = 0; k < nodes.length; k++) {
        var n2 = nodes[k];
        if (hideOthers && connected !== null && !connected[n2.id]) { continue; }
        ctx.fillText(n2.id, px(n2), py(n2));
    }
}

function drawTooltip() {
    if (hoverNode === null) { return; }
    var n = nodeIndex[hoverNode];
    if (!n) { return; }
    var t = nodeTotals(hoverNode);
    var lines = [hoverNode, '连接数：' + n.degree, '流入：' + fmt(t.in), '流出：' + fmt(t.out)];
    if (mode === 'classic') {
        ctx.font = '12px Microsoft YaHei, SimHei';
        var tw = 0;
        for (var i = 0; i < lines.length; i++) {
            var wl = ctx.measureText(lines[i]).width;
            if (wl > tw) { tw = wl; }
        }
        var bx = 12, by = 12;
        ctx.fillStyle = 'rgba(255,255,255,0.95)';
        ctx.strokeStyle = '#bbb';
        ctx.lineWidth = 1;
        ctx.fillRect(bx, by, tw + 16, lines.length * 18 + 10);
        ctx.strokeRect(bx, by, tw + 16, lines.length * 18 + 10);
        ctx.fillStyle = '#333';
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        for (var j = 0; j < lines.length; j++) {
            ctx.fillText(lines[j], bx + 8, by + 8 + j * 18);
        }
        return;
    }
    // iOS 风格：深色圆角悬浮卡 + 柔和阴影
    ctx.font = '12px -apple-system, "Helvetica Neue", "PingFang SC", "Microsoft YaHei", sans-serif';
    var tw2 = 0;
    for (var m = 0; m < lines.length; m++) {
        var wm = ctx.measureText(lines[m]).width;
        if (wm > tw2) { tw2 = wm; }
    }
    var bx2 = 12, by2 = 12;
    var bw2 = tw2 + 20, bh = lines.length * 19 + 14;
    ctx.save();
    ctx.shadowColor = 'rgba(0,0,0,0.28)';
    ctx.shadowBlur = 12;
    ctx.shadowOffsetY = 4;
    ctx.fillStyle = C.tooltipBg;
    roundRect(bx2, by2, bw2, bh, 12);
    ctx.fill();
    ctx.restore();
    ctx.fillStyle = C.tooltipText;
    ctx.textAlign = 'left';
    ctx.textBaseline = 'top';
    for (var p = 0; p < lines.length; p++) {
        ctx.fillText(lines[p], bx2 + 10, by2 + 7 + p * 19);
    }
}

function getMousePos(e) {
    var rect = canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
}

function nodeAt(mx, my) {
    var conn = (hideOthers && activeNode !== null) ? buildConnected() : null;
    for (var i = 0; i < nodes.length; i++) {
        var n = nodes[i];
        if (conn !== null && !conn[n.id]) { continue; }
        var cx = px(n), cy = py(n);
        var r = n.size * scale * radiusScale + 2;
        var dx = mx - cx, dy = my - cy;
        if (dx * dx + dy * dy <= r * r) { return n.id; }
    }
    return null;
}

canvas.addEventListener('click', function(e) {
    var m = getMousePos(e);
    var id = nodeAt(m.x, m.y);
    activeNode = (id !== null) ? id : null;
    draw();
});

canvas.addEventListener('mousemove', function(e) {
    var m = getMousePos(e);
    hoverNode = nodeAt(m.x, m.y);
    draw();
});

canvas.addEventListener('wheel', function(e) {
    e.preventDefault();
    var factor = (e.deltaY < 0) ? 1.1 : 0.9;
    scale = Math.min(8, Math.max(0.2, scale * factor));
    draw();
});

var dragging = false, lastX = 0, lastY = 0;
canvas.addEventListener('mousedown', function(e) {
    var m = getMousePos(e);
    if (nodeAt(m.x, m.y) !== null) { return; }
    dragging = true; lastX = e.clientX; lastY = e.clientY;
    canvas.style.cursor = 'grabbing';
});
window.addEventListener('mousemove', function(e) {
    if (!dragging) { return; }
    panX += e.clientX - lastX;
    panY += e.clientY - lastY;
    lastX = e.clientX; lastY = e.clientY;
    draw();
});
window.addEventListener('mouseup', function() {
    dragging = false; canvas.style.cursor = 'grab';
});

/* 触屏手势：点选 / 单指平移 / 双指缩放 */
var touchStartDist = 0;
canvas.addEventListener('touchstart', function(e) {
    var t = e.touches;
    if (t.length === 2) {
        var dx = t[0].clientX - t[1].clientX, dy = t[0].clientY - t[1].clientY;
        touchStartDist = Math.sqrt(dx * dx + dy * dy);
        dragging = false;
        return;
    }
    if (t.length !== 1) { return; }
    var m = getMousePos(t[0]);
    var hit = nodeAt(m.x, m.y);
    if (hit !== null) { activeNode = hit; tap(); draw(); return; }
    dragging = true; lastX = t[0].clientX; lastY = t[0].clientY;
    canvas.style.cursor = 'grabbing';
}, { passive: true });

canvas.addEventListener('touchmove', function(e) {
    var t = e.touches;
    if (t.length === 2) {
        var dx = t[0].clientX - t[1].clientX, dy = t[0].clientY - t[1].clientY;
        var d = Math.sqrt(dx * dx + dy * dy);
        if (touchStartDist > 0) {
            scale = Math.min(8, Math.max(0.2, scale * (d / touchStartDist)));
            draw();
        }
        touchStartDist = d;
        if (e.cancelable) { e.preventDefault(); }
        return;
    }
    if (!dragging) { return; }
    var p = t[0];
    panX += p.clientX - lastX;
    panY += p.clientY - lastY;
    lastX = p.clientX; lastY = p.clientY;
    draw();
    if (e.cancelable) { e.preventDefault(); }
}, { passive: false });

canvas.addEventListener('touchend', function() {
    dragging = false; touchStartDist = 0;
    canvas.style.cursor = 'grab';
});

document.getElementById('cbAmount').addEventListener('change', function() {
    showAmount = this.checked; draw();
});
document.getElementById('cbHideOther').addEventListener('change', function() {
    hideOthers = this.checked;
    draw();
});
var titleText = document.getElementById('titleText');
document.getElementById('titleInput').addEventListener('input', function() {
    var v = this.value || '资金流向图';
    titleText.textContent = v;
    document.title = v;
});
var settingsPanel = document.getElementById('settingsPanel');
function openSettings() {
    settingsPanel.className = 'show';
    settingsPanel.setAttribute('aria-hidden', 'false');
}
function closeSettings() {
    settingsPanel.className = '';
    settingsPanel.setAttribute('aria-hidden', 'true');
}
document.getElementById('settingsBtn').addEventListener('click', function() {
    tap();
    if (settingsPanel.className === 'show') { closeSettings(); } else { openSettings(); }
});
document.getElementById('settingsClose').addEventListener('click', function() {
    tap(); closeSettings();
});
settingsPanel.addEventListener('click', function(e) {
    if (e.target === settingsPanel) { closeSettings(); }
});
window.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && settingsPanel.className === 'show') { closeSettings(); }
});
var segBtns = document.querySelectorAll('.seg-btn');
for (var q = 0; q < segBtns.length; q++) {
    (function(btn) {
        btn.addEventListener('click', function() {
            tap();
            applyMode(btn.getAttribute('data-mode'));
        });
    })(segBtns[q]);
}
document.getElementById('refreshBtn').addEventListener('click', function() {
    tap();
    activeNode = null;
    scale = 1; panX = 0; panY = 0;
    draw();
});
document.getElementById('zoomIn').addEventListener('click', function() {
    tap();
    scale = Math.min(8, scale * 1.25);
    draw();
});
document.getElementById('zoomOut').addEventListener('click', function() {
    tap();
    scale = Math.max(0.2, scale / 1.25);
    draw();
});

window.addEventListener('resize', resizeCanvas);
</script>
</body>
</html>'''


html = generate_html(contract, title='资金流水分析演示图', hide_other=True)
with open('资金流向图.html', 'w', encoding='utf-8') as f:
    f.write(html)
if sys.stdout is not None:
    print('已生成 资金流向图.html，双击即可在浏览器中离线使用。')