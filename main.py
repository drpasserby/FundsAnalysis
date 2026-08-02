"""
资金流水走向分析工具
版本：1.0.3
作者：wulvxinchen
"""


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import sys
import json
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from collections import defaultdict


opt_root = tk.Tk()
opt_root.title('生成选项')
opt_root.geometry('360x280+600+400')

show_amount = tk.BooleanVar(value=True)
merge_edges = tk.BooleanVar(value=False)
custom_title = tk.StringVar(value='演示图')
file_path_var = tk.StringVar(value='未选择文件')

tk.Label(opt_root, text='请选择绘图选项：', font=('', 10, 'bold')).pack(pady=5)

cb1 = tk.Checkbutton(opt_root, text='显示金额', variable=show_amount)
cb1.pack(anchor='w', padx=20, pady=3)

cb2 = tk.Checkbutton(opt_root, text='收入/支出合并显示', variable=merge_edges)
cb2.pack(anchor='w', padx=20, pady=3)

tk.Label(opt_root, text='自定义标题：').pack(anchor='w', padx=20, pady=(8,0))
title_entry = tk.Entry(opt_root, textvariable=custom_title, width=30)
title_entry.pack(padx=20, pady=3)

def choose_file():
    path = filedialog.askopenfilename(
        title='选择数据文件',
        filetypes=[('Excel 文件', '*.xlsx;*.xls'), ('所有文件', '*.*')])
    if path:
        file_path_var.set(path)

tk.Label(opt_root, text='数据文件：').pack(anchor='w', padx=20, pady=(8,0))
file_frame = tk.Frame(opt_root)
file_frame.pack(fill='x', padx=20, pady=3)
tk.Button(file_frame, text='选择文件...', command=choose_file, width=12).pack(side='left')
tk.Label(file_frame, textvariable=file_path_var, fg='gray', anchor='w', width=20).pack(side='left', padx=8)

def confirm():
    if file_path_var.get() == '未选择文件':
        messagebox.showwarning('提示', '请先选择数据文件！')
        return
    opt_root.destroy()

tk.Button(opt_root, text='开始生成', command=confirm, width=12).pack(pady=10)
opt_root.mainloop()

file_path = file_path_var.get()
if file_path == '未选择文件':
    sys.exit('未选择数据文件，程序退出。')

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

# ================= 图构建（统一，一次遍历产出两套图） =================
def build_graphs(df):
    """一次遍历构建两套有向图：
    G_sep —— 收入/支出不合并：支出边 a→c，收入边 c→a；
    G_mer —— 同一对节点合并存储收支，双向分别画边。
    """
    edge_sep = defaultdict(lambda: {'weight': 0.0, 'type': None})
    edge_mer = defaultdict(lambda: {'to_weight': 0.0, 'from_weight': 0.0})

    for _, row in df.iterrows():
        a = str(row.iloc[0]).strip()
        direction = str(row.iloc[1]).strip()
        c = str(row.iloc[2]).strip()
        amount = float(row.iloc[3])

        if a == c:
            continue

        if '支出' in direction:
            edge_sep[(a, c)]['weight'] += amount
            edge_sep[(a, c)]['type'] = '支出'
            edge_mer[(a, c)]['to_weight'] += amount
        elif '收入' in direction:
            edge_sep[(c, a)]['weight'] += amount
            edge_sep[(c, a)]['type'] = '收入'
            edge_mer[(a, c)]['from_weight'] += amount

    G_sep = nx.DiGraph()
    for (u, v), info in edge_sep.items():
        G_sep.add_edge(u, v, weight=info['weight'], etype=info['type'])

    G_mer = nx.DiGraph()
    for (u, v), info in edge_mer.items():
        if info['to_weight'] > 0:
            G_mer.add_edge(u, v, weight=info['to_weight'], etype='支出')
        if info['from_weight'] > 0:
            G_mer.add_edge(v, u, weight=info['from_weight'], etype='收入')

    return G_sep, G_mer


G_sep, G_mer = build_graphs(df)
G = G_mer if merge_edges.get() else G_sep

# ================= 数据契约（供 HTML 渲染使用） =================
def build_contract(G_sep, G_mer):
    """生成 HTML 端使用的数据契约：坐标归一化到 [0,1]，节点大小设上限。
    返回 {'nodes': [...], 'edgesSep': [...], 'edgesMer': [...]}
    """
    all_nodes = list(set(G_sep.nodes()) | set(G_mer.nodes()))

    pos = nx.spring_layout(G_sep, k=4, seed=42) if len(G_sep.nodes()) > 0 else {}

    degree_dict = {}
    for node in all_nodes:
        pred = set(G_sep.predecessors(node)) if node in G_sep else set()
        succ = set(G_sep.successors(node)) if node in G_sep else set()
        degree_dict[node] = len(pred | succ)

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

    # 节点大小：基础值 + 度 * 缩放，设上限防止枢纽节点过大
    base_size = 500
    scale = 300
    max_size = 3000
    size_dict = {n: min(base_size + scale * degree_dict[n], max_size) for n in all_nodes}
    max_s = max(size_dict.values()) if size_dict else 1
    min_s = min(size_dict.values()) if size_dict else 1

    nodes_json = []
    for node in all_nodes:
        if node in pos:
            x = float((pos[node][0] - minx) / (maxx - minx))
            y = float((pos[node][1] - miny) / (maxy - miny))
        else:
            x, y = 0.5, 0.5
        raw_s = size_dict[node]
        display_s = 15 + (raw_s - min_s) / (max_s - min_s) * 40 if max_s != min_s else 30
        nodes_json.append({'id': node, 'x': x, 'y': y, 'size': display_s, 'degree': degree_dict[node]})

    def edge_list(G):
        edges = []
        for u, v in G.edges():
            info = G[u][v]
            rad = 0.0
            if G.has_edge(v, u):
                rad = 0.25 if u < v else -0.25
            edges.append({
                'source': u,
                'target': v,
                'type': info.get('etype', ''),
                'amount': info['weight'],
                'rad': rad
            })
        return edges

    return {
        'nodes': nodes_json,
        'edgesSep': edge_list(G_sep),
        'edgesMer': edge_list(G_mer),
    }


contract = build_contract(G_sep, G_mer)
with open('资金流向图.json', 'w', encoding='utf-8') as f:
    json.dump(contract, f, ensure_ascii=False, indent=2)
if sys.stdout is not None:
    print('已生成 资金流向图.json（数据契约，坐标已归一化到 [0,1]）。')

degree_dict = {}
for node in G.nodes():
    neighbors = set(G.predecessors(node)) | set(G.successors(node))
    degree_dict[node] = len(neighbors)

base_size = 500
scale = 300
node_size = [base_size + scale * degree_dict[node] for node in G.nodes()]

plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei']
plt.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(20, 16))
plt.subplots_adjust(bottom=0.1)
pos = nx.spring_layout(G, k=4, seed=42)

node_list = list(G.nodes())
node_colors = ['lightblue'] * len(node_list)
node_edgecolors = ['black'] * len(node_list)
node_linewidths = [0.8] * len(node_list)
node_zorders = [1] * len(node_list)

drawn_nodes = nx.draw_networkx_nodes(G, pos, node_size=node_size, node_color=node_colors,
                       edgecolors=node_edgecolors, linewidths=node_linewidths, ax=ax)
drawn_labels = nx.draw_networkx_labels(G, pos, font_size=10, ax=ax)

for text in drawn_labels.values():
    text.set_zorder(12)

edge_lines = {}
edge_texts = {}

for u, v in G.edges():
    edge_info = G[u][v]
    etype = edge_info.get('etype', '')
    weight = edge_info['weight']

    if show_amount.get():
        label = f'{etype} {weight:,.0f}' if etype else f'{weight:,.0f}'
    else:
        label = ''

    rad = 0.0
    if G.has_edge(v, u):
        if u < v:
            rad = 0.25
        else:
            rad = -0.25

    line_list = nx.draw_networkx_edges(G, pos, edgelist=[(u, v)],
                           connectionstyle=f'arc3, rad={rad}',
                           arrowstyle='-|>', arrowsize=15,
                           edge_color='gray', alpha=0.7, ax=ax)
    edge_lines[(u, v)] = line_list[0]

    if show_amount.get() and label:
        x1, y1 = pos[u]
        x2, y2 = pos[v]
        dx, dy = x2 - x1, y2 - y1
        length = (dx**2 + dy**2) ** 0.5
        if length == 0:
            continue
        nx_dir = -dy / length
        ny_dir = dx / length
        offset = rad * length
        xm = (x1 + x2) / 2 + offset * nx_dir
        ym = (y1 + y2) / 2 + offset * ny_dir

        text = ax.text(xm, ym, label, fontsize=8, color='darkred', ha='center', va='center',
                bbox=dict(facecolor='white', alpha=0.8, edgecolor='none', boxstyle='round,pad=0.2'))
        edge_texts[(u, v)] = text

ax.axis('off')
plt.title(custom_title.get(), fontsize=14)

ax.text(0.02, 0.02, '资金流水走向分析工具Github@drpasserby(wlxc)', transform=ax.transAxes,
        fontsize=15, color='#bbb', alpha=0.6, ha='left', va='bottom', zorder=20)

initial_node_colors = node_colors.copy()
initial_node_edgecolors = node_edgecolors.copy()
initial_node_linewidths = node_linewidths.copy()
initial_edge_alphas = {e: 0.7 for e in edge_lines}
initial_text_alphas = {e: 1.0 for e in edge_texts}

def reset_graph(event):
    for i in range(len(node_list)):
        node_colors[i] = initial_node_colors[i]
        node_edgecolors[i] = initial_node_edgecolors[i]
        node_linewidths[i] = initial_node_linewidths[i]
        node_zorders[i] = 1

    drawn_nodes.set_facecolor(node_colors)
    drawn_nodes.set_edgecolor(node_edgecolors)
    drawn_nodes.set_linewidth(node_linewidths)

    for e, obj in edge_lines.items():
        obj.set_alpha(initial_edge_alphas[e])
        obj.set_zorder(1)
    for e, text in edge_texts.items():
        text.set_alpha(initial_text_alphas[e])
        text.set_zorder(2)

    for text in drawn_labels.values():
        text.set_zorder(12)

    fig.canvas.draw_idle()

reset_ax = plt.axes([0.8, 0.02, 0.1, 0.04])
reset_btn = Button(reset_ax, '重置')
reset_btn.on_clicked(reset_graph)

def on_node_click(event):
    if event.inaxes != ax:
        return

    clicked_node = None
    min_dist = float('inf')
    for node, (x, y) in pos.items():
        dist = ((event.xdata - x)**2 + (event.ydata - y)**2)**0.5
        idx = node_list.index(node)
        threshold = (node_size[idx] / 3.14)**0.5 * 0.015
        if dist < threshold and dist < min_dist:
            clicked_node = node
            min_dist = dist

    if clicked_node is None:
        return

    connected_nodes = set()
    for neighbor in G.predecessors(clicked_node):
        connected_nodes.add(neighbor)
    for neighbor in G.successors(clicked_node):
        connected_nodes.add(neighbor)
    connected_nodes.add(clicked_node)

    for i, node in enumerate(node_list):
        if node == clicked_node:
            node_colors[i] = 'orange'
            node_edgecolors[i] = 'red'
            node_linewidths[i] = 3.0
            node_zorders[i] = 10
        elif node in connected_nodes:
            node_colors[i] = 'lightgreen'
            node_edgecolors[i] = 'green'
            node_linewidths[i] = 2.0
            node_zorders[i] = 9
        else:
            node_colors[i] = 'lightgray'
            node_edgecolors[i] = 'gray'
            node_linewidths[i] = 0.5
            node_zorders[i] = 1

    drawn_nodes.set_facecolor(node_colors)
    drawn_nodes.set_edgecolor(node_edgecolors)
    drawn_nodes.set_linewidth(node_linewidths)

    for text in drawn_labels.values():
        text.set_zorder(12)

    for (u, v), obj in edge_lines.items():
        if u == clicked_node or v == clicked_node:
            obj.set_alpha(1.0)
            obj.set_zorder(9)
        elif u in connected_nodes and v in connected_nodes:
            obj.set_alpha(0.9)
            obj.set_zorder(8)
        else:
            obj.set_alpha(0.15)
            obj.set_zorder(1)

    for (u, v), text in edge_texts.items():
        if u == clicked_node or v == clicked_node:
            text.set_alpha(1.0)
            text.set_zorder(10)
        elif u in connected_nodes and v in connected_nodes:
            text.set_alpha(0.9)
            text.set_zorder(9)
        else:
            text.set_alpha(0.3)
            text.set_zorder(1)

    fig.canvas.draw_idle()

fig.canvas.mpl_connect('button_press_event', on_node_click)

plt.show()