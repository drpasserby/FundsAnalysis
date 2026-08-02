"""
资金流水走向分析工具
版本：1.0.0
作者：wulvxinchen
"""


import tkinter as tk
from tkinter import messagebox
import sys
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from matplotlib.widgets import Button
from collections import defaultdict


opt_root = tk.Tk()
opt_root.title('生成选项')
opt_root.geometry('350x200+600+400')

show_amount = tk.BooleanVar(value=True)
merge_edges = tk.BooleanVar(value=False)
custom_title = tk.StringVar(value='演示图')

tk.Label(opt_root, text='请选择绘图选项：', font=('', 10, 'bold')).pack(pady=5)

cb1 = tk.Checkbutton(opt_root, text='显示金额', variable=show_amount)
cb1.pack(anchor='w', padx=20, pady=3)

cb2 = tk.Checkbutton(opt_root, text='收入/支出合并显示', variable=merge_edges)
cb2.pack(anchor='w', padx=20, pady=3)

tk.Label(opt_root, text='自定义标题：').pack(anchor='w', padx=20, pady=(8,0))
title_entry = tk.Entry(opt_root, textvariable=custom_title, width=30)
title_entry.pack(padx=20, pady=3)

def confirm():
    opt_root.destroy()

tk.Button(opt_root, text='开始生成', command=confirm, width=12).pack(pady=10)
opt_root.mainloop()

df = pd.read_excel('流水透视.xlsx', sheet_name='Sheet1', header=0)

if merge_edges.get():
    edge_data = defaultdict(lambda: {'to_weight': 0.0, 'from_weight': 0.0})

    for _, row in df.iterrows():
        a = str(row.iloc[0]).strip()
        direction = str(row.iloc[1]).strip()
        c = str(row.iloc[2]).strip()
        amount = float(row.iloc[3])

        if a == c:
            continue

        if '支出' in direction:
            edge_data[(a, c)]['to_weight'] += amount
        elif '收入' in direction:
            edge_data[(a, c)]['from_weight'] += amount

    G = nx.DiGraph()
    for (u, v), info in edge_data.items():
        if info['to_weight'] > 0:
            G.add_edge(u, v, weight=info['to_weight'], etype='支出', merged=True, opposite_weight=info['from_weight'])
        if info['from_weight'] > 0:
            G.add_edge(v, u, weight=info['from_weight'], etype='收入', merged=True, opposite_weight=info['to_weight'])

else:
    edge_data = defaultdict(lambda: {'weight': 0.0, 'type': None})

    for _, row in df.iterrows():
        a = str(row.iloc[0]).strip()
        direction = str(row.iloc[1]).strip()
        c = str(row.iloc[2]).strip()
        amount = float(row.iloc[3])

        if '支出' in direction:
            source, target, etype = a, c, '支出'
        elif '收入' in direction:
            source, target, etype = c, a, '收入'
        else:
            continue

        if source == target:
            continue

        key = (source, target)
        edge_data[key]['weight'] += amount
        edge_data[key]['type'] = etype

    G = nx.DiGraph()
    for (u, v), info in edge_data.items():
        G.add_edge(u, v, weight=info['weight'], etype=info['type'])

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