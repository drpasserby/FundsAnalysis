"""
资金流水走向分析工具
版本：1.2.3
作者：wulvxinchen
"""

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import sys
import json
import math
import re
import pandas as pd
import networkx as nx
from collections import defaultdict


# ================= 选择数据文件与生成风格 =================
def choose_file_and_style(root):
    """弹窗：使用提示 + 选择数据文件 + 选择生成的 HTML 风格（默认经典布局）。
    返回 (file_path, style)，style 为 'classic' 或 'ios'；用户取消则返回 None。"""
    dialog = tk.Toplevel(root)
    dialog.title('生成资金流向图')
    dialog.resizable(False, False)
    dialog.attributes('-topmost', True)  # 置顶显示
    dialog.grab_set()  # 模态窗口

    result = {'path': None, 'style': 'classic'}

    # 使用提示：说明可读文件格式与表格内容格式
    tk.Label(dialog, text='可读取的文件格式：Excel（.xlsx / .xls）',
             anchor='w', justify='left').pack(fill='x', padx=16, pady=(14, 2))
    tk.Label(dialog, text='表格内容格式（第 1 行为表头）：',
             anchor='w', justify='left').pack(fill='x', padx=16)
    tk.Label(dialog, text='用户方 | 支出/收入 | 客户方 | 金额（元）',
             anchor='w', justify='left', font=('Microsoft YaHei UI', 10, 'bold')).pack(fill='x', padx=16, pady=(0, 2))
    tk.Label(dialog, text='（若表头与预设不一致，仍按默认列位读取：\n第 1 列=用户方、第 2 列=支出/收入、第 3 列=客户方、第 4 列=金额）',
             anchor='w', justify='left', fg='#555555').pack(fill='x', padx=16, pady=(0, 8))

    # 文件选择行
    file_var = tk.StringVar(value='（未选择文件）')
    file_frame = tk.Frame(dialog)
    tk.Entry(file_frame, textvariable=file_var, width=34, state='readonly').pack(side='left', padx=(12, 6))

    def pick_file():
        p = filedialog.askopenfilename(
            parent=dialog, title='选择数据文件',
            filetypes=[('Excel 文件', '*.xlsx;*.xls'), ('所有文件', '*.*')])
        if p:
            result['path'] = p
            file_var.set(os.path.basename(p))

    tk.Button(file_frame, text='选择文件...', command=pick_file).pack(side='left')
    file_frame.pack(pady=6)

    # 风格选择行
    style_var = tk.StringVar(value='classic')
    tk.Label(dialog, text='HTML 风格：', anchor='w').pack(fill='x', padx=16)
    tk.Radiobutton(dialog, text='经典布局（默认）', value='classic', variable=style_var).pack(anchor='w', padx=28)
    tk.Radiobutton(dialog, text='iOS 布局', value='ios', variable=style_var).pack(anchor='w', padx=28)

    def confirm():
        if not result['path']:
            messagebox.showwarning('提示', '请先选择数据文件。', parent=dialog)
            return
        result['style'] = style_var.get()
        dialog.destroy()

    def cancel():
        result['path'] = None
        dialog.destroy()

    tk.Button(dialog, text='生成', width=14, command=confirm).pack(pady=(8, 4))
    dialog.protocol('WM_DELETE_WINDOW', cancel)

    dialog.wait_window()  # 阻塞直到对话框关闭

    if not result['path']:
        return None
    return result['path'], result['style']


root = tk.Tk()
root.withdraw()  # 隐藏主窗口

choice = choose_file_and_style(root)
if choice is None:
    sys.exit()  # 用户取消选择，正常退出
file_path, html_style = choice

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

    # 表头软校验（仅提示，不阻断）：列位始终按默认顺序读取，不依赖表头文字。
    # 若表头与预设不一致，仍按 第 1 列=用户方、第 2 列=支出/收入、第 3 列=客户方、第 4 列=金额 读取。
    headers = [str(h) for h in raw_df.columns[:4]]
    if not ('收入' in headers[1] or '支出' in headers[1]) or '金额' not in headers[3]:
        messages.append('提示：表头与预设不一致（应为 用户方|支出/收入|客户方|金额），'
                        '已按默认列位读取：第 1 列=用户方，第 2 列=支出/收入，第 3 列=客户方，第 4 列=金额')

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

        # 金额必须是有限的非负数（拦截 NaN / ±Inf / 负数，避免污染后续求和）
        if not math.isfinite(amount) or amount < 0:
            bad_rows.append('第 {} 行：金额“{}”不是有效的非负数字'.format(row_no, amount))
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
# 模板：与 main.py 同目录的 template.html，占位符 {{TOKEN}} 由 generate_html 单遍替换。
# 输出仍是自包含单文件 HTML；模板与数据分离，避免在三引号字符串里拼接代码的引号边界问题。
TEMPLATE_FILE = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'template.html')


def generate_html(contract, title='资金流水分析演示图', show_amount=True, hide_other=True,
                  default_mode='classic', template_file=None):
    """生成自包含的交互式 HTML：数据内嵌、离线可用（JS 保持 ES5 兼容旧浏览器）。
    default_mode：初始界面风格（'classic' 经典布局 / 'ios' iOS 布局），默认经典布局。
    template_file：HTML 模板路径，缺省用同目录 template.html（测试可显式注入）。
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

    # 初始界面风格：经典布局（默认）或 iOS 布局
    if default_mode == 'ios':
        body_class = 'mode-ios'
        mode_js = 'ios'
        seg_classic = 'seg-btn'
        seg_ios = 'seg-btn active'
    else:
        body_class = 'mode-classic'
        mode_js = 'classic'
        seg_classic = 'seg-btn active'
        seg_ios = 'seg-btn'

    tokens = {
        'NODES_JSON': nodes_json,
        'EDGES_JSON': edges_json,
        'SAFE_TITLE': safe_title,
        'AMT_CHECKED': amt_checked,
        'HIDE_CHECKED': hide_checked,
        'SHOW_JS': show_js,
        'HIDE_JS': hide_js,
        'BODY_CLASS': body_class,
        'SEG_CLASSIC': seg_classic,
        'SEG_IOS': seg_ios,
        'MODE_JS': mode_js,
    }
    path = template_file or TEMPLATE_FILE
    with open(path, encoding='utf-8') as f:
        html = f.read()
    # 单遍替换：只扫描模板原文，占位符的值不会再被当作占位符二次替换
    html = re.sub(r'\{\{([A-Z_]+)\}\}', lambda m: tokens[m.group(1)], html)
    return html



html = generate_html(contract, title='资金流水分析演示图', hide_other=True, default_mode=html_style)
with open('资金流向图.html', 'w', encoding='utf-8') as f:
    f.write(html)
if sys.stdout is not None:
    print('已生成 资金流向图.html，双击即可在浏览器中离线使用。')