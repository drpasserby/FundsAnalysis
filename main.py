"""
资金流水走向分析工具
版本：1.1.1
作者：wulvxinchen
"""


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import sys
import json
import pandas as pd
import networkx as nx
from collections import defaultdict


# ================= 选择数据文件 =================
root = tk.Tk()
root.withdraw()  # 隐藏主窗口，直接弹出文件选择框
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


# ================= HTML 输出（自包含交互式查看器） =================
def generate_html(contract, title='演示图', show_amount=True, merge_edges=False):
    """生成自包含的交互式 HTML：数据内嵌、离线可用、兼容旧浏览器（ES5 语法）。
    页面内可实时切换：显示金额 / 合并收支 / 自定义标题；
    支持点击高亮、悬浮提示、滚轮缩放、拖拽平移、重置。
    """
    nodes_json = json.dumps(contract['nodes'], ensure_ascii=False).replace('</', '<\\/')
    edges_sep_json = json.dumps(contract['edgesSep'], ensure_ascii=False).replace('</', '<\\/')
    edges_mer_json = json.dumps(contract['edgesMer'], ensure_ascii=False).replace('</', '<\\/')
    safe_title = (title.replace('&', '&amp;').replace('<', '&lt;')
                       .replace('>', '&gt;').replace('"', '&quot;'))
    amt_checked = ' checked' if show_amount else ''
    merge_checked = ' checked' if merge_edges else ''
    show_js = 'true' if show_amount else 'false'
    merge_js = 'true' if merge_edges else 'false'

    return '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>资金流向图</title>
<style>
body { margin:0; display:flex; flex-direction:column; align-items:center;
       font-family:'Microsoft YaHei','SimHei',sans-serif; background:#fafafa; }
#controls { margin:12px 0 4px; display:flex; flex-wrap:wrap; align-items:center; gap:14px; font-size:14px; }
#controls label { cursor:pointer; }
#titleText { margin:8px 0 2px; font-size:20px; color:#333; }
#canvas { border:1px solid #ccc; background:#fff; cursor:grab; }
#legend { font-size:13px; color:#555; }
</style>
</head>
<body>
<h2 id="titleText">''' + safe_title + '''</h2>
<div id="controls">
  <label><input type="checkbox" id="cbAmount"''' + amt_checked + '''> 显示金额</label>
  <label><input type="checkbox" id="cbMerge"''' + merge_checked + '''> 收入/支出合并显示</label>
  <span>标题：<input type="text" id="titleInput" value="''' + safe_title + '''"></span>
  <button id="resetBtn" type="button">重置</button>
  <span id="legend"><span style="color:#c0392b;">——支出</span>
  <span style="color:#2e8b57;">——收入</span></span>
</div>
<canvas id="canvas"></canvas>
<script>
var nodes = ''' + nodes_json + ''';
var edgesSep = ''' + edges_sep_json + ''';
var edgesMer = ''' + edges_mer_json + ''';
var PADDING = 80;
var scale = 1, panX = 0, panY = 0;
var activeNode = null, hoverNode = null;
var showAmount = ''' + show_js + ''';
var mergeEdges = ''' + merge_js + ''';
var currentEdges = mergeEdges ? edgesMer : edgesSep;
var canvas = document.getElementById('canvas');
var ctx = canvas.getContext('2d');
var W, H;

var nodeIndex = Object.create(null);
for (var k = 0; k < nodes.length; k++) { nodeIndex[nodes[k].id] = nodes[k]; }

function resizeCanvas() {
    var w = Math.min(1400, (window.innerWidth || 1200) - 40);
    var h = Math.min(900, (window.innerHeight || 800) - 130);
    canvas.width = Math.max(600, w);
    canvas.height = Math.max(400, h);
    W = canvas.width; H = canvas.height;
    draw();
}
resizeCanvas();

function px(n) { return (PADDING + n.x * (W - 2 * PADDING)) * scale + panX; }
function py(n) { return (PADDING + n.y * (H - 2 * PADDING)) * scale + panY; }

function fmt(n) { return Math.round(n).toLocaleString(); }

function nodeTotals(id) {
    var tin = 0, tout = 0;
    for (var i = 0; i < currentEdges.length; i++) {
        var e = currentEdges[i];
        if (e.target === id) { tin += e.amount; }
        if (e.source === id) { tout += e.amount; }
    }
    return { in: tin, out: tout };
}

function draw() {
    ctx.clearRect(0, 0, W, H);
    drawEdges();
    drawWatermark();
    drawNodes();
    drawTooltip();
}

function drawWatermark() {
    ctx.font = '2em Microsoft YaHei, SimHei';
    ctx.fillStyle = '#ddd';
    ctx.textAlign = 'right';
    ctx.textBaseline = 'bottom';
    ctx.fillText('资金流水走向分析工具 Github@drpasserby(WLXC)', W - 16, H - 12);
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

function edgeColor(type) {
    if (type === '支出') { return '#c0392b'; }
    if (type === '收入') { return '#2e8b57'; }
    return '#888';
}

function drawEdges() {
    var connected = null;
    if (activeNode !== null) { connected = Object.create(null); connected[activeNode] = true; }
    for (var i = 0; i < currentEdges.length; i++) {
        var e = currentEdges[i];
        var sn = nodeIndex[e.source];
        var tn = nodeIndex[e.target];
        if (!sn || !tn) { continue; }
        var sx = px(sn), sy = py(sn);
        var tx = px(tn), ty = py(tn);
        var dx = tx - sx, dy = ty - sy;
        var len = Math.sqrt(dx * dx + dy * dy);
        if (len === 0) { continue; }
        var nvx = -dy / len, nvy = dx / len;
        var rad = e.rad || 0;
        var off = rad * len;
        var mx = (sx + tx) / 2 + off * nvx;
        var my = (sy + ty) / 2 + off * nvy;

        var alpha = 0.7, lineColor = edgeColor(e.type), textAlpha = 1.0;
        if (activeNode !== null) {
            if (e.source === activeNode || e.target === activeNode) { alpha = 1.0; }
            else if (connected[e.source] && connected[e.target]) { alpha = 0.9; }
            else { alpha = 0.15; textAlpha = 0.3; }
        }

        ctx.beginPath();
        ctx.moveTo(sx, sy);
        if (rad !== 0) { ctx.quadraticCurveTo(mx, my, tx, ty); }
        else { ctx.lineTo(tx, ty); }
        ctx.strokeStyle = lineColor;
        ctx.lineWidth = 1.5;
        ctx.globalAlpha = alpha;
        ctx.stroke();
        ctx.globalAlpha = 1;

        var ax2, ay2, aAng;
        if (rad !== 0) {
            var tgx = tx - mx, tgy = ty - my;
            var tl = Math.sqrt(tgx * tgx + tgy * tgy) || 1;
            ax2 = tx - tgx / tl * 12;
            ay2 = ty - tgy / tl * 12;
            aAng = Math.atan2(tgy, tgx);
        } else {
            ax2 = tx - dx / len * 12;
            ay2 = ty - dy / len * 12;
            aAng = Math.atan2(dy, dx);
        }
        drawArrow(ax2, ay2, aAng, lineColor, (activeNode === null) ? 0.9 : alpha);

        if (showAmount) {
            var label = e.type + ' ' + fmt(e.amount);
            ctx.font = '11px Microsoft YaHei, SimHei';
            var tw = ctx.measureText(label).width;
            var lx = (rad !== 0) ? mx : (sx + tx) / 2;
            var ly = (rad !== 0) ? my : (sy + ty) / 2;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.globalAlpha = textAlpha;
            ctx.fillStyle = '#fff';
            ctx.fillRect(lx - tw / 2 - 4, ly - 9, tw + 8, 18);
            ctx.fillStyle = '#333';
            ctx.fillText(label, lx, ly);
            ctx.globalAlpha = 1;
        }
    }
}

function drawNodes() {
    var connected = null;
    if (activeNode !== null) {
        connected = Object.create(null);
        for (var i = 0; i < currentEdges.length; i++) {
            var e = currentEdges[i];
            if (e.source === activeNode) { connected[e.target] = true; }
            if (e.target === activeNode) { connected[e.source] = true; }
        }
        connected[activeNode] = true;
    }
    for (var j = 0; j < nodes.length; j++) {
        var n = nodes[j];
        var cx = px(n), cy = py(n);
        var r = n.size * scale;
        var fill = '#dbe9fb', border = '#333', bw = 0.8;
        if (activeNode !== null) {
            if (n.id === activeNode) { fill = '#f39c12'; border = '#c0392b'; bw = 3; }
            else if (connected[n.id]) { fill = '#a9dfbf'; border = '#2e8b57'; bw = 2; }
            else { fill = '#e5e5e5'; border = '#999'; bw = 0.5; }
        }
        ctx.beginPath();
        ctx.arc(cx, cy, r, 0, 2 * Math.PI);
        ctx.fillStyle = fill;
        ctx.fill();
        ctx.strokeStyle = border;
        ctx.lineWidth = bw;
        ctx.stroke();
        ctx.font = '12px Microsoft YaHei, SimHei';
        ctx.fillStyle = '#000';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(n.id, cx, cy);
    }
}

function drawTooltip() {
    if (hoverNode === null) { return; }
    var n = nodeIndex[hoverNode];
    if (!n) { return; }
    var t = nodeTotals(hoverNode);
    var lines = [hoverNode, '连接数：' + n.degree, '流入：' + fmt(t.in), '流出：' + fmt(t.out)];
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
}

function getMousePos(e) {
    var rect = canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
}

function nodeAt(mx, my) {
    for (var i = 0; i < nodes.length; i++) {
        var n = nodes[i];
        var cx = px(n), cy = py(n);
        var r = n.size * scale + 2;
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

document.getElementById('cbAmount').addEventListener('change', function() {
    showAmount = this.checked; draw();
});
document.getElementById('cbMerge').addEventListener('change', function() {
    mergeEdges = this.checked;
    currentEdges = mergeEdges ? edgesMer : edgesSep;
    activeNode = null;
    draw();
});
var titleText = document.getElementById('titleText');
document.getElementById('titleInput').addEventListener('input', function() {
    var v = this.value || '资金流向图';
    titleText.textContent = v;
    document.title = v;
});
document.getElementById('resetBtn').addEventListener('click', function() {
    activeNode = null;
    scale = 1; panX = 0; panY = 0;
    draw();
});

window.addEventListener('resize', resizeCanvas);
</script>
</body>
</html>'''


html = generate_html(contract)
with open('资金流向图.html', 'w', encoding='utf-8') as f:
    f.write(html)
if sys.stdout is not None:
    print('已生成 资金流向图.html，双击即可在浏览器中离线使用。')