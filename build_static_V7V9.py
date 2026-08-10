"""
build_static_V7V9.py
把 V7/V9 16y trades + Daily PENDING + K 線 PNG 打包成靜態檔, 推到 GitHub Pages.

Output:
  static_V7V9_site/
    index.html           ← 主頁
    data/v7.json         ← V7 16y trades
    data/v9.json         ← V9 16y trades
    data/daily.json      ← Daily PENDING trades (從 CSV)
    charts/v7/*.png      ← V7 K 線
    charts/v9/*.png      ← V9 K 線

Note: GitHub Pages 是純 client-side, 所以 daily 的 PENDING 顯示靜態資料
(出場價 = N/A, 現價 = N/A - 沒 yfinance). 完整即時抓要看本地 FastAPI 版
(paper_trade_V7_V9_web.py @ localhost:8768).
"""

import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent
OUT = ROOT / 'static_V7V9_site'

EXCEL_V7 = ROOT / 'paper_trade_V7_16y_trades.xlsx'
EXCEL_V9 = ROOT / 'paper_trade_V9_16y_trades.xlsx'
CHARTS_V7 = ROOT / 'paper_trade_V7_16y_charts'
CHARTS_V9 = ROOT / 'paper_trade_V9_16y_charts'
CSV_DAILY = ROOT / 'paper_trade_V7_trades.csv'


def load_trades(excel_path, label):
    import openpyxl
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    trades = []
    for r_idx in range(5, ws.max_row + 1):
        sym = ws.cell(row=r_idx, column=3).value
        if not (sym and isinstance(sym, str) and sym.isdigit() and len(sym) == 4):
            continue
        idx = ws.cell(row=r_idx, column=1).value
        png_rel = f"charts/{label}/{idx:03d}_{sym}_{ws.cell(row=r_idx, column=4).value}_{ws.cell(row=r_idx, column=13).value}.png"
        trades.append({
            'idx': idx,
            'win': ws.cell(row=r_idx, column=2).value,
            'sym': sym,
            'entry_date': ws.cell(row=r_idx, column=4).value,
            'first_cost': ws.cell(row=r_idx, column=5).value,
            'buy_count': ws.cell(row=r_idx, column=6).value,
            'avg_cost': ws.cell(row=r_idx, column=7).value,
            'exit_date': ws.cell(row=r_idx, column=8).value,
            'exit_price': ws.cell(row=r_idx, column=9).value,
            'hold_days': ws.cell(row=r_idx, column=10).value,
            'pnl_pct': ws.cell(row=r_idx, column=11).value,
            'pnl_dollar': ws.cell(row=r_idx, column=12).value,
            'reason': ws.cell(row=r_idx, column=13).value,
            'detail': ws.cell(row=r_idx, column=14).value,
            'png_url': png_rel,
        })
    return trades


def load_daily_trades(csv_path):
    """從 paper_trade_V7_trades.csv 載入 daily trades (含 PENDING)"""
    if not csv_path.exists():
        return [], 0
    import pandas as pd
    df = pd.read_csv(csv_path)
    if df.empty:
        return [], 0
    all_trades = {}
    final_pnl = 0
    for _, row in df.iterrows():
        sym = str(row['sym'])
        if row['action'] == 'entry':
            all_trades[sym] = {
                'idx': len(all_trades) + 1,
                'win': 'PENDING',
                'sym': sym,
                'entry_date': row['date'],
                'first_cost': row['first_cost'],
                'buy_count': 1,
                'avg_cost': row['first_cost'],
                'exit_date': '',
                'exit_price': 0,
                'hold_days': 0,
                'pnl_pct': 0,
                'pnl_dollar': 0,
                'reason': 'Holding',
                'detail': '',
                'png_url': '',
            }
        elif row['action'] == 'add' and sym in all_trades:
            all_trades[sym]['buy_count'] += 1
        elif row['action'] == 'exit' and sym in all_trades:
            t = all_trades[sym]
            t['win'] = 'WIN' if row['pnl_pct'] > 0 else 'LOSS'
            t['exit_date'] = row['date']
            t['exit_price'] = row['exit_price']
            t['pnl_pct'] = row['pnl_pct']
            t['pnl_dollar'] = row['pnl_dollar']
            t['reason'] = row['reason']
            final_pnl += row['pnl_dollar']
    return list(all_trades.values()), final_pnl


def main():
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()
    (OUT / 'data').mkdir()
    (OUT / 'charts' / 'v7').mkdir(parents=True)
    (OUT / 'charts' / 'v9').mkdir(parents=True)

    # 1. 讀 16y trades
    v7 = load_trades(EXCEL_V7, 'v7')
    v9 = load_trades(EXCEL_V9, 'v9')
    (OUT / 'data' / 'v7.json').write_text(json.dumps(v7, default=str))
    (OUT / 'data' / 'v9.json').write_text(json.dumps(v9, default=str))
    print(f'V7: {len(v7)} trades, V9: {len(v9)} trades')

    # 2. 讀 daily trades (PENDING + 完成的)
    daily, daily_pnl = load_daily_trades(CSV_DAILY)
    (OUT / 'data' / 'daily.json').write_text(json.dumps(daily, default=str))
    print(f'Daily: {len(daily)} trades (pending + completed)')

    # 3. 複製 PNG
    n7 = sum(1 for _ in CHARTS_V7.glob('*.png'))
    n9 = sum(1 for _ in CHARTS_V9.glob('*.png'))
    for f in CHARTS_V7.glob('*.png'):
        shutil.copy(f, OUT / 'charts' / 'v7' / f.name)
    for f in CHARTS_V9.glob('*.png'):
        shutil.copy(f, OUT / 'charts' / 'v9' / f.name)
    print(f'PNG: V7={n7}, V9={n9}')

    # 4. 寫 index.html (3 tabs: V7 / V9 / Daily)
    html = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<title>V7 / V9 paper trade (GitHub Pages)</title>
<style>
  body { font-family: 'Microsoft JhengHei', Arial, sans-serif; margin: 0; background: #f5f5f5; color: #333; }
  .top { background: #2c3e50; color: #fff; padding: 12px 16px; display: flex; flex-direction: column; gap: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
  .top-main { display: flex; justify-content: space-between; align-items: center; }
  .top h1 { margin: 0; font-size: 18px; }
  .tabs { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 4px; }
  .tab { padding: 6px 12px; border: 1px solid #fff; border-radius: 4px; cursor: pointer; font-size: 13px; white-space: nowrap; }
  .tab.active { background: #3498db; border-color: #3498db; }
  .stats { display: flex; gap: 12px; font-size: 12px; overflow-x: auto; padding-bottom: 4px; }
  .stat { display: flex; flex-direction: column; align-items: center; min-width: 60px; }
  .stat-label { opacity: 0.7; font-size: 10px; }
  .stat-value { font-weight: bold; font-size: 14px; margin-top: 2px; }
  .stat-value.win { color: #2ecc71; }
  .stat-value.loss { color: #e74c3c; }
  .strategy { background: #fff3cd; color: #856404; padding: 10px 16px; border-bottom: 1px solid #ffeeba; font-size: 12px; line-height: 1.6; }
  .strategy b { color: #533f03; }
  .strategy summary { cursor: pointer; font-weight: bold; color: #533f03; }
  .strategy code { background: rgba(0,0,0,0.05); padding: 1px 4px; border-radius: 3px; font-size: 11px; }
  .container { display: flex; flex-direction: column; height: calc(100vh - 120px); }
  .left { width: 100%; overflow-y: auto; background: #fff; border-right: 1px solid #ddd; }
  .filter { padding: 10px 16px; background: #ecf0f1; border-bottom: 1px solid #ddd; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  .filter label { font-size: 12px; margin-right: 4px; }
  .filter select, .filter input { padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 12px; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; min-width: 600px; }
  thead { background: #34495e; color: #fff; position: sticky; top: 0; z-index: 10; }
  th, td { padding: 8px 6px; text-align: left; border-bottom: 1px solid #ecf0f1; }
  th { font-weight: 600; }
  tr:hover { background: #f8f9fa; cursor: pointer; }
  thead tr:hover { background: #34495e; cursor: default; }
  tr.selected { background: #d5e8f7 !important; }
  .win-bg { background: #d5f5e3; }
  .loss-bg { background: #fadbd8; }
  .pending-bg { background: #fcfcfc; }
  .right { width: 100%; display: none; flex-direction: column; background: #2c3e50; }
  .right.show { display: flex; position: fixed; top: 0; right: 0; width: 100%; height: 100%; z-index: 1000; }
  .right-header { background: #34495e; color: #fff; padding: 10px 16px; }
  .right-header h3 { margin: 0 0 4px 0; font-size: 15px; }
  .right-header .detail { font-size: 11px; opacity: 0.9; line-height: 1.5; }
  .close-btn { position: absolute; top: 8px; right: 12px; color: #fff; font-size: 28px; cursor: pointer; }
  .chart-area { flex: 1; display: flex; align-items: center; justify-content: center; padding: 12px; overflow: auto; }
  .chart-area img { max-width: 100%; max-height: 100%; cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3); }
  .empty { color: #95a5a6; font-size: 14px; text-align: center; padding: 60px 20px; }
  .pos { color: #27ae60; font-weight: bold; }
  .neg { color: #c0392b; font-weight: bold; }
  .modal { display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); align-items: center; justify-content: center; }
  .modal.show { display: flex; }
  .modal img { max-width: 95%; max-height: 95%; }
  .modal-close { position: absolute; top: 20px; right: 30px; color: #fff; font-size: 36px; cursor: pointer; }
  .updated-at { text-align: center; padding: 8px; background: #ecf0f1; font-size: 11px; color: #7f8c8d; }
</style>
</head>
<body>
<div class="top">
  <div class="top-main">
    <h1>📊 V7 / V9 paper trade</h1>
  </div>
  <div class="tabs">
    <div class="tab active" onclick="switchTab('v7')" id="tab-v7">V7 (16y)</div>
    <div class="tab" onclick="switchTab('v9')" id="tab-v9">V9 (16y)</div>
    <div class="tab" onclick="switchTab('daily')" id="tab-daily">V7 Daily</div>
  </div>
  <div class="stats" id="stats">
    <div class="stat"><div class="stat-label">期末</div><div class="stat-value win" id="stat-final"></div></div>
    <div class="stat"><div class="stat-label">報酬</div><div class="stat-value win" id="stat-ret"></div></div>
    <div class="stat"><div class="stat-label">trades</div><div class="stat-value" id="stat-total"></div></div>
    <div class="stat"><div class="stat-label">勝率</div><div class="stat-value" id="stat-wr"></div></div>
    <div class="stat"><div class="stat-label">avg hold</div><div class="stat-value" id="stat-hold"></div></div>
  </div>
</div>

<div class="strategy">
  <details>
    <summary>📋 策略說明 (V7 / V9 martingale)</summary>
    <div style="margin-top:6px;">
      <b>V7 (16y)</b>: 0050 限定. 進場後每跌 <code>10%</code> 加碼 1 張, 最多 3 張 (跌 20% / 30% / 40% 各加 1 張),
      跌 <code>50%</code> 停損. 漲 <code>+20%</code> 出場. <b>16y 報酬 +732.54%</b>, 勝率 62.91% (302 trades).
    </div>
    <div style="margin-top:4px;">
      <b>V9 (16y)</b>: V7 基礎 + 跌破 5MA 提前出場. 16y 報酬 <b>+129%</b> (5.5y 區間).
      V9 表現 <span style="color:#c00;font-weight:bold;">輸 V7</span>, 因為 5MA 出場規則錯失波段.
    </div>
    <div style="margin-top:4px;">
      <b>Daily (V7 8/3~)</b>: 真實即時跑 V7 規則. <b>GitHub Pages 每天 06:00 / 18:00 更新</b>, PENDING 出場價 = N/A
      (即時價格需本地 FastAPI 版, 看 <code>localhost:8768</code>).
    </div>
  </details>
</div>

<div class="container">
  <div class="left">
    <div class="filter">
      <label>結果:</label>
      <select id="fResult"><option value="all">全部</option><option value="WIN">WIN</option><option value="LOSS">LOSS</option><option value="PENDING">持倉中</option></select>
      <label>理由:</label>
      <select id="fReason"><option value="all">全部</option><option value="tp">tp</option><option value="sl">sl</option><option value="ma5_break">ma5_break</option><option value="end_of_test">end</option><option value="Holding">Holding</option></select>
      <label>加碼:</label>
      <select id="fAdd"><option value="all">全部</option><option value="1">1 張</option><option value="2">2 張</option><option value="3">3 張</option></select>
      <label>排序:</label>
      <select id="fSort">
        <option value="entry_desc">進場日 (新→舊)</option>
        <option value="entry_asc">進場日 (舊→新)</option>
        <option value="pnl_desc">報酬 (高→低)</option>
        <option value="pnl_asc">報酬 (低→高)</option>
      </select>
      <input type="text" id="fSym" placeholder="搜尋股票代碼" style="margin-left:auto;width:80px;">
    </div>
    <table id="tradeTable">
      <thead><tr><th>#</th><th>結果</th><th>代碼</th><th>進場</th><th>出場</th><th>成本</th><th>張數</th><th>持倉</th><th>出場價</th><th>損益</th><th>理由</th></tr></thead>
      <tbody></tbody>
    </table>
  </div>
  <div class="right" id="rightPanel">
    <div class="right-header">
      <h3 id="chartTitle">點擊左側任一筆查看 K 線圖</h3>
      <div class="detail" id="chartDetail"></div>
      <div class="close-btn" onclick="closeRightPanel()">&times;</div>
    </div>
    <div class="chart-area" id="chartArea"><div class="empty">← 從左側選一筆 trade</div></div>
  </div>
</div>

<div class="updated-at" id="updatedAt"></div>

<div class="modal" id="modal" onclick="closeModal()">
  <span class="modal-close">&times;</span>
  <img id="modalImg" src="">
</div>

<script>
const data = { v7: null, v9: null, daily: null };
let current = 'v7';
let allRows = [];

function calcStats(trades) {
  const wins = trades.filter(t => t.win === 'WIN').length;
  const completed = trades.filter(t => t.win !== 'PENDING');
  const total = completed.length;
  const final_cap = 2000000 + completed.reduce((s, t) => s + (t.pnl_dollar || 0), 0);
  const ret_pct = (final_cap - 2000000) / 2000000 * 100;
  const avg_hold = total > 0 ? completed.reduce((s, t) => s + (t.hold_days || 0), 0) / total : 0;
  const pending = trades.length - completed.length;
  return {
    final_cap: final_cap.toLocaleString(),
    ret: (ret_pct >= 0 ? '+' : '') + ret_pct.toFixed(2) + '%',
    total: trades.length + (pending > 0 ? ' (' + pending + ' 持倉)' : ''),
    wr: total > 0 ? (wins/total*100).toFixed(2) : '—',
    avg_hold: total > 0 ? avg_hold.toFixed(1) : '—',
  };
}

function renderStats(key) {
  const s = calcStats(data[key] || []);
  document.getElementById('stat-final').textContent = s.final_cap;
  document.getElementById('stat-ret').textContent = s.ret;
  document.getElementById('stat-total').textContent = s.total;
  document.getElementById('stat-wr').textContent = s.wr + '%';
  document.getElementById('stat-hold').textContent = s.avg_hold + '天';
}

function loadRows() {
  const trades = data[current] || [];
  allRows = trades.map(t => {
    const tr = document.createElement('tr');
    tr.dataset.idx = t.idx;
    tr.dataset.sym = t.sym;
    tr.dataset.reason = t.reason;
    tr.dataset.win = t.win;
    tr.dataset.buy = t.buy_count;
    tr.dataset.pnl = t.pnl_pct;
    tr.dataset.entry = t.entry_date;
    tr.dataset.png = t.png_url;
    tr.dataset.detail = t.detail || '';
    tr.className = t.win === 'WIN' ? 'win-bg' : (t.win === 'LOSS' ? 'loss-bg' : 'pending-bg');
    const isPending = t.win === 'PENDING';
    const pnl = (t.pnl_pct || 0) * 100;
    const pnlClass = isPending ? 'pending' : (t.pnl_pct > 0 ? 'pos' : 'neg');
    const exitPriceDisplay = isPending ? 'N/A' : (+t.exit_price).toFixed(2);
    tr.innerHTML = '<td>' + t.idx + '</td><td>' + t.win + '</td><td><b>' + t.sym + '</b></td><td>' + t.entry_date + '</td><td>' + t.exit_date + '</td><td>' + (+t.first_cost).toFixed(2) + '</td><td>' + t.buy_count + '</td><td>' + t.hold_days + '天</td><td>' + exitPriceDisplay + '</td><td class="' + pnlClass + '">' + (isPending ? '⏳' : (pnl >= 0 ? '+' : '') + pnl.toFixed(2) + '%') + '</td><td>' + t.reason + '</td>';
    return tr;
  });
  applyFilter();
}

function applyFilter() {
  const fr = document.getElementById('fResult').value;
  const fr2 = document.getElementById('fReason').value;
  const fa = document.getElementById('fAdd').value;
  const fs = document.getElementById('fSort').value;
  const fsym = document.getElementById('fSym').value.trim();

  let rows = allRows.slice();
  if (fr !== 'all') rows = rows.filter(r => r.dataset.win === fr);
  if (fr2 !== 'all') rows = rows.filter(r => r.dataset.reason === fr2);
  if (fa !== 'all') rows = rows.filter(r => r.dataset.buy === fa);
  if (fsym) rows = rows.filter(r => r.dataset.sym.includes(fsym));

  if (fs === 'entry_desc') rows.sort((a,b) => b.dataset.entry.localeCompare(a.dataset.entry));
  else if (fs === 'entry_asc') rows.sort((a,b) => a.dataset.entry.localeCompare(b.dataset.entry));
  else if (fs === 'pnl_desc') rows.sort((a,b) => parseFloat(b.dataset.pnl) - parseFloat(a.dataset.pnl));
  else if (fs === 'pnl_asc') rows.sort((a,b) => parseFloat(a.dataset.pnl) - parseFloat(b.dataset.pnl));

  const tbody = document.querySelector('#tradeTable tbody');
  tbody.innerHTML = '';
  rows.forEach(r => {
    r.onclick = () => showTrade(r);
    tbody.appendChild(r);
  });
}

['fResult','fReason','fAdd','fSort'].forEach(id => {
  document.getElementById(id).addEventListener('change', applyFilter);
});
document.getElementById('fSym').addEventListener('input', applyFilter);

function showTrade(row) {
  allRows.forEach(r => r.classList.remove('selected'));
  row.classList.add('selected');
  document.getElementById('chartTitle').textContent = '#' + row.dataset.idx + ' ' + row.dataset.sym + ' (' + row.dataset.entry + ', ' + row.dataset.reason + ')';
  document.getElementById('chartDetail').textContent = row.dataset.detail || '';
  const png = row.dataset.png;
  if (!png) {
    document.getElementById('chartArea').innerHTML = '<div class="empty">PENDING (持倉中) — 無出場 K 線圖<br><br>GitHub Pages 為純靜態. 即時 K 線要看本地 FastAPI: <code>localhost:8768</code></div>';
  } else {
    document.getElementById('chartArea').innerHTML = '<img src="' + png + '" onclick="openModal(this.src)" onerror="this.outerHTML=\\'<div class=&quot;empty&quot;>找不到圖</div>\\'">';
  }
  document.getElementById('rightPanel').classList.add('show');
}

function closeRightPanel() {
  document.getElementById('rightPanel').classList.remove('show');
}

function switchTab(key) {
  current = key;
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('tab-' + key).classList.add('active');
  renderStats(key);
  loadRows();
  closeRightPanel();
}

function openModal(url) {
  document.getElementById('modalImg').src = url;
  document.getElementById('modal').classList.add('show');
}
function closeModal() {
  document.getElementById('modal').classList.remove('show');
}
document.addEventListener('keydown', e => { if (e.key === 'Escape') { closeModal(); closeRightPanel(); } });

// 載入 JSON
Promise.all([
  fetch('data/v7.json').then(r => r.json()),
  fetch('data/v9.json').then(r => r.json()),
  fetch('data/daily.json').then(r => r.json()).catch(() => [])
]).then(([v7, v9, daily]) => {
  data.v7 = v7;
  data.v9 = v9;
  data.daily = daily || [];
  renderStats('v7');
  loadRows();
  document.getElementById('updatedAt').textContent = '📅 最後更新: ' + new Date().toISOString().replace('T', ' ').slice(0, 16) + ' UTC (GitHub Pages 每 12 小時 rebuild)';
});
</script>
</body>
</html>
"""
    (OUT / 'index.html').write_text(html)
    print(f'index.html: {len(html)} chars')

    print(f'\n靜態檔產出: {OUT}')
    print('總檔案數:')
    for sub in ['data', 'charts/v7', 'charts/v9']:
        files = list((OUT / sub).glob('*'))
        print(f'  {sub}: {len(files)} files')

    print('\n[DONE] 接下來可以:')
    print(f'  1) cd {OUT}')
    print(f'  2) 推到 GitHub Pages repo 的 data/ + index.html + charts/')


if __name__ == '__main__':
    main()
