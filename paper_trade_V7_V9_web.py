"""V7/V9 paper trade 網頁版 (16y + Daily) - RWD Optimized"""

from fastapi import FastAPI, Request, Query
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import openpyxl
from pathlib import Path
from datetime import datetime
from jinja2 import Template
import pandas as pd

ROOT = Path(__file__).parent
EXCEL_V7 = ROOT / 'paper_trade_V7_16y_trades.xlsx'
CHARTS_V7 = ROOT / 'paper_trade_V7_16y_charts'
EXCEL_V9 = ROOT / 'paper_trade_V9_16y_trades.xlsx'
CHARTS_V9 = ROOT / 'paper_trade_V9_16y_charts'
CSV_V7_DAILY = ROOT / 'paper_trade_V7_trades.csv'

app = FastAPI()
app.mount('/charts_v7', StaticFiles(directory=str(CHARTS_V7)), name='charts_v7')
app.mount('/charts_v9', StaticFiles(directory=str(CHARTS_V9)), name='charts_v9')


def load_trades(excel_path, charts_dir, label):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    trades = []
    total_pnl = 0
    for r_idx in range(5, ws.max_row + 1):
        sym = ws.cell(row=r_idx, column=3).value
        if not (sym and isinstance(sym, str) and sym.isdigit() and len(sym) == 4):
            continue
        pnl_dollar = ws.cell(row=r_idx, column=12).value or 0
        total_pnl += pnl_dollar
        idx = ws.cell(row=r_idx, column=1).value
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
            'pnl_dollar': pnl_dollar,
            'reason': ws.cell(row=r_idx, column=13).value,
            'detail': ws.cell(row=r_idx, column=14).value,
            'png_url': f'/charts_{label}/{idx:03d}_{sym}_{ws.cell(row=r_idx, column=4).value}_{ws.cell(row=r_idx, column=13).value}.png',
        })
    return trades, total_pnl

def load_trades_daily(csv_path):
    if not csv_path.exists():
        return [], 0
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
                'png_url': ''
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


INDEX_HTML = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>V7 / V9 paper trade 16y & Daily</title>
<style>
  body { font-family: 'Microsoft JhengHei', Arial, sans-serif; margin: 0; background: #f5f5f5; color: #333; }
  .top { background: #2c3e50; color: #fff; padding: 12px 16px; display: flex; flex-direction: column; gap: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
  .top-main { display: flex; justify-content: space-between; align-items: center; }
  .top h1 { margin: 0; font-size: 18px; }
  .tabs { display: flex; gap: 8px; overflow-x: auto; padding-bottom: 4px; -webkit-overflow-scrolling: touch; }
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
  .strategy details { margin-top: 4px; }
  .strategy summary { cursor: pointer; font-weight: bold; color: #533f03; }
  .strategy code { background: rgba(0,0,0,0.05); padding: 1px 4px; border-radius: 3px; font-size: 11px; }
  .container { display: flex; flex-direction: column; height: calc(100vh - 120px); }
  .left { width: 100%; overflow-y: auto; background: #fff; border-right: 1px solid #ddd; }
  .filter { padding: 10px 16px; background: #ecf0f1; border-bottom: 1px solid #ddd; display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  .filter label { font-size: 12px; margin-right: 4px; }
  .filter select, .filter input { padding: 4px 8px; border: 1px solid #ccc; border-radius: 3px; font-size: 12px; }
  .table-wrapper { overflow-x: auto; }
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
  .right-header { background: #34495e; color: #fff; padding: 10px 16px; display: flex; justify-content: space-between; align-items: center; }
  .right-header h3 { margin: 0; font-size: 15px; }
  .close-btn { cursor: pointer; font-size: 24px; font-weight: bold; padding: 0 10px; }
  .right-header .detail { font-size: 11px; opacity: 0.9; line-height: 1.5; margin-top: 4px; }
  .chart-area { flex: 1; display: flex; align-items: center; justify-content: center; padding: 12px; overflow: auto; }
  .chart-area img { max-width: 100%; max-height: 100%; cursor: pointer; box-shadow: 0 4px 12px rgba(0,0,0,0.3); }
  .empty { color: #95a5a6; font-size: 14px; text-align: center; padding: 60px 20px; }
  .pos { color: #27ae60; font-weight: bold; }
  .neg { color: #c0392b; font-weight: bold; }
  .modal { display: none; position: fixed; z-index: 2000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); align-items: center; justify-content: center; }
  .modal.show { display: flex; }
  .modal img { max-width: 95%; max-height: 95%; }
  .modal-close { position: absolute; top: 20px; right: 30px; color: #fff; font-size: 36px; cursor: pointer; }
</style>
</head>
<body>
<div class="top">
  <div class="top-main">
    <h1>📊 Paper trade 歷史記錄</h1>
  </div>
  <div class="tabs">
    <div class="tab active" onclick="switchTab('v7')" id="tab-v7">V7 (16y)</div>
    <div class="tab" onclick="switchTab('v9')" id="tab-v9">V9 (16y)</div>
    <div class="tab" onclick="switchTab('v7_daily')" id="tab-v7_daily">V7 (Daily 8/3~)</div>
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
      <b>Daily (V7 8/3~)</b>: 真實即時跑 V7 規則, 持倉中會顯示未實現損益 (即時抓 yfinance 最新價).
    </div>
    <div style="margin-top:4px; font-size:11px; opacity:0.8;">
      ⚠️ <b>未實現損益</b> 用 yfinance 最新收盤價計算, 僅供參考. 持倉中的 K 線圖會顯示進場日 ~ 現在.
    </div>
  </details>
</div>

<div class="container">
  <div class="left">
    <div class="filter">
      <label>結果:</label>
      <select id="fResult"><option value="all">全部</option><option value="WIN">WIN</option><option value="LOSS">LOSS</option><option value="PENDING">持倉中</option></select>
      <label>出場:</label>
      <select id="fReason"><option value="all">全部</option><option value="tp">tp</option><option value="sl">sl</option><option value="ma5_break">ma5_break</option><option value="end_of_test">end</option></select>
      <label>加碼:</label>
      <select id="fAdd"><option value="all">全部</option><option value="1">1 張</option><option value="2">2 張</option><option value="3">3 張</option></select>
      <label>排序:</label>
      <select id="fSort">
        <option value="entry_desc">進場日 (新→舊)</option>
        <option value="entry_asc">進場日 (舊→新)</option>
        <option value="pnl_desc">報酬 (高→低)</option>
        <option value="pnl_asc">報酬 (低→高)</option>
      </select>
      <input type="text" id="fSym" placeholder="搜尋代碼" style="margin-left:auto;width:80px;">
    </div>
    <div class="table-wrapper">
      <table id="tradeTable">
        <thead><tr><th>#</th><th>結果</th><th>代碼</th><th>進場</th><th>出場</th><th>成本</th><th>張數</th><th>持倉</th><th>出場價</th><th>現價</th><th>損益</th><th>理由</th></tr></thead>
        <tbody></tbody>
      </table\>
    </div>
  </div>
  <div class="right" id="rightPanel">
    <div class="right-header">
      <div>
        <h3 id="chartTitle">K 線圖</h3>
        <div class="detail" id="chartDetail"></div>
      </div>
      <div class="close-btn" onclick="closeRightPanel()">&times;</div>
    </div>
    <div class="chart-area" id="chartArea"><div class="empty">← 從左側選一筆 trade</div></div>
  </div>
</div>

<div class="modal" id="modal" onclick="closeModal()">
  <span class="modal-close">&times;</span>
  <img id="modalImg" src="">
</div>

<script>
const data = {
  v7: { trades: {{ v7_trades|safe }}, stats: {{ v7_stats|safe }} },
  v9: { trades: {{ v9_trades|safe }}, stats: {{ v9_stats|safe }} },
  v7_daily: { trades: {{ v7_daily_trades|safe }}, stats: {{ v7_daily_stats|safe }} },
};
let current = 'v7';
let allRows = [];

function renderStats(key) {
  const s = data[key].stats;
  document.getElementById('stat-final').textContent = s.final_cap;
  document.getElementById('stat-ret').textContent = s.ret;
  document.getElementById('stat-total').textContent = s.total;
  document.getElementById('stat-wr').textContent = s.wr + '%';
  document.getElementById('stat-hold').textContent = s.avg_hold + '天';
}

function loadRows() {
  allRows = data[current].trades.map(t => {
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
    const exitPriceDisplay = isPending ? 'N/A' : (+t.exit_price).toFixed(2);
    tr.innerHTML = `<td>${t.idx}</td><td>${t.win}</td><td><b>${t.sym}</b></td><td>${t.entry_date}</td><td>${t.exit_date}</td><td>${(+t.first_cost).toFixed(2)}</td><td>${t.buy_count}</td><td>${t.hold_days}天</td><td id="exit-${t.idx}" class="exit-price">${exitPriceDisplay}</td><td id="now-${t.idx}" class="now-price">${isPending ? '<span style="opacity:0.5;">⏳</span>' : '—'}</td><td id="pnl-${t.idx}" class="pnl-cell ${isPending ? 'pending' : (t.pnl_pct > 0 ? 'pos' : 'neg')}">${isPending ? '<i>⏳</i>' : (t.pnl_pct*100).toFixed(2)+'%'}</td><td>${t.reason}</td>`;
    return tr;
  });
  applyFilter();
  // 對 PENDING 持倉即時抓未實現損益 + 現價
  fetchPendingUnrealized();
}

function fetchPendingUnrealized() {
  // 找出所有 PENDING row (用 win filter 後的 tbody)
  const pendingRows = allRows.filter(r => r.dataset.win === 'PENDING');
  if (pendingRows.length === 0) return;

  // 並行 fetch (每個 PENDING 一個 holding endpoint)
  pendingRows.forEach(row => {
    const sym = row.dataset.sym;
    const idx = row.dataset.idx;
    fetch(`/api/holding/${sym}`)
      .then(r => {
        if (!r.ok) return null;
        const latestPrice = r.headers.get('X-Latest-Price');
        const unrealizedPct = r.headers.get('X-Unrealized-Pct');
        if (latestPrice && unrealizedPct !== null) {
          const pct = parseFloat(unrealizedPct);
          const nowCell = document.getElementById(`now-${idx}`);
          if (nowCell) {
            nowCell.innerHTML = `<b style="color:#0066cc;">${parseFloat(latestPrice).toFixed(2)}</b>`;
            nowCell.title = '即時收盤價 (yfinance)';
          }
          const pnlCell = document.getElementById(`pnl-${idx}`);
          if (pnlCell) {
            pnlCell.innerHTML = `<i>${pct.toFixed(2)}%</i>`;
            pnlCell.className = `pnl-cell ${pct >= 0 ? 'pos' : 'neg'}`;
            pnlCell.title = '未實現損益 (即時)';
          }
          row.dataset.pnl = (pct / 100).toString();
        }
      })
      .catch(err => {
        const nowCell = document.getElementById(`now-${idx}`);
        if (nowCell) nowCell.innerHTML = `<span style="color:#c00;" title="error: ${err.message}">N/A</span>`;
      });
  });
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

  document.getElementById('chartTitle').textContent = `#${row.dataset.idx} ${row.dataset.sym}`;
  document.getElementById('chartDetail').textContent = `${row.dataset.entry}, ${row.dataset.reason}`;

  // PENDING (持倉中) → 即時畫 K 線 + 未實現損益
  if (row.dataset.win === 'PENDING') {
    const url = `/api/holding/${row.dataset.sym}`;
    document.getElementById('chartArea').innerHTML =
      `<div class="empty" id="holding-loading">⏳ 抓即時資料中...</div>`;
    fetch(url)
      .then(r => {
        const entryCost = r.headers.get('X-Entry-Cost');
        const latestPrice = r.headers.get('X-Latest-Price');
        const unrealizedPct = r.headers.get('X-Unrealized-Pct');
        const entryDate = r.headers.get('X-Entry-Date');
        if (!r.ok) {
          return r.json().then(err => { throw new Error(err.error || 'fetch failed'); });
        }
        // 更新 detail 區顯示未實現損益
        document.getElementById('chartDetail').innerHTML =
          `進場 ${entryDate} @ ${parseFloat(entryCost).toFixed(2)} | 現價 ${parseFloat(latestPrice).toFixed(2)} | 未實現 <b style="color:${parseFloat(unrealizedPct)>=0?'#2ecc71':'#e74c3c'}">${parseFloat(unrealizedPct).toFixed(2)}%</b>`;
        return r.blob();
      })
      .then(blob => {
        const imgUrl = URL.createObjectURL(blob);
        document.getElementById('chartArea').innerHTML =
          `<img src="${imgUrl}" onclick="openModalUrl('${imgUrl}')">`;
      })
      .catch(err => {
        document.getElementById('chartArea').innerHTML =
          `<div class="empty">❌ 抓取失敗: ${err.message}<br><br>可能原因: 非台股代碼 / yfinance 沒有資料 / 網路問題</div>`;
      });
  } else {
    // WIN/LOSS → 用預存的 png
    const png = row.dataset.png;
    if (!png) {
      document.getElementById('chartArea').innerHTML = `<div class="empty">此筆無預存 K 線圖</div>`;
    } else {
      document.getElementById('chartArea').innerHTML =
        `<img src="${png}" onclick="openModal('${png}')" onerror="this.outerHTML='<div class=\\\\'empty\\\\'>找不到圖</div>'">`;
    }
  }

  document.getElementById('rightPanel').classList.add('show');
}

function openModalUrl(url) {
  document.getElementById('modalImg').src = url;
  document.getElementById('modal').classList.add('show');
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
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

renderStats('v7');
loadRows();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
async def index():
    v7_trades, v7_pnl = load_trades(EXCEL_V7, CHARTS_V7, 'v7')
    v9_trades, v9_pnl = load_trades(EXCEL_V9, CHARTS_V9, 'v9')
    v7_daily_trades, v7_daily_pnl = load_trades_daily(CSV_V7_DAILY)

    def stats(trades, pnl):
        wins = sum(1 for t in trades if t['win'] == 'WIN')
        total = len(trades)
        final_cap = 2_000_000 + pnl
        ret_pct = (final_cap - 2_000_000) / 2_000_000 * 100
        avg_hold = sum(t['hold_days'] for t in trades) / total if total > 0 else 0
        return {
            'final_cap': f'{final_cap:,.0f}',
            'ret': f'{ret_pct:+.2f}%',
            'total': total,
            'wr': f'{wins/total*100:.2f}' if total > 0 else '0',
            'avg_hold': f'{avg_hold:.1f}',
        }

    v7_stats = stats(v7_trades, v7_pnl)
    v9_stats = stats(v9_trades, v9_pnl)
    v7_daily_stats = stats(v7_daily_trades, v7_daily_pnl)

    import json
    template = Template(INDEX_HTML)
    html = template.render(
        v7_trades=json.dumps(v7_trades, default=str),
        v9_trades=json.dumps(v9_trades, default=str),
        v7_daily_trades=json.dumps(v7_daily_trades, default=str),
        v7_stats=json.dumps(v7_stats),
        v9_stats=json.dumps(v9_stats),
        v7_daily_stats=json.dumps(v7_daily_stats),
    )
    return HTMLResponse(html)


@app.get("/api/holding/{sym}")
async def holding_chart(sym: str):
    """PENDING 持倉中, 即時抓 yfinance 畫 K 線 + 算未實現損益"""
    import yfinance as yf
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from io import BytesIO
    import json as _json

    TW_SUFFIX = '.TW'
    symbol = sym if sym.endswith(TW_SUFFIX) else f'{sym}{TW_SUFFIX}'

    # 1. 抓 60 天資料
    try:
        df = yf.download(symbol, period='60d', progress=False)
        if df.empty or len(df) < 5:
            return {'error': f'抓不到 {symbol} 資料'}
    except Exception as e:
        return {'error': f'yfinance error: {e}'}

    # 2. 取收盤 (處理 multi-index 欄位)
    close = df['Close']
    if hasattr(close, 'columns'):
        close = close.iloc[:, 0]
    latest_price = float(close.iloc[-1])
    prev_close = float(close.iloc[-2]) if len(close) > 1 else latest_price
    daily_chg = (latest_price - prev_close) / prev_close * 100

    # 3. 找這筆 PENDING 的 entry_cost (從 daily CSV)
    if not CSV_V7_DAILY.exists():
        return {'error': 'CSV_V7_DAILY not found'}
    df_csv = pd.read_csv(CSV_V7_DAILY)
    # sym 可能是 int 或 str, 都試
    df_csv['sym_str'] = df_csv['sym'].astype(str)
    entries = df_csv[(df_csv['sym_str'] == sym) & (df_csv['action'] == 'entry')]
    if entries.empty:
        # 也試 str → int
        try:
            sym_int = int(sym)
            entries = df_csv[(df_csv['sym'] == sym_int) & (df_csv['action'] == 'entry')]
        except (ValueError, TypeError):
            pass
    if entries.empty:
        return {'error': f'找不到 {sym} entry record'}

    entry_row = entries.iloc[-1]
    entry_cost = float(entry_row['first_cost'])
    entry_date = str(entry_row['date'])

    # 4. 算未實現損益
    unrealized_pct = (latest_price - entry_cost) / entry_cost * 100

    # 5. 畫 K 線 (簡化版: line chart + 標記進場價)
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial Unicode MS']
    plt.rcParams['axes.unicode_minus'] = False

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(close.index, close.values, color='#1f77b4', linewidth=1.4, label=f'{symbol} 收盤')

    # 標記進場價
    ax.axhline(y=entry_cost, color='gray', linestyle='--', alpha=0.7, label=f'進場價 {entry_cost:.2f}')

    # 標記最新價
    last_date = close.index[-1]
    ax.scatter([last_date], [latest_price], color='red', s=60, zorder=5)
    ax.annotate(
        f'現價 {latest_price:.2f} ({daily_chg:+.2f}%)\n損益 {unrealized_pct:+.2f}%',
        xy=(last_date, latest_price), xytext=(10, 15), textcoords='offset points',
        fontsize=10, color='red', fontweight='bold',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='white', edgecolor='red', alpha=0.9)
    )

    pnl_color = '#2ecc71' if unrealized_pct >= 0 else '#e74c3c'
    ax.set_title(f'{symbol} 持倉中 (進場 {entry_date} @ {entry_cost:.2f}) — 未實現 {unrealized_pct:+.2f}%',
                 fontsize=12, color=pnl_color, fontweight='bold')
    ax.set_ylabel('價格')
    ax.legend(loc='upper left', fontsize=9)
    ax.grid(alpha=0.3)

    # 6. 輸出 PNG bytes
    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format='png', dpi=110, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)

    from fastapi.responses import Response
    return Response(
        content=buf.getvalue(),
        media_type='image/png',
        headers={
            'X-Entry-Cost': str(entry_cost),
            'X-Latest-Price': str(latest_price),
            'X-Unrealized-Pct': str(unrealized_pct),
            'X-Entry-Date': entry_date,
        }
    )


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='0.0.0.0', port=8768)
