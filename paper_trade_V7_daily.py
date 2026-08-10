"""
V7 波段 paper trade - 每日 18:15 跑
- 用 V7 邏輯 (跌 20% 進場, 加碼, 出場)
- 推 Telegram 群組 -1003990238955
- 持久化: state.json (持倉), trades.csv (成交記錄)
"""

import sys
import json
import sqlite3
import csv
import os
import asyncio
from pathlib import Path
from datetime import datetime, timedelta

MARKET_DATA = Path.home() / "market_data"
DB = MARKET_DATA / "tw_ohlcv_listed.sqlite"
STATE_JSON = MARKET_DATA / "paper_trade_V7_state.json"
TRADES_CSV = MARKET_DATA / "paper_trade_V7_trades.csv"

TW50 = set()
with open(MARKET_DATA / 'tw50_constituents.txt') as f:
    for line in f:
        c = line.strip()
        if c:
            TW50.add(c)

BLACKLIST = {'2368'}
DIRTY_DATES = {'2026-04-20'}

STOCK_NAMES_PATH = MARKET_DATA / 'tw_stock_names.txt'
STOCK_NAMES = {}
if STOCK_NAMES_PATH.exists():
    with open(STOCK_NAMES_PATH, encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split('\t')
            if len(parts) == 2:
                STOCK_NAMES[parts[0]] = parts[1]

INITIAL_CAPITAL = 2_000_000
SHARES_PER_LOT = 1000
ENTRY_DROP_PCT = 20.0
ADD_DOWN_PCT = 10.0
ADD_UP_PCT = 5.0
MAX_LOTS = 3
ADD_DAYS_LIMIT = 14
TP_PCT = 20.0
STOP_LOSS_PCT = 50.0

CHAT_ID = "-1003990238955"  # 群組


def load_state():
    if STATE_JSON.exists():
        return json.loads(STATE_JSON.read_text())
    return {
        'capital': INITIAL_CAPITAL,
        'positions': {},  # sym -> {lots: [(date, price)], first_cost, high_20, start_date}
    }


def save_state(state):
    STATE_JSON.write_text(json.dumps(state, indent=2, default=str))


def load_ohlcv_today():
    """讀取最近 30 天的 OHLCV (含今日)"""
    c = sqlite3.connect(str(DB), timeout=60)
    end = datetime.now().strftime('%Y-%m-%d')
    start = (datetime.now() - timedelta(days=60)).strftime('%Y-%m-%d')
    rows = c.execute("""
        SELECT symbol, date, open, high, low, close
        FROM ohlcv_day
        WHERE market='TWSE' AND date BETWEEN ? AND ?
        ORDER BY symbol, date
    """, (start, end)).fetchall()
    c.close()
    data = {}
    for sym, d, o, h, l, cl in rows:
        data.setdefault(sym, []).append({
            'date': d, 'open': o, 'high': h, 'low': l, 'close': cl
        })
    return data


def get_latest_trading_date():
    c = sqlite3.connect(str(DB), timeout=60)
    r = c.execute("SELECT MAX(date) FROM ohlcv_day WHERE market='TWSE'").fetchone()
    c.close()
    return r[0]


def run_paper_trade(verbose=True):
    state = load_state()
    capital = state['capital']
    positions = state['positions']

    asof = get_latest_trading_date()
    data = load_ohlcv_today()

    msg_lines = [f"V7 波段 paper trade (asof={asof})", ""]
    msg_lines.append(f"本金: {capital:,.0f}")
    msg_lines.append(f"持倉: {len(positions)} 檔")
    msg_lines.append("")

    trades_done = []

    # 1. 檢查出場 (從現有持倉)
    to_close = []
    for sym, pos in positions.items():
        if sym not in data:
            continue
        daily = data[sym]
        today = next((d_ for d_ in daily if d_['date'] == asof), None)
        if not today:
            continue
        cur = today['close']
        # 計算 MA5
        if len(daily) >= 5:
            ma5 = sum(daily[-5]['close'] for _ in [0]) / 5 if daily[-5] else None
        # 取最近 5 根
        recent = [d_ for d_ in daily if d_['date'] <= asof][-5:]
        if len(recent) >= 5:
            ma5 = sum(d['close'] for d in recent) / 5
            # 出場 1: +20%
            if (cur - pos['first_cost']) / pos['first_cost'] * 100 >= TP_PCT:
                to_close.append((sym, 'tp', today, cur, pos))
                continue
            # 出場 3: -50%
            high20 = pos.get('high_20', pos['first_cost'])
            if high20 > 0 and (high20 - cur) / high20 * 100 >= STOP_LOSS_PCT:
                to_close.append((sym, 'sl', today, cur, pos))

    for sym, reason, today, cur, pos in to_close:
        total_cost = sum(p * SHARES_PER_LOT for _, p in pos['lots'])
        total_shares = len(pos['lots']) * SHARES_PER_LOT
        exit_price = cur
        pnl_dollar = exit_price * total_shares - total_cost
        pnl_pct = (exit_price - total_cost/total_shares) / (total_cost/total_shares) * 100
        capital += total_cost + pnl_dollar
        msg_lines.append(f"[出場] {sym} {pos['lots'][0][0]}→{asof} ({len(pos['lots'])}張)")
        msg_lines.append(f"  進場 {pos['first_cost']:.2f} → 賣 {exit_price:.2f}, {pnl_pct:+.2f}% ({pnl_dollar:+,.0f})")
        msg_lines.append(f"  理由: {reason}")
        trades_done.append({
            'date': asof, 'sym': sym, 'action': 'exit', 'reason': reason,
            'entry_date': pos['lots'][0][0], 'first_cost': pos['first_cost'],
            'exit_price': exit_price, 'lots': len(pos['lots']),
            'pnl_pct': pnl_pct, 'pnl_dollar': pnl_dollar,
        })
        positions.pop(sym)

    # 2. 進場 (V7: 跌 20%)
    bought_today = False
    candidates = []
    for sym in TW50:
        if sym in BLACKLIST or sym in positions or positions or sym not in data:
            continue
        daily = data[sym]
        today = next((d_ for d_ in daily if d_['date'] == asof), None)
        if not today:
            continue
        if len(daily) < 21:
            continue
        window = [d_ for d_ in daily[max(0, len(daily)-21):-1] if d_['date'] not in DIRTY_DATES]
        if len(window) < 5:
            continue
        high20 = max(d_['high'] for d_ in window)
        if high20 <= 0:
            continue
        drop_pct = (high20 - today['close']) / high20 * 100
        if drop_pct >= ENTRY_DROP_PCT:
            candidates.append((sym, drop_pct, today, high20))

    if candidates and not positions:  # 雙重保險: 沒持倉才進場
        candidates.sort(key=lambda x: -x[1])
        sym, drop_pct, today, high20 = candidates[0]
        cost = SHARES_PER_LOT * today['close']
        if cost <= capital:
            positions[sym] = {
                'lots': [(asof, today['close'])],
                'first_cost': today['close'],
                'high_20': high20,
                'start_date': asof,
            }
            capital -= cost
            bought_today = True
            msg_lines.append(f"[進場] {sym} {asof}")
            msg_lines.append(f"  跌 {drop_pct:.1f}% 從高 {high20:.2f} → 收 {today['close']:.2f}")
            msg_lines.append(f"  1 張 @ {today['close']:.2f} = {cost:,.0f}")
            trades_done.append({
                'date': asof, 'sym': sym, 'action': 'entry', 'reason': 'drop_20',
                'first_cost': today['close'], 'lots': 1, 'high_20': high20,
            })

    # 3. 加碼
    for sym, pos in list(positions.items()):
        if sym not in data:
            continue
        daily = data[sym]
        today = next((d_ for d_ in daily if d_['date'] == asof), None)
        if not today:
            continue
        from datetime import datetime as _dt
        days_since = (_dt.strptime(asof, '%Y-%m-%d') - _dt.strptime(pos['start_date'], '%Y-%m-%d')).days
        if days_since > ADD_DAYS_LIMIT or len(pos['lots']) >= MAX_LOTS:
            continue
        cur = today['close']
        pct_to_first = (cur - pos['first_cost']) / pos['first_cost'] * 100
        if pct_to_first <= -ADD_DOWN_PCT or pct_to_first >= ADD_UP_PCT:
            cost = SHARES_PER_LOT * cur
            if cost <= capital:
                pos['lots'].append((asof, cur))
                capital -= cost
                msg_lines.append(f"[加碼] {sym} +1 張 @ {cur:.2f} ({pct_to_first:+.1f}% 從首張)")
                trades_done.append({
                    'date': asof, 'sym': sym, 'action': 'add',
                    'first_cost': cur, 'lots': len(pos['lots']),
                })

    # 4. 持倉清單
    if positions:
        msg_lines.append("")
        msg_lines.append(f"持倉 ({len(positions)} 檔):  [股票名稱 | 買入日 | 進場價 | 今日收盤 | 損益]")
        for sym, pos in positions.items():
            daily = data.get(sym, [])
            today = next((d_ for d_ in daily if d_['date'] == asof), None)
            cur = today['close'] if today else pos['first_cost']
            pnl = (cur - pos['first_cost']) / pos['first_cost'] * 100
            lots = pos['lots']
            first_buy_date = lots[0][0] if lots else '?'
            last_buy_date = lots[-1][0] if len(lots) > 1 else ''
            if last_buy_date == '' or first_buy_date == last_buy_date:
                buy_date_str = first_buy_date
            else:
                buy_date_str = f"{first_buy_date}~{last_buy_date}"
            stock_name = STOCK_NAMES.get(sym, '')
            name_part = f"{stock_name} " if stock_name else ""
            msg_lines.append(f"  {sym} {name_part}{len(pos['lots'])}張 買入:{buy_date_str} 進 {pos['first_cost']:.2f} → 今 {cur:.2f} ({pnl:+.2f}%)")

    msg_lines.append("")
    msg_lines.append(f"現金: {capital:,.0f}")
    msg_lines.append(f"持倉成本: {sum(sum(p*SHARES_PER_LOT for _, p in pos['lots']) for pos in positions.values()):,.0f}")
    total_value = capital + sum(sum(p*SHARES_PER_LOT for _, p in pos['lots']) for pos in positions.values())
    msg_lines.append(f"總資產: {total_value:,.0f} ({(total_value-INITIAL_CAPITAL)/INITIAL_CAPITAL*100:+.2f}%)")

    # 存 state
    state['capital'] = capital
    state['positions'] = positions
    save_state(state)

    # 寫 trades CSV
    if trades_done:
        new_file = not TRADES_CSV.exists()
        with open(TRADES_CSV, 'a', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=['date','sym','action','reason','entry_date','first_cost','exit_price','lots','pnl_pct','pnl_dollar','high_20'])
            if new_file:
                w.writeheader()
            for t in trades_done:
                w.writerow(t)

    # 推 Telegram
    text = '\n'.join(msg_lines)
    if verbose:
        print(text)

    # 送到 Telegram
    try:
        from hermes_cli.send_cmd import _load_hermes_env
        _load_hermes_env()
        token = os.environ.get("TELEGRAM_BOT_TOKEN")
        if token:
            from telegram import Bot
            async def _send():
                bot = Bot(token=token)
                await bot.send_message(chat_id=CHAT_ID, text=text)
            asyncio.run(_send())
            print(f'已推 Telegram 群組 {CHAT_ID}')
        else:
            print('TELEGRAM_BOT_TOKEN 未設, 跳過推播')
    except Exception as e:
        print(f'Telegram 推播失敗: {e}')

    return text


if __name__ == '__main__':
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('--dry-run', action='store_true', help='不推 Telegram')
    a = p.parse_args()
    run_paper_trade(verbose=True)