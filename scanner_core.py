"""
================================================================================
 NSE INTRADAY MOMENTUM SCANNER v2.0 -- FULL RANKING + EXCEL EXPORT
================================================================================

WHAT THIS DOES
--------------
- Ranks ALL liquid NSE stocks (not just those passing hard gates)
- Shows Entry, Stop Loss, Target 1, Target 2 for every ranked stock
- Exports to Excel with formatting and multiple sheets
- Configurable --top N (default 20)
- Confidence scores with star ratings

USAGE
-----
    pip install yfinance pandas numpy openpyxl

    # Rank top 20 stocks
    python intraday_scanner.py --capital 100000 --leverage 5

    # Rank top 50 stocks
    python intraday_scanner.py --top 50 --capital 100000 --leverage 5

    # Rank all stocks
    python intraday_scanner.py --top 150 --capital 100000 --leverage 5
================================================================================
"""

from __future__ import annotations

import argparse
import math
import os
import sys
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

try:
    import yfinance as yf
except ImportError:
    sys.exit("pip install yfinance pandas numpy openpyxl")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
except ImportError:
    sys.exit("pip install openpyxl")

import logging
logging.getLogger("yfinance").setLevel(logging.CRITICAL)
logging.getLogger("urllib3").setLevel(logging.CRITICAL)

IST = "Asia/Kolkata"


# =============================================================================
# 1. CONFIG
# =============================================================================

@dataclass
class Config:
    capital: float = 100_000.0
    leverage: float = 5.0
    risk_pct: float = 0.5
    top_n: int = 20
    signal_time: str = "09:45"
    allow_shorts: bool = True
    min_turnover_cr: float = 100.0
    min_atr_pct: float = 0.8
    max_atr_pct: float = 8.0
    min_price: float = 30.0
    min_rvol: float = 0.0
    max_ext_atr: float = 999.0
    min_sl_pct: float = 0.3
    max_sl_pct: float = 1.5
    t1_r: float = 1.5
    t2_r: float = 2.5
    respect_market_regime: bool = False
    outdir: str = "scan_output"
    no_save: bool = False


# =============================================================================
# 2. UNIVERSE
# =============================================================================

UNIVERSE: Dict[str, str] = {
    # Banks
    "HDFCBANK": "Private Bank", "ICICIBANK": "Private Bank", "AXISBANK": "Private Bank",
    "KOTAKBANK": "Private Bank", "INDUSINDBK": "Private Bank", "IDFCFIRSTB": "Private Bank",
    "FEDERALBNK": "Private Bank", "BANDHANBNK": "Private Bank", "AUBANK": "Private Bank",
    "SBIN": "PSU Bank", "BANKBARODA": "PSU Bank", "PNB": "PSU Bank",
    "CANBK": "PSU Bank", "UNIONBANK": "PSU Bank",
    # NBFC
    "BAJFINANCE": "NBFC", "BAJAJFINSV": "NBFC", "CHOLAFIN": "NBFC", "SHRIRAMFIN": "NBFC",
    "MUTHOOTFIN": "NBFC", "LICHSGFIN": "NBFC", "PFC": "NBFC", "RECLTD": "NBFC",
    "HDFCLIFE": "Insurance", "SBILIFE": "Insurance", "ICICIGI": "Insurance",
    "HDFCAMC": "Capital Mkt", "ANGELONE": "Capital Mkt", "BSE": "Capital Mkt",
    # IT
    "TCS": "IT", "INFY": "IT", "HCLTECH": "IT", "WIPRO": "IT", "TECHM": "IT",
    "LTM": "IT", "COFORGE": "IT", "PERSISTENT": "IT", "MPHASIS": "IT",
    # Auto
    "MARUTI": "Auto", "TMPV": "Auto", "TMCV": "Auto", "M&M": "Auto", "BAJAJ-AUTO": "Auto",
    "HEROMOTOCO": "Auto", "EICHERMOT": "Auto", "TVSMOTOR": "Auto", "ASHOKLEY": "Auto",
    "BOSCHLTD": "Auto Anc", "MOTHERSON": "Auto Anc", "BHARATFORG": "Auto Anc",
    "TIINDIA": "Auto Anc", "EXIDEIND": "Auto Anc",
    # Metals
    "TATASTEEL": "Metals", "JSWSTEEL": "Metals", "HINDALCO": "Metals", "VEDL": "Metals",
    "JINDALSTEL": "Metals", "SAIL": "Metals", "NATIONALUM": "Metals", "HINDZINC": "Metals",
    # Energy
    "RELIANCE": "Energy", "ONGC": "Energy", "BPCL": "Energy", "IOC": "Energy",
    "HINDPETRO": "Energy", "GAIL": "Energy", "OIL": "Energy", "PETRONET": "Energy",
    # Power
    "NTPC": "Power", "POWERGRID": "Power", "TATAPOWER": "Power", "ADANIPOWER": "Power",
    "JSWENERGY": "Power", "NHPC": "Power", "SJVN": "Power",
    # Infra
    "LT": "Cap Goods", "SIEMENS": "Cap Goods", "ABB": "Cap Goods", "BEL": "Defence",
    "HAL": "Defence", "BDL": "Defence", "MAZDOCK": "Defence", "CUMMINSIND": "Cap Goods",
    "THERMAX": "Cap Goods", "POLYCAB": "Cap Goods", "HAVELLS": "Cap Goods",
    # Adani
    "ADANIENT": "Adani", "ADANIPORTS": "Adani", "ADANIGREEN": "Adani", "AMBUJACEM": "Cement",
    # Cement
    "ULTRACEMCO": "Cement", "SHREECEM": "Cement", "ACC": "Cement", "DALBHARAT": "Cement",
    # Pharma
    "SUNPHARMA": "Pharma", "DRREDDY": "Pharma", "CIPLA": "Pharma", "DIVISLAB": "Pharma",
    "LUPIN": "Pharma", "AUROPHARMA": "Pharma", "ZYDUSLIFE": "Pharma", "TORNTPHARM": "Pharma",
    "ALKEM": "Pharma", "LAURUSLABS": "Pharma", "GLENMARK": "Pharma",
    "APOLLOHOSP": "Healthcare", "MAXHEALTH": "Healthcare",
    # FMCG
    "ITC": "FMCG", "HINDUNILVR": "FMCG", "NESTLEIND": "FMCG", "BRITANNIA": "FMCG",
    "TATACONSUM": "FMCG", "DABUR": "FMCG", "GODREJCP": "FMCG", "MARICO": "FMCG",
    "VBL": "FMCG", "COLPAL": "FMCG", "UNITDSPR": "FMCG",
    # Retail
    "TITAN": "Retail", "TRENT": "Retail", "DMART": "Retail", "ZOMATO": "New Age",
    "NYKAA": "New Age", "PAYTM": "New Age", "POLICYBZR": "New Age", "IRCTC": "New Age",
    # Telecom
    "BHARTIARTL": "Telecom", "IDEA": "Telecom", "INDUSTOWER": "Telecom",
    # Chemicals
    "PIDILITIND": "Chemicals", "SRF": "Chemicals", "UPL": "Chemicals", "TATACHEM": "Chemicals",
    "DEEPAKNTR": "Chemicals", "PIIND": "Chemicals",
    # Real Estate
    "DLF": "Realty", "GODREJPROP": "Realty", "OBEROIRLTY": "Realty", "LODHA": "Realty",
    "PRESTIGE": "Realty", "PHOENIXLTD": "Realty",
    # PSU
    "COALINDIA": "PSU", "IRFC": "PSU", "RVNL": "PSU", "IRCON": "PSU",
    "CONCOR": "Logistics", "INDIGO": "Aviation", "DIXON": "Electronics",
    "KAYNES": "Electronics", "CGPOWER": "Electronics",
}

SYMBOL_ALIASES: Dict[str, str] = {
    "LTIM": "LTM",
    "TATAMOTORS": "TMPV",
    "MINDTREE": "LTM",
    "HDFC": "HDFCBANK",
    "SBICARD": "SBICARDS",
}

_DEAD_SYMBOLS: set = set()


def yf_symbol(sym: str) -> str:
    return f"{SYMBOL_ALIASES.get(sym, sym)}.NS"


# =============================================================================
# 3. DATA LAYER
# =============================================================================

def _chunk(lst: List[str], n: int) -> List[List[str]]:
    return [lst[i:i + n] for i in range(0, len(lst), n)]


def _extract(raw: pd.DataFrame, ticker: str) -> Optional[pd.DataFrame]:
    try:
        if isinstance(raw.columns, pd.MultiIndex):
            if ticker in raw.columns.get_level_values(0):
                df = raw[ticker].copy()
            elif ticker in raw.columns.get_level_values(1):
                df = raw.xs(ticker, axis=1, level=1).copy()
            else:
                return None
        else:
            df = raw.copy()
        df = df.dropna(how="all")
        need = {"Open", "High", "Low", "Close", "Volume"}
        if not need.issubset(set(df.columns)) or df.empty:
            return None
        return df
    except Exception:
        return None


_CACHE: Dict[tuple, Dict[str, pd.DataFrame]] = {}


def download_bulk(symbols: List[str], period: str, interval: str,
                  chunk_size: int = 40) -> Dict[str, pd.DataFrame]:
    symbols = [s for s in symbols if s not in _DEAD_SYMBOLS]
    key = (tuple(sorted(symbols)), period, interval)
    if key in _CACHE:
        return _CACHE[key]

    out: Dict[str, pd.DataFrame] = {}
    tickers = [yf_symbol(s) for s in symbols]
    for grp in _chunk(tickers, chunk_size):
        try:
            raw = yf.download(
                " ".join(grp), period=period, interval=interval,
                group_by="ticker", auto_adjust=False, progress=False,
                threads=True, prepost=False,
            )
        except Exception:
            continue
        if raw is None or raw.empty:
            continue
        for t in grp:
            df = _extract(raw, t)
            if df is not None and len(df) > 5:
                out[t.replace(".NS", "")] = df
    if interval == "1d":
        for s in symbols:
            if s not in out:
                _DEAD_SYMBOLS.add(s)
    _CACHE[key] = out
    return out


def download_index(interval: str, period: str) -> Optional[pd.DataFrame]:
    try:
        df = yf.download("^NSEI", period=period, interval=interval,
                         auto_adjust=False, progress=False)
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        return df.dropna(how="all")
    except Exception:
        return None


# =============================================================================
# 4. INDICATORS
# =============================================================================

def atr_wilder(df: pd.DataFrame, period: int = 14) -> pd.Series:
    h, l, c = df["High"], df["Low"], df["Close"]
    pc = c.shift(1)
    tr = pd.concat([h - l, (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)
    return tr.ewm(alpha=1 / period, adjust=False).mean()


def ema(s: pd.Series, span: int) -> pd.Series:
    return s.ewm(span=span, adjust=False).mean()


def session_vwap(df: pd.DataFrame) -> pd.Series:
    tp = (df["High"] + df["Low"] + df["Close"]) / 3.0
    pv = (tp * df["Volume"]).cumsum()
    vv = df["Volume"].cumsum().replace(0, np.nan)
    return pv / vv


def to_ist(df: pd.DataFrame) -> pd.DataFrame:
    idx = df.index
    try:
        if idx.tz is None:
            df.index = idx.tz_localize("UTC").tz_convert(IST)
        else:
            df.index = idx.tz_convert(IST)
    except Exception:
        pass
    return df


def _completed_sessions_only(df: pd.DataFrame, session_date) -> pd.DataFrame:
    try:
        return df[np.array([d < session_date for d in df.index.date])]
    except Exception:
        return df


# =============================================================================
# 5. STOCK EVALUATION
# =============================================================================

@dataclass
class StockData:
    symbol: str
    sector: str
    close: float
    prev_close: float
    atr: float
    atr_pct: float
    turnover_cr: float
    ema20: float
    ema50: float
    above_ema20: bool
    above_ema50: bool
    dist_52w_high_pct: float
    vwap: float
    last: float
    or_high: float
    or_low: float
    gap_pct: float
    rvol: float
    day_high: float
    day_low: float
    open_px: float
    range_used_pct: float
    bar_close_strength: float


@dataclass
class RankedTrade:
    rank: int
    symbol: str
    sector: str
    direction: str
    score: float
    confidence: str
    entry: float
    stop: float
    target1: float
    target2: float
    qty: int
    notional: float
    margin: float
    risk_rs: float
    reward_rs: float
    rr: float
    reasons: List[str]
    warnings: List[str]
    price_vs_vwap: float
    range_used: float
    atr_pct: float
    rvol: float
    gap_pct: float
    trend: str


def evaluate_stock(sym: str, daily: pd.DataFrame, intraday: pd.DataFrame,
                   cfg: Config, session_date) -> Optional[StockData]:
    """Extract all data for a stock, no gates - just gather info."""
    try:
        daily = _completed_sessions_only(daily.dropna(), session_date)
        if len(daily) < 30:
            return None

        close = float(daily["Close"].iloc[-1])
        if close < cfg.min_price:
            return None

        turnover_cr = float((daily["Close"] * daily["Volume"]).tail(20).mean() / 1e7)
        if turnover_cr < cfg.min_turnover_cr:
            return None

        atr = float(atr_wilder(daily, 14).iloc[-1])
        atr_pct = atr / close * 100.0
        if not (cfg.min_atr_pct <= atr_pct <= cfg.max_atr_pct):
            return None

        e20 = float(ema(daily["Close"], 20).iloc[-1])
        e50 = float(ema(daily["Close"], 50).iloc[-1])
        hi52 = float(daily["High"].tail(250).max())

        intraday = to_ist(intraday.copy())
        days = sorted(set(intraday.index.date))
        if not days:
            return None
        target = session_date.date() if session_date is not None else days[-1]
        if target not in days:
            return None

        today = intraday[intraday.index.date == target]
        hh, mm = map(int, cfg.signal_time.split(":"))
        today = today[(today.index.hour * 60 + today.index.minute) <= hh * 60 + mm]
        if len(today) < 2:
            return None

        or_bar = today.iloc[0]
        vw = session_vwap(today)
        last_bar = today.iloc[-1]
        last = float(last_bar["Close"])

        n = len(today)
        prev_sessions = [intraday[intraday.index.date == d] for d in days if d < target]
        baselines = []
        for p in prev_sessions[-5:]:
            if len(p) >= n:
                vol = float(p["Volume"].iloc[:n].sum())
                if vol > 0:
                    baselines.append(vol)
        rvol = float(today["Volume"].sum()) / float(np.mean(baselines)) if baselines else 1.0

        return StockData(
            symbol=sym,
            sector=UNIVERSE.get(sym, "Other"),
            close=close,
            prev_close=float(daily["Close"].iloc[-2]) if len(daily) > 1 else close,
            atr=atr,
            atr_pct=atr_pct,
            turnover_cr=turnover_cr,
            ema20=e20,
            ema50=e50,
            above_ema20=close > e20,
            above_ema50=close > e50,
            dist_52w_high_pct=(close / hi52 - 1) * 100.0,
            vwap=float(vw.iloc[-1]),
            last=last,
            or_high=float(or_bar["High"]),
            or_low=float(or_bar["Low"]),
            gap_pct=(float(today["Open"].iloc[0]) / float(daily["Close"].iloc[-1]) - 1) * 100.0,
            rvol=rvol,
            day_high=float(today["High"].max()),
            day_low=float(today["Low"].min()),
            open_px=float(today["Open"].iloc[0]),
            range_used_pct=(float(today["High"].max() - today["Low"].min()) / atr * 100.0 if atr else 0),
            bar_close_strength=(last - float(last_bar["Low"])) / (float(last_bar["High"]) - float(last_bar["Low"]) + 0.001),
        )
    except Exception:
        return None


def rank_stock(data: StockData, cfg: Config) -> Optional[RankedTrade]:
    """Score the stock and generate trade plan."""
    above_vwap = data.last > data.vwap
    below_vwap = data.last < data.vwap
    broke_orh = data.last > data.or_high
    broke_orl = data.last < data.or_low

    buy_score = 0.0
    sell_score = 0.0

    if above_vwap and broke_orh:
        buy_score += 40
    elif below_vwap and broke_orl and cfg.allow_shorts:
        sell_score += 40
    elif above_vwap:
        buy_score += 20
    elif below_vwap and cfg.allow_shorts:
        sell_score += 20

    if data.above_ema20:
        buy_score += 15
    else:
        sell_score += 15
    if data.above_ema50:
        buy_score += 10
    else:
        sell_score += 10

    vol_score = min(data.rvol / 2.0, 1.0) * 20
    buy_score += vol_score
    sell_score += vol_score

    if abs(data.gap_pct) < 1.0:
        buy_score += 5
        sell_score += 5

    range_penalty = min(data.range_used_pct / 100.0, 1.0) * 10
    buy_score -= range_penalty
    sell_score -= range_penalty

    if buy_score > sell_score:
        direction = "BUY"
        score = min(buy_score, 100)
    elif sell_score > buy_score and cfg.allow_shorts:
        direction = "SELL"
        score = min(sell_score, 100)
    else:
        return None

    entry = data.last
    if direction == "BUY":
        sl = min(data.or_low, data.vwap - 0.3 * data.atr)
        sl = max(sl, entry * (1 - cfg.max_sl_pct / 100))
        sl = min(sl, entry * (1 - cfg.min_sl_pct / 100))
        t1 = entry + cfg.t1_r * abs(entry - sl)
        t2 = entry + cfg.t2_r * abs(entry - sl)
        trend = "Bullish" if data.above_ema20 else "Neutral"
    else:
        sl = max(data.or_high, data.vwap + 0.3 * data.atr)
        sl = min(sl, entry * (1 + cfg.max_sl_pct / 100))
        sl = max(sl, entry * (1 + cfg.min_sl_pct / 100))
        t1 = entry - cfg.t1_r * abs(entry - sl)
        t2 = entry - cfg.t2_r * abs(entry - sl)
        trend = "Bearish" if not data.above_ema20 else "Neutral"

    risk_per_share = abs(entry - sl)
    if risk_per_share <= 0 or risk_per_share / entry > 0.05:
        return None

    risk_budget = cfg.capital * cfg.risk_pct / 100.0
    qty_by_risk = math.floor(risk_budget / risk_per_share)
    buying_power = cfg.capital * cfg.leverage
    qty_by_margin = math.floor(buying_power * 0.4 / entry)
    qty = max(min(qty_by_risk, qty_by_margin), 0)
    if qty == 0:
        return None

    reasons = []
    if direction == "BUY":
        reasons.append(f"Price ({data.last:.2f}) above VWAP ({data.vwap:.2f}) and above OR high ({data.or_high:.2f})")
    else:
        reasons.append(f"Price ({data.last:.2f}) below VWAP ({data.vwap:.2f}) and below OR low ({data.or_low:.2f})")
    reasons.append(f"ATR: {data.atr_pct:.1f}% (Rs {data.atr:.2f}) - good volatility")
    reasons.append(f"Volume: {data.rvol:.1f}x normal - {'strong' if data.rvol > 1.5 else 'moderate'} participation")
    if data.above_ema20:
        reasons.append(f"Above 20-day EMA ({data.ema20:.2f}) - bullish trend")
    else:
        reasons.append(f"Below 20-day EMA ({data.ema20:.2f}) - bearish trend")
    reasons.append(f"Sector: {data.sector}")

    warnings = []
    if abs(data.gap_pct) > 1.5:
        warnings.append(f"Large gap {data.gap_pct:+.1f}% - may fill")
    if data.range_used_pct > 70:
        warnings.append(f"Used {data.range_used_pct:.0f}% of daily range already")

    if score >= 80:
        confidence = "★★★★★"
    elif score >= 70:
        confidence = "★★★★"
    elif score >= 60:
        confidence = "★★★"
    elif score >= 50:
        confidence = "★★"
    else:
        confidence = "★"

    return RankedTrade(
        rank=0,
        symbol=data.symbol,
        sector=data.sector,
        direction=direction,
        score=round(score, 1),
        confidence=confidence,
        entry=round(entry, 2),
        stop=round(sl, 2),
        target1=round(t1, 2),
        target2=round(t2, 2),
        qty=qty,
        notional=round(qty * entry, 0),
        margin=round(qty * entry / cfg.leverage, 0),
        risk_rs=round(qty * risk_per_share, 0),
        reward_rs=round(qty * abs(t1 - entry), 0),
        rr=round(abs(t1 - entry) / risk_per_share, 2),
        reasons=reasons,
        warnings=warnings,
        price_vs_vwap=round((data.last / data.vwap - 1) * 100, 2),
        range_used=round(data.range_used_pct, 1),
        atr_pct=round(data.atr_pct, 1),
        rvol=round(data.rvol, 2),
        gap_pct=round(data.gap_pct, 2),
        trend=trend,
    )


# =============================================================================
# 6. EXCEL EXPORT
# =============================================================================

def export_to_excel(trades: List[RankedTrade], cfg: Config, label: str) -> str:
    """Export trades to formatted Excel with multiple sheets."""
    os.makedirs(cfg.outdir, exist_ok=True)
    filepath = os.path.join(cfg.outdir, f"picks_{label}.xlsx")

    rows = []
    for t in trades:
        rows.append({
            "Rank": t.rank,
            "Symbol": t.symbol,
            "Sector": t.sector,
            "Direction": t.direction,
            "Score": t.score,
            "Confidence": t.confidence,
            "Entry": t.entry,
            "Stop Loss": t.stop,
            "Target 1": t.target1,
            "Target 2": t.target2,
            "Qty": t.qty,
            "Notional (Rs)": t.notional,
            "Margin (Rs)": t.margin,
            "Risk (Rs)": t.risk_rs,
            "Reward (Rs)": t.reward_rs,
            "R:R": t.rr,
            "ATR %": t.atr_pct,
            "RVOL": t.rvol,
            "Gap %": t.gap_pct,
            "Range Used %": t.range_used,
            "Price vs VWAP %": t.price_vs_vwap,
            "Trend": t.trend,
        })

    df = pd.DataFrame(rows)

    wb = Workbook()

    # Sheet 1: Rankings
    ws1 = wb.active
    ws1.title = "Rankings"

    headers = list(df.columns)
    for col, header in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(df.values, 2):
        for c_idx, val in enumerate(row, 1):
            ws1.cell(row=r_idx, column=c_idx, value=val)

    for col in ws1.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws1.column_dimensions[col_letter].width = min(max_len + 2, 25)

    for row in range(2, len(df) + 2):
        dir_cell = ws1.cell(row=row, column=4)
        if dir_cell.value == "BUY":
            dir_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            dir_cell.font = Font(color="006100")
        elif dir_cell.value == "SELL":
            dir_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            dir_cell.font = Font(color="9C0006")

    # Sheet 2: Trade Plan
    ws2 = wb.create_sheet("Trade Plan")
    plan_headers = ["Rank", "Symbol", "Direction", "Entry", "Stop", "Target 1", "Target 2", "Qty", "Risk", "Reward"]
    for col, h in enumerate(plan_headers, 1):
        cell = ws2.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")

    for r_idx, t in enumerate(trades, 2):
        ws2.cell(row=r_idx, column=1, value=t.rank)
        ws2.cell(row=r_idx, column=2, value=t.symbol)
        ws2.cell(row=r_idx, column=3, value=t.direction)
        ws2.cell(row=r_idx, column=4, value=t.entry)
        ws2.cell(row=r_idx, column=5, value=t.stop)
        ws2.cell(row=r_idx, column=6, value=t.target1)
        ws2.cell(row=r_idx, column=7, value=t.target2)
        ws2.cell(row=r_idx, column=8, value=t.qty)
        ws2.cell(row=r_idx, column=9, value=t.risk_rs)
        ws2.cell(row=r_idx, column=10, value=t.reward_rs)

    for col in ws2.columns:
        ws2.column_dimensions[col[0].column_letter].width = 14

    # Sheet 3: Reasons
    ws3 = wb.create_sheet("Reasons")
    ws3.cell(row=1, column=1, value="Symbol").font = Font(bold=True)
    ws3.cell(row=1, column=2, value="Direction").font = Font(bold=True)
    ws3.cell(row=1, column=3, value="Reasons").font = Font(bold=True)
    ws3.cell(row=1, column=4, value="Watch-outs").font = Font(bold=True)
    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 12
    ws3.column_dimensions["C"].width = 50
    ws3.column_dimensions["D"].width = 40

    for r_idx, t in enumerate(trades, 2):
        ws3.cell(row=r_idx, column=1, value=t.symbol)
        ws3.cell(row=r_idx, column=2, value=t.direction)
        ws3.cell(row=r_idx, column=3, value="\n".join(t.reasons))
        ws3.cell(row=r_idx, column=4, value="\n".join(t.warnings) if t.warnings else "None")
        for col in range(1, 5):
            ws3.cell(row=r_idx, column=col).alignment = Alignment(wrap_text=True, vertical="top")

    # Sheet 4: Summary
    ws4 = wb.create_sheet("Summary")
    ws4.cell(row=1, column=1, value="Metric").font = Font(bold=True, size=12)
    ws4.cell(row=1, column=2, value="Value").font = Font(bold=True, size=12)

    metrics = [
        ("Date", label),
        ("Total Trades", len(trades)),
        ("Capital", f"Rs {cfg.capital:,.0f}"),
        ("Leverage", f"{cfg.leverage}x"),
        ("Risk per Trade", f"{cfg.risk_pct}%"),
        ("Total Risk if All SL Hit", f"Rs {sum(t.risk_rs for t in trades):,.0f}"),
        ("Total Margin Needed", f"Rs {sum(t.margin for t in trades):,.0f}"),
        ("Buy/Sell Split", f"{sum(1 for t in trades if t.direction == 'BUY')}/{sum(1 for t in trades if t.direction == 'SELL')}"),
        ("Avg Score", f"{sum(t.score for t in trades) / len(trades):.1f}" if trades else "N/A"),
        ("Avg R:R", f"{sum(t.rr for t in trades) / len(trades):.2f}" if trades else "N/A"),
    ]

    for r_idx, (k, v) in enumerate(metrics, 2):
        ws4.cell(row=r_idx, column=1, value=k)
        ws4.cell(row=r_idx, column=2, value=v)

    ws4.column_dimensions["A"].width = 25
    ws4.column_dimensions["B"].width = 25

    # Sheet 5: Full Data
    ws5 = wb.create_sheet("Full Data")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws5.append(r)
    for col in ws5.columns:
        ws5.column_dimensions[col[0].column_letter].width = 14

    wb.save(filepath)
    return filepath


# =============================================================================
# 7. MAIN SCAN ENGINE
# =============================================================================

def run_scan(cfg: Config, session: Optional[pd.Timestamp] = None) -> List[RankedTrade]:
    label = session.strftime("%Y-%m-%d") if session is not None else datetime.now().strftime("%Y-%m-%d")

    print(f"\n{'='*70}")
    print(f"  NSE INTRADAY SCANNER - {label}")
    print(f"  Capital: Rs {cfg.capital:,.0f}  |  Leverage: {cfg.leverage}x  |  Risk: {cfg.risk_pct}%")
    print(f"{'='*70}")

    print(f"\n[1/4] Downloading daily data for {len(UNIVERSE)} stocks...")
    daily_data = download_bulk(list(UNIVERSE.keys()), period="1y", interval="1d")
    print(f"      Got {len(daily_data)} stocks")

    print(f"[2/4] Downloading intraday data for {len(daily_data)} stocks...")
    intraday_data = download_bulk(list(daily_data.keys()), period="1mo", interval="15m")
    print(f"      Got {len(intraday_data)} stocks")

    print(f"[3/4] Evaluating and ranking stocks...")
    session_date = session if session is not None else pd.Timestamp.now(tz=IST)

    all_trades: List[RankedTrade] = []
    for sym, daily in daily_data.items():
        if sym not in intraday_data:
            continue
        data = evaluate_stock(sym, daily, intraday_data[sym], cfg, session_date)
        if data is None:
            continue
        trade = rank_stock(data, cfg)
        if trade is not None:
            all_trades.append(trade)

    all_trades.sort(key=lambda x: -x.score)

    # Assign ranks and limit
    for i, t in enumerate(all_trades[:cfg.top_n], 1):
        t.rank = i

    top_trades = all_trades[:cfg.top_n]

    print(f"      Found {len(all_trades)} tradable stocks, showing top {cfg.top_n}")

    # Print summary to terminal
    print(f"\n{'='*70}")
    print(f"  TOP {len(top_trades)} STOCKS")
    print(f"{'='*70}")
    for t in top_trades:
        print(f"  #{t.rank:2} {t.direction:4} {t.symbol:<12} Score: {t.score:5.1f} {t.confidence}  Entry: {t.entry:8.2f}  SL: {t.stop:8.2f}  T1: {t.target1:8.2f}  T2: {t.target2:8.2f}")

    # Export
    if top_trades and not cfg.no_save:
        print(f"\n[4/4] Exporting to Excel...")
        xlsx_path = export_to_excel(top_trades, cfg, label)

        # Also CSV
        csv_path = os.path.join(cfg.outdir, f"picks_{label}.csv")
        rows = []
        for t in top_trades:
            rows.append({
                "Rank": t.rank, "Symbol": t.symbol, "Sector": t.sector,
                "Direction": t.direction, "Score": t.score, "Confidence": t.confidence,
                "Entry": t.entry, "Stop": t.stop, "Target1": t.target1, "Target2": t.target2,
                "Qty": t.qty, "Notional": t.notional, "Margin": t.margin,
                "Risk": t.risk_rs, "Reward": t.reward_rs, "RR": t.rr,
                "Reasons": " | ".join(t.reasons), "Warnings": " | ".join(t.warnings)
            })
        pd.DataFrame(rows).to_csv(csv_path, index=False)

        print(f"\n  ✅ Excel saved to: {xlsx_path}")
        print(f"  ✅ CSV saved to:  {csv_path}")
        print(f"\n  📊 Open the Excel file - it has 5 sheets:")
        print(f"     1. Rankings - all data with color coding")
        print(f"     2. Trade Plan - just entry/SL/targets")
        print(f"     3. Reasons - why each stock was selected")
        print(f"     4. Summary - overall metrics")
        print(f"     5. Full Data - raw data table")

    return top_trades


# NOTE: CLI entrypoint removed here — this module is imported by app.py (Streamlit UI).
# All core logic (Config, run_scan, evaluate_stock, rank_stock, export_to_excel, etc.)
# above this point is UNCHANGED from the original script.