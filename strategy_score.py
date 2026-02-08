"""
strategy_score.py  (v6.3.29-F4)

集中管理：
- 綜合分數（實盤精煉版）
- 燈號：嘎空壓力 / 周轉率 / 成交值
"""
from __future__ import annotations
import math
import pandas as pd


def light_label(x: float, t1: float, t2: float, reverse: bool=False) -> str:
    """Return emoji light based on thresholds.
    reverse=False: higher is better (green). reverse=True: higher is worse (red).
    """
    try:
        if x is None or pd.isna(x):
            return "N/A"
        v = float(x)
    except Exception:
        return "N/A"
    if reverse:
        if v >= t2:
            return "🔴"
        if v >= t1:
            return "🟡"
        return "🟢"
    else:
        if v >= t2:
            return "🟢"
        if v >= t1:
            return "🟡"
        return "🔴"

def _to_num(s):
    return pd.to_numeric(s, errors="coerce")

def amount_threshold_by_price(price: float) -> float:
    """依股價估算最低成交值門檻(NTD)。可自行調整。"""
    try:
        p = float(price)
    except Exception:
        return 3_000_000.0
    if p < 20:   return 3_000_000.0
    if p < 50:   return 5_000_000.0
    if p < 100:  return 8_000_000.0
    if p < 200:  return 12_000_000.0
    return 18_000_000.0

def _rank01(series, higher_better=True):
    s = _to_num(series)
    if s.isna().all():
        return pd.Series([0.0]*len(series), index=series.index)
    r = s.rank(pct=True)
    if not higher_better:
        r = 1.0 - r
    return r.fillna(0.0)

def _light_from_levels(x, hi, mid, reverse=False):
    """回傳：🟢/🟡/🔴/N/A。reverse=True 表示數值越大越危險。"""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "N/A"
    try:
        v = float(x)
    except Exception:
        return "N/A"
    if v == 0:
        return "N/A"
    if not reverse:
        if v >= hi:  return "🟢"
        if v >= mid: return "🟡"
        return "🔴"
    else:
        if v >= hi:  return "🔴"
        if v >= mid: return "🟡"
        return "🟢"

def add_lights(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    # --- Row-wise derive/repair squeeze pressure (嘎空壓力) from SMR ---
    smr_col = None
    for _c in ["券資比(%)", "short_margin_ratio(%)", "SMR(%)"]:
        if _c in out.columns:
            smr_col = _c
            break

    if "嘎空壓力" not in out.columns:
        out["嘎空壓力"] = float("nan")

    if smr_col:
        smr = pd.to_numeric(out[smr_col], errors="coerce")
        sq = pd.to_numeric(out.get("嘎空壓力"), errors="coerce")
        need = sq.isna()
        derived = ((smr - 9.0) / (30.0 - 9.0)).clip(lower=0.0, upper=1.0)
        out.loc[need, "嘎空壓力"] = derived.loc[need].round(4)

    # --- Turnover light (周轉率燈號) ---
    if "周轉率燈號" not in out.columns:
        out["周轉率燈號"] = ""
    tr_col = None
    for _c in ["周轉率(%)", "turnover_rate(%)", "turnover_rate"]:
        if _c in out.columns:
            tr_col = _c
            break
    if tr_col:
        tr = pd.to_numeric(out[tr_col], errors="coerce")
        out["周轉率燈號"] = tr.apply(lambda v: light_label(v, 0.2, 0.5, reverse=False))

    # --- Squeeze light (嘎空壓力燈號) ---
    if "嘎空壓力燈號" not in out.columns:
        out["嘎空壓力燈號"] = ""
    sqv = pd.to_numeric(out.get("嘎空壓力"), errors="coerce")
    out["嘎空壓力燈號"] = sqv.apply(lambda v: light_label(v, 0.33, 0.66, reverse=True))



    
    # --- derive/repair squeeze pressure for ALL markets from SMR (券資比%) ---
    # If "嘎空壓力" is missing OR some rows are empty/NaN, derive from SMR row-wise.
    smr_col = None
    for _c in ["券資比(%)", "short_margin_ratio(%)", "SMR(%)"]:
        if _c in out.columns:
            smr_col = _c
            break

    if smr_col:
        smr = pd.to_numeric(out[smr_col], errors="coerce")
        derived = ((smr - 9.0) / (30.0 - 9.0)).clip(lower=0.0, upper=1.0)
        if "嘎空壓力" in out.columns:
            cur = pd.to_numeric(out["嘎空壓力"], errors="coerce")
            out["嘎空壓力"] = cur.where(cur.notna(), derived).round(4)
        else:
            out["嘎空壓力"] = derived.round(4)
    else:
        if "嘎空壓力" not in out.columns:
            out["嘎空壓力"] = float("nan")


    # --- derive squeeze pressure for ALL markets from SMR (券資比%) if missing ---
    if "嘎空壓力" not in out.columns:
        smr_col = None
        for _c in ["券資比(%)", "short_margin_ratio(%)", "SMR(%)"]:
            if _c in out.columns:
                smr_col = _c
                break
        if smr_col:
            smr = pd.to_numeric(out[smr_col], errors="coerce")
            out["嘎空壓力"] = ((smr - 9.0) / (30.0 - 9.0)).clip(lower=0.0, upper=1.0).round(4)
        else:
            out["嘎空壓力"] = float("nan")


    col_pressure = "嘎空壓力" if "嘎空壓力" in out.columns else ("otc_short_pressure" if "otc_short_pressure" in out.columns else None)
    col_turn = "周轉率(%)" if "周轉率(%)" in out.columns else ("turnover_rate(%)" if "turnover_rate(%)" in out.columns else None)
    col_tv = "成交值(元)" if "成交值(元)" in out.columns else ("traded_value_ntd" if "traded_value_ntd" in out.columns else None)
    col_price = "進場價" if "進場價" in out.columns else ("entry_price" if "entry_price" in out.columns else ("收盤價" if "收盤價" in out.columns else ("close" if "close" in out.columns else None)))

    if col_pressure:
        pr = _to_num(out[col_pressure])
        out["嘎空壓力燈號"] = pr.apply(lambda x: _light_from_levels(x, hi=0.90, mid=0.70, reverse=True))
    else:
        out["嘎空壓力燈號"] = "N/A"

    if col_turn:
        tr = _to_num(out[col_turn])
        out["周轉率燈號"] = tr.apply(lambda x: _light_from_levels(x, hi=1.00, mid=0.30, reverse=False))
    else:
        out["周轉率燈號"] = "N/A"

    if col_tv and col_price:
        tv = _to_num(out[col_tv])
        px = _to_num(out[col_price])
        lights=[]
        for t, p in zip(tv.tolist(), px.tolist()):
            if t is None or (isinstance(t,float) and math.isnan(t)) or float(t) <= 0:
                lights.append("N/A"); continue
            thr = amount_threshold_by_price(p)
            if float(t) >= thr*2: lights.append("🟢")
            elif float(t) >= thr: lights.append("🟡")
            else: lights.append("🔴")
        out["成交值燈號"] = lights
    else:
        out["成交值燈號"] = "N/A"
    return out

def compute_composite_score_live(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()
    out = df.copy()

    col_score = "策略分數" if "策略分數" in out.columns else ("strategy_score" if "strategy_score" in out.columns else None)
    col_bias = "乖離率(%)" if "乖離率(%)" in out.columns else ("bias20" if "bias20" in out.columns else None)
    col_turn = "周轉率(%)" if "周轉率(%)" in out.columns else None
    col_tv = "成交值(元)" if "成交值(元)" in out.columns else None
    col_pos = "建議部位(元)" if "建議部位(元)" in out.columns else ("position_size" if "position_size" in out.columns else None)
    col_vola = "年化波動" if "年化波動" in out.columns else ("vol_annual" if "vol_annual" in out.columns else None)
    col_price = "進場價" if "進場價" in out.columns else ("entry_price" if "entry_price" in out.columns else None)

    w_score = _rank01(out[col_score] if col_score else pd.Series([0.0]*len(out), index=out.index), True)
    w_pos = _rank01(out[col_pos] if col_pos else pd.Series([0.0]*len(out), index=out.index), True)

    if col_bias:
        b = _to_num(out[col_bias]).fillna(0.0).abs()
        w_bias = _rank01(b, True)
    else:
        w_bias = pd.Series([0.0]*len(out), index=out.index)

    w_turn = _rank01(out[col_turn] if col_turn else pd.Series([0.0]*len(out), index=out.index), True)
    w_tv = _rank01(out[col_tv] if col_tv else pd.Series([0.0]*len(out), index=out.index), True)

    out["成交值排名"] = w_tv.round(4)

    prices = _to_num(out[col_price]) if col_price else pd.Series([math.nan]*len(out), index=out.index)
    tvv = _to_num(out[col_tv]) if col_tv else pd.Series([math.nan]*len(out), index=out.index)
    vola = _to_num(out[col_vola]) if col_vola else pd.Series([math.nan]*len(out), index=out.index)

    pen=[]
    for pr, tv, va in zip(prices.tolist(), tvv.tolist(), vola.tolist()):
        p=0.0
        try:
            thr_amt = amount_threshold_by_price(pr)
            if not (isinstance(tv,float) and math.isnan(tv)):
                if float(tv) < float(thr_amt): p += 0.25
            if (not (isinstance(va,float) and math.isnan(va))) and (not (isinstance(tv,float) and math.isnan(tv))):
                if float(va) > 0.50 and float(tv) < float(thr_amt)*1.5: p += 0.25
        except Exception:
            pass
        pen.append(p)
    penalty = pd.Series(pen, index=out.index)
    out["流動性扣分"] = penalty.round(2)

    comp = 0.35*w_score + 0.20*w_bias + 0.20*w_tv + 0.10*w_turn + 0.15*w_pos - penalty
    out["綜合分數"] = comp.round(4)
    return out

def apply_live_scoring(df: pd.DataFrame) -> pd.DataFrame:
    out = add_lights(df)
    out = compute_composite_score_live(out)
    return out
