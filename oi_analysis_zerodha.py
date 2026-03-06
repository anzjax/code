# ════════════════════════════════════════════════════════════════
#  📊 LIVE OI ANALYSIS OPTION CHAIN — ZERODHA KITE API (POWER QUERY)
# ════════════════════════════════════════════════════════════════
#  What this script does:
#  ✅ Greeks removed (no Delta/Gamma/Theta/Vega)
#  ✅ Output = NSE layout CSV (CALLS | STRIKE | PUTS) for Power Query
#  ✅ Excel file open rahe, Power Query refresh ho
#  ✅ Python auto-updates CSV every N seconds
#
#  IMPORTANT TRUTH:
#  - Power Query "auto refresh" Excel ke andar setting hoti hai (Query Properties).
#    Python code Excel ko auto refresh force nahi kar sakta reliably.
#  - Python ka kaam: CSV continuously update karna. Excel PQ refresh: every 1 minute.
#
#  Output files:
#   1) oi_chain_nse_layout_latest.csv
#   2) oi_dashboard_latest.csv
#
#  Login:
#   Reads api_key + access_token from zarodhalogin.xlsx (sheet: logindata)
#    - api_key      -> B2 (df.iloc[1,1])
#    - access_token -> B4 (df.iloc[3,1])
#
#  Run:
#    python oi_analysis_zerodha.py
# ════════════════════════════════════════════════════════════════

from __future__ import annotations

import os
import time
import json
import warnings
from datetime import datetime, date
from typing import Any

import pandas as pd
from kiteconnect import KiteConnect

warnings.filterwarnings("ignore")

# ==============================================================================
# ⚙️ CONFIG (EDIT ONLY THIS SECTION)
# ==============================================================================
BASE_FOLDER = r"C:\New folder\naya work"
LOGIN_FILE_PATH = r"D:\zarodhatradeonlevel16122025\excel group\zarodhalogin.xlsx"
LOGIN_SHEET = "logindata"

CSV_NSE_LAYOUT = os.path.join(BASE_FOLDER, "oi_chain_nse_layout_latest.csv")
CSV_DASHBOARD = os.path.join(BASE_FOLDER, "oi_dashboard_latest.csv")
INSTRUMENTS_FILE = os.path.join(BASE_FOLDER, "instruments_cache.csv")

# Change-in-OI snapshot (so Chg in OI is not always 0)
SNAPSHOT_OI_FILE = os.path.join(BASE_FOLDER, "oi_last_snapshot_by_strike.json")

# Rate-limit/network safe settings
QUOTE_BATCH_SIZE = 120
QUOTE_BATCH_SLEEP = 0.40
QUOTE_RETRIES = 3
QUOTE_RETRY_SLEEP_BASE = 1.5

INSTRUMENTS_RETRIES = 5
INSTRUMENTS_RETRY_SLEEP_BASE = 2.0

DEFAULT_REFRESH_SECONDS = 60  # Python updates CSV every 60 sec


# ==============================================================================
# 📌 NSE-LAYOUT COLUMNS (NO GREEKS)
# ==============================================================================
# CALLS side (14 columns)
CE_COLS = [
    "CE_OI",
    "CE_Chg_in_OI",
    "CE_Volume",
    "CE_IV",
    "CE_LTP",
    "CE_Net_Chg",
    "CE_Bid_Qty",
    "CE_Bid_Price",
    "CE_Ask_Price",
    "CE_Ask_Qty",
    "CE_Open",
    "CE_High",
    "CE_Low",
    "CE_Close",
]

# PUTS side (14 columns)
PE_COLS = [
    "PE_Bid_Qty",
    "PE_Bid_Price",
    "PE_Ask_Price",
    "PE_Ask_Qty",
    "PE_Net_Chg",
    "PE_LTP",
    "PE_IV",
    "PE_Volume",
    "PE_Chg_in_OI",
    "PE_OI",
    "PE_Open",
    "PE_High",
    "PE_Low",
    "PE_Close",
]

ALL_COLS = CE_COLS + ["Strike"] + PE_COLS  # 29 columns


# ==============================================================================
# 🧩 UTILS
# ==============================================================================
def chunk_list(items: list[Any], size: int) -> list[list[Any]]:
    return [items[i:i + size] for i in range(0, len(items), size)]


def retry_call(fn, *, retries: int, sleep_base: float, what: str):
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            return fn()
        except Exception as e:
            last_err = e
            wait = min(sleep_base * (2 ** (attempt - 1)), 30)
            print(f"⚠️ {what} failed (attempt {attempt}/{retries}): {e}")
            print(f"   retrying in {wait:.1f}s...")
            time.sleep(wait)
    raise last_err


# ==============================================================================
# 🔐 LOGIN (Excel)
# ==============================================================================
def kite_login_from_excel() -> KiteConnect:
    print("🔐 Login Process Started (Excel)...")
    df = pd.read_excel(LOGIN_FILE_PATH, sheet_name=LOGIN_SHEET, header=None)
    api_key = str(df.iloc[1, 1]).strip()
    access_token = str(df.iloc[3, 1]).strip()

    kite = KiteConnect(api_key=api_key)
    kite.set_access_token(access_token)

    # validate token
    kite.profile()
    print("✅ Login Successful!")
    return kite


# ==============================================================================
# 📦 INSTRUMENTS (daily cache + retry)
# ==============================================================================
def download_instruments(kite: KiteConnect, force: bool = False) -> pd.DataFrame:
    os.makedirs(BASE_FOLDER, exist_ok=True)

    if not force and os.path.exists(INSTRUMENTS_FILE):
        try:
            file_date = date.fromtimestamp(os.path.getmtime(INSTRUMENTS_FILE))
            if file_date == date.today():
                print("✅ Instruments cache is up to date")
                return pd.read_csv(INSTRUMENTS_FILE, parse_dates=["expiry"])
        except Exception:
            pass

    def _do():
        print("⬇️  Downloading instruments...")
        inst = pd.DataFrame(kite.instruments())
        inst["expiry"] = pd.to_datetime(inst["expiry"], errors="coerce")
        inst.to_csv(INSTRUMENTS_FILE, index=False)
        print(f"✅ {len(inst)} instruments saved!")
        return inst

    return retry_call(
        _do,
        retries=INSTRUMENTS_RETRIES,
        sleep_base=INSTRUMENTS_RETRY_SLEEP_BASE,
        what="instruments download",
    )


# ==============================================================================
# 🎛️ SELECT SYMBOL + EXPIRY
# ==============================================================================
POPULAR = [
    "NIFTY", "BANKNIFTY", "FINNIFTY", "MIDCPNIFTY",
    "RELIANCE", "TCS", "INFY", "HDFCBANK", "ICICIBANK",
    "SBIN", "ITC", "LT", "AXISBANK", "KOTAKBANK",
    "BAJFINANCE", "MARUTI", "WIPRO", "HCLTECH",
]


def get_fno_symbols(instruments: pd.DataFrame) -> list[str]:
    fno = instruments[instruments["segment"] == "NFO-OPT"]
    return sorted(fno["name"].unique().tolist())


def show_stock_selector(instruments: pd.DataFrame) -> str:
    symbols = get_fno_symbols(instruments)
    print("\n" + "═" * 60)
    print("  📊  OI ANALYSIS — STOCK / INDEX SELECTOR")
    print("═" * 60)
    print("\n🔥 POPULAR:")
    for i, s in enumerate(POPULAR, 1):
        if s in symbols:
            print(f"   {i:2d}. {s}")
    print(f"\n📋 Total F&O Symbols: {len(symbols)}")
    print("   (Type name directly ya number enter karein)\n")

    while True:
        choice = input("🔎 Stock/Index: ").strip().upper()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(POPULAR) and POPULAR[idx] in symbols:
                return POPULAR[idx]
        if choice in symbols:
            return choice
        matches = [s for s in symbols if choice in s]
        if matches:
            print(f"   Did you mean: {', '.join(matches[:10])} ?")
        else:
            print("   ❌ Not found, try again")


def show_expiry_selector(instruments: pd.DataFrame, symbol: str):
    opts = instruments[
        (instruments["name"] == symbol) &
        (instruments["segment"] == "NFO-OPT")
    ]
    expiries = sorted(opts["expiry"].dropna().unique())
    print(f"\n📅 {symbol} — Available Expiries:")
    for i, e in enumerate(expiries, 1):
        tag = " 👈 Nearest" if i == 1 else ""
        print(f"   {i}. {str(e)[:10]}{tag}")

    while True:
        c = input(f"\nExpiry (1-{len(expiries)}) [1]: ").strip()
        if c == "":
            return expiries[0]
        if c.isdigit() and 1 <= int(c) <= len(expiries):
            return expiries[int(c) - 1]
        print("   ❌ Invalid")


# ==============================================================================
# 📈 SPOT + SAFE QUOTE
# ==============================================================================
INDEX_MAP = {
    "NIFTY": "NSE:NIFTY 50",
    "BANKNIFTY": "NSE:NIFTY BANK",
    "FINNIFTY": "NSE:NIFTY FIN SERVICE",
    "MIDCPNIFTY": "NSE:NIFTY MID SELECT",
}


def get_spot_price(kite: KiteConnect, symbol: str) -> float | None:
    key = INDEX_MAP.get(symbol, f"NSE:{symbol}")
    try:
        return kite.ltp(key)[key]["last_price"]
    except Exception:
        return None


def safe_quote(kite: KiteConnect, kite_symbols: list[str]) -> dict:
    quotes_all: dict = {}
    batches = chunk_list(kite_symbols, QUOTE_BATCH_SIZE)

    for bi, batch in enumerate(batches, 1):
        def _do(b=batch):
            return kite.quote(b)

        q = retry_call(
            _do,
            retries=QUOTE_RETRIES,
            sleep_base=QUOTE_RETRY_SLEEP_BASE,
            what=f"quote batch {bi}/{len(batches)}",
        )
        quotes_all.update(q)

        if bi < len(batches):
            time.sleep(QUOTE_BATCH_SLEEP)

    return quotes_all


# ==============================================================================
# 🔎 FETCH OPTION CHAIN (raw rows, no greeks)
# ==============================================================================
def fetch_option_chain(kite: KiteConnect, instruments: pd.DataFrame, symbol: str, expiry):
    print(f"\n🔄 Fetching {symbol} option chain for {str(expiry)[:10]}...")

    opts = instruments[
        (instruments["name"] == symbol) &
        (instruments["segment"] == "NFO-OPT") &
        (instruments["expiry"] == expiry)
    ].copy()

    if opts.empty:
        print("❌ No options found!")
        return None

    spot = get_spot_price(kite, symbol)
    print(f"💰 Spot Price : {spot}")

    expiry_dt = pd.Timestamp(expiry)
    dte = max((expiry_dt - pd.Timestamp.now()).days, 1)
    print(f"📅 Days to Exp: {dte}")

    ts_list = opts["tradingsymbol"].tolist()
    kite_symbols = [f"NFO:{s}" for s in ts_list]

    print(f"📡 Fetching quotes for {len(kite_symbols)} options...")
    all_quotes = safe_quote(kite, kite_symbols)

    rows = []
    for _, r in opts.iterrows():
        ts = r["tradingsymbol"]
        strike = float(r["strike"])
        otype = r["instrument_type"]  # CE/PE
        key = f"NFO:{ts}"
        q = all_quotes.get(key)
        if not q:
            continue

        ohlc = q.get("ohlc", {}) or {}
        depth = q.get("depth", {}) or {}
        buy = depth.get("buy", [{}]) or [{}]
        sell = depth.get("sell", [{}]) or [{}]

        oi = int(q.get("oi", 0) or 0)
        vol = int(q.get("volume", 0) or 0)
        iv = float(q.get("implied_volatility", 0) or 0)
        ltp = float(q.get("last_price", 0) or 0)

        prev_close = float(ohlc.get("close", 0) or 0)
        net_chg = round(ltp - prev_close, 2) if prev_close else 0.0

        bid_p = float(buy[0].get("price", 0) or 0)
        bid_q = int(buy[0].get("quantity", 0) or 0)
        ask_p = float(sell[0].get("price", 0) or 0)
        ask_q = int(sell[0].get("quantity", 0) or 0)

        rows.append({
            "Strike": strike,
            "Type": otype,
            "OI": oi,
            "Volume": vol,
            "IV": round(iv, 2),
            "LTP": ltp,
            "Net_Chg": net_chg,
            "Bid_Qty": bid_q,
            "Bid_Price": bid_p,
            "Ask_Price": ask_p,
            "Ask_Qty": ask_q,
            "Open": float(ohlc.get("open", 0) or 0),
            "High": float(ohlc.get("high", 0) or 0),
            "Low": float(ohlc.get("low", 0) or 0),
            "Close": prev_close,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        print("⚠️ Option chain empty (quotes missing).")
        return None

    return df, spot, dte


# ==============================================================================
# 📌 Change-in-OI snapshot (so Chg_in_OI works)
# ==============================================================================
def load_last_oi_snapshot() -> dict:
    if os.path.exists(SNAPSHOT_OI_FILE):
        try:
            with open(SNAPSHOT_OI_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_last_oi_snapshot(data: dict) -> None:
    with open(SNAPSHOT_OI_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def apply_chg_in_oi(symbol: str, expiry, df: pd.DataFrame) -> pd.DataFrame:
    expiry_str = str(expiry)[:10]
    snap_key = f"{symbol}|{expiry_str}"

    snap = load_last_oi_snapshot()
    prev_map = snap.get(snap_key, {})  # "Strike_Type" -> OI

    df = df.copy()
    chg_list = []
    curr_map = {}

    for _, r in df.iterrows():
        k = f"{float(r['Strike'])}_{r['Type']}"
        curr_oi = int(r["OI"])
        prev_oi = int(prev_map.get(k, curr_oi))
        chg_list.append(curr_oi - prev_oi)
        curr_map[k] = curr_oi

    df["Chg_in_OI"] = chg_list
    snap[snap_key] = curr_map
    save_last_oi_snapshot(snap)
    return df


# ==============================================================================
# 🧩 BUILD NSE-LAYOUT (merge CE + PE per strike)
# ==============================================================================
def build_chain_dataframe(df: pd.DataFrame, spot: float | None):
    ce = df[df["Type"] == "CE"].copy()
    pe = df[df["Type"] == "PE"].copy()

    # Rename to NSE-layout expected column names
    ce = ce.rename(columns={
        "OI": "CE_OI",
        "Chg_in_OI": "CE_Chg_in_OI",
        "Volume": "CE_Volume",
        "IV": "CE_IV",
        "LTP": "CE_LTP",
        "Net_Chg": "CE_Net_Chg",
        "Bid_Qty": "CE_Bid_Qty",
        "Bid_Price": "CE_Bid_Price",
        "Ask_Price": "CE_Ask_Price",
        "Ask_Qty": "CE_Ask_Qty",
        "Open": "CE_Open",
        "High": "CE_High",
        "Low": "CE_Low",
        "Close": "CE_Close",
    })

    pe = pe.rename(columns={
        "OI": "PE_OI",
        "Chg_in_OI": "PE_Chg_in_OI",
        "Volume": "PE_Volume",
        "IV": "PE_IV",
        "LTP": "PE_LTP",
        "Net_Chg": "PE_Net_Chg",
        "Bid_Qty": "PE_Bid_Qty",
        "Bid_Price": "PE_Bid_Price",
        "Ask_Price": "PE_Ask_Price",
        "Ask_Qty": "PE_Ask_Qty",
        "Open": "PE_Open",
        "High": "PE_High",
        "Low": "PE_Low",
        "Close": "PE_Close",
    })

    # Keep only Strike + renamed columns
    ce_keep = ["Strike"] + CE_COLS
    pe_keep = ["Strike"] + PE_COLS

    for c in ce_keep:
        if c not in ce.columns:
            ce[c] = 0
    for c in pe_keep:
        if c not in pe.columns:
            pe[c] = 0

    ce = ce[ce_keep]
    pe = pe[pe_keep]

    merged = ce.merge(pe, on="Strike", how="outer").sort_values("Strike").reset_index(drop=True)

    atm_idx = None
    if spot:
        merged["_d"] = (merged["Strike"] - float(spot)).abs()
        atm_idx = int(merged["_d"].idxmin())
        merged.drop(columns=["_d"], inplace=True)

    return merged, atm_idx


# ==============================================================================
# 📤 EXPORT FOR POWER QUERY (CSV)
# ==============================================================================
def export_powerquery_nse_layout(df: pd.DataFrame, spot, symbol: str, expiry, dte: int):
    os.makedirs(BASE_FOLDER, exist_ok=True)

    merged, atm_idx = build_chain_dataframe(df, spot)

    # Ensure all expected columns exist
    for col in ALL_COLS:
        if col not in merged.columns:
            merged[col] = 0

    merged = merged[ALL_COLS].copy()

    # ISO format for Power Query stability
    ts_iso = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    expiry_iso = str(expiry)[:10]

    merged.insert(0, "timestamp", ts_iso)
    merged.insert(1, "symbol", symbol)
    merged.insert(2, "expiry", expiry_iso)
    merged.insert(3, "spot", float(spot) if spot is not None else 0.0)
    merged.insert(4, "dte", int(dte))
    merged.insert(5, "atm_row", 0)

    if atm_idx is not None:
        merged.loc[atm_idx, "atm_row"] = 1

    merged = merged.fillna(0)
    merged.to_csv(CSV_NSE_LAYOUT, index=False)

    # Dashboard file
    ce_df = df[df["Type"] == "CE"]
    pe_df = df[df["Type"] == "PE"]

    total_ce_oi = int(ce_df["OI"].sum()) if not ce_df.empty else 0
    total_pe_oi = int(pe_df["OI"].sum()) if not pe_df.empty else 0
    total_ce_vol = int(ce_df["Volume"].sum()) if not ce_df.empty else 0
    total_pe_vol = int(pe_df["Volume"].sum()) if not pe_df.empty else 0

    pcr = round(total_pe_oi / max(total_ce_oi, 1), 4)
    support = float(pe_df.sort_values("OI", ascending=False).iloc[0]["Strike"]) if not pe_df.empty else 0.0
    resistance = float(ce_df.sort_values("OI", ascending=False).iloc[0]["Strike"]) if not ce_df.empty else 0.0

    dash = pd.DataFrame([{
        "timestamp": ts_iso,
        "symbol": symbol,
        "expiry": expiry_iso,
        "spot": float(spot) if spot is not None else 0.0,
        "dte": int(dte),
        "total_ce_oi": total_ce_oi,
        "total_pe_oi": total_pe_oi,
        "total_ce_vol": total_ce_vol,
        "total_pe_vol": total_pe_vol,
        "pcr": pcr,
        "support": support,
        "resistance": resistance,
    }])
    dash.to_csv(CSV_DASHBOARD, index=False)

    print("✅ Power Query CSV updated:")
    print(f"   - {CSV_NSE_LAYOUT}")
    print(f"   - {CSV_DASHBOARD}")


# ==============================================================================
# 🔁 LIVE MODE
# ==============================================================================
def live_mode(kite: KiteConnect, instruments: pd.DataFrame, symbol: str, expiry, interval_sec: int):
    print(f"\n🔴 LIVE MODE STARTED — Update CSV every {interval_sec}s")
    print("   Excel Power Query ko Query Properties me 'Refresh every 1 minute' set karein.")
    print("   Stop: Ctrl+C\n")

    run = 0
    while True:
        run += 1
        try:
            print("\n" + "─" * 50)
            print(f"🔄 Refresh #{run} — {datetime.now().strftime('%H:%M:%S')}")

            result = fetch_option_chain(kite, instruments, symbol, expiry)
            if result is None:
                print("⚠️ No data. Waiting...")
                time.sleep(interval_sec)
                continue

            df, spot, dte = result
            df = apply_chg_in_oi(symbol, expiry, df)
            export_powerquery_nse_layout(df, spot, symbol, expiry, dte)

            time.sleep(interval_sec)

        except KeyboardInterrupt:
            print("\n🛑 Live mode stopped by user.")
            break
        except Exception as e:
            print(f"❌ Error: {e}")
            time.sleep(10)


# ==============================================================================
# 🚀 MAIN
# ==============================================================================
def main():
    os.makedirs(BASE_FOLDER, exist_ok=True)

    print("═" * 60)
    print("  📊  ZERODHA OI ANALYSIS — POWER QUERY MODE (NO GREEKS)")
    print("═" * 60)

    kite = kite_login_from_excel()
    kite.timeout = 30  # network stability

    instruments = download_instruments(kite, force=False)

    symbol = show_stock_selector(instruments)
    expiry = show_expiry_selector(instruments, symbol)

    print("\n📌 MODE SELECT:")
    print("   1. Single Snapshot (CSV update once)")
    print("   2. Live Mode (CSV update every X seconds)")
    mode = input("\nMode (1/2) [2]: ").strip() or "2"

    if mode == "1":
        result = fetch_option_chain(kite, instruments, symbol, expiry)
        if result is None:
            print("❌ Data fetch failed.")
            return
        df, spot, dte = result
        df = apply_chg_in_oi(symbol, expiry, df)
        export_powerquery_nse_layout(df, spot, symbol, expiry, dte)
        print("\n✅ Done. Excel Power Query refresh karega.")
        return

    interval = input(f"Refresh interval seconds (default {DEFAULT_REFRESH_SECONDS}): ").strip()
    interval_sec = int(interval) if interval.isdigit() else DEFAULT_REFRESH_SECONDS
    live_mode(kite, instruments, symbol, expiry, interval_sec)


if __name__ == "__main__":
    main()
