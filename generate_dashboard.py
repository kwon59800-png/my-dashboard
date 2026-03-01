"""
Dashboard Generator
===================
daily_sp500.xlsx + chart_data.json 을 읽어
portfolio-dashboard-live.html 기반으로 dashboard_today.html 을 자동 생성.
"""

import json
import math
import re
import sys
import logging
from pathlib import Path
from datetime import datetime

import pandas as pd

SCRIPT_DIR = Path(__file__).parent
EXCEL_FILE = SCRIPT_DIR / "daily_sp500.xlsx"
CHART_FILE = SCRIPT_DIR / "chart_data.json"
ANALYST_FILE = SCRIPT_DIR / "analyst_targets.json"
TEMPLATE_FILE = SCRIPT_DIR / "portfolio-dashboard-live.html"
MERGED_TEMPLATE = SCRIPT_DIR / "대시보드_PC_v3_프리장실시간_로고 (1).html"
OUTPUT_FILE = SCRIPT_DIR / "dashboard_today.html"
MOBILE_OUTPUT = SCRIPT_DIR / "dashboard_mobile.html"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# ── 포트폴리오 정의 ──────────────────────────────
MYPORT = ["NVDA", "TSM", "MU", "CEG", "SOXL", "IONQ", "CRCL"]
TC = {"NVDA": "#76B900", "TSM": "#E4002B", "MU": "#FDB515", "CEG": "#00A3E0",
      "SOXL": "#FF6B35", "IONQ": "#7B2D8E", "CRCL": "#3498DB"}
NAMES = {"NVDA": "NVIDIA", "TSM": "Taiwan Semi", "MU": "Micron",
         "CEG": "Constellation Energy", "SOXL": "Direxion Semi 3X",
         "IONQ": "IonQ", "CRCL": "Circle Internet"}
PORT_SEC = {"NVDA": "정보기술", "TSM": "정보기술", "MU": "정보기술",
            "CEG": "유틸리티", "SOXL": "레버리지ETF", "IONQ": "정보기술", "CRCL": "금융"}
ALLOC_W = {"NVDA": 25, "TSM": 18, "MU": 18, "CEG": 12, "SOXL": 12, "IONQ": 8, "CRCL": 7}

# ── GICS 섹터 정의 ──────────────────────────────
SECTORS = [
    {"id": "IT", "nm": "정보기술", "etf": "XLK", "co": "#3b82f6", "gics": "Information Technology"},
    {"id": "HC", "nm": "헬스케어", "etf": "XLV", "co": "#10b981", "gics": "Health Care"},
    {"id": "FN", "nm": "금융", "etf": "XLF", "co": "#f59e0b", "gics": "Financials"},
    {"id": "CD", "nm": "경기소비재", "etf": "XLY", "co": "#ef4444", "gics": "Consumer Discretionary"},
    {"id": "CS", "nm": "커뮤니케이션", "etf": "XLC", "co": "#8b5cf6", "gics": "Communication Services"},
    {"id": "IN", "nm": "산업재", "etf": "XLI", "co": "#6366f1", "gics": "Industrials"},
    {"id": "CP", "nm": "필수소비재", "etf": "XLP", "co": "#14b8a6", "gics": "Consumer Staples"},
    {"id": "EN", "nm": "에너지", "etf": "XLE", "co": "#f97316", "gics": "Energy"},
    {"id": "UT", "nm": "유틸리티", "etf": "XLU", "co": "#06b6d4", "gics": "Utilities"},
    {"id": "RE", "nm": "부동산", "etf": "XLRE", "co": "#ec4899", "gics": "Real Estate"},
    {"id": "MT", "nm": "소재", "etf": "XLB", "co": "#84cc16", "gics": "Materials"},
]
SECTOR_KR = {s["gics"]: s["nm"] for s in SECTORS}
SECTOR_ETFS = [s["etf"] for s in SECTORS]

# ── 포트폴리오 매수 정보 (평단가/수량) ──
PORT_DETAIL = {
    "NVDA": {"qty": 114, "avg": 183.7547},
    "TSM":  {"qty": 87,  "avg": 293.1331},
    "MU":   {"qty": 78,  "avg": 284.9344},
    "CEG":  {"qty": 304, "avg": 296.5757},
    "SOXL": {"qty": 80,  "avg": 62.2500},
    "IONQ": {"qty": 64,  "avg": 45.9200},
    "CRCL": {"qty": 56,  "avg": 58.0000},
}


def fetch_fx_rate():
    """USD/KRW 환율 조회 (여러 소스 폴백)."""
    import urllib.request
    apis = [
        "https://open.er-api.com/v6/latest/USD",
        "https://api.exchangerate-api.com/v4/latest/USD",
    ]
    for url in apis:
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read().decode())
                rate = data.get("rates", {}).get("KRW", 0)
                if rate > 0:
                    log.info(f"FX rate fetched: USD/KRW = {rate:.2f}")
                    return round(rate, 2)
        except Exception as e:
            log.warning(f"FX fetch failed ({url}): {e}")
    log.warning("All FX APIs failed, using fallback 1440.86")
    return 1440.86


NAMES_KR = {
    "CEG": "컨스텔레이션 에너지", "MU": "마이크론 테크놀로지", "TSM": "TSMC ADR",
    "NVDA": "엔비디아", "SOXL": "Direxion Semi 3×", "CRCL": "서클 인터넷 그룹",
    "IONQ": "아이온큐",
}
ACCENT = {
    "CEG": "#5b8ff9", "MU": "#f05c5c", "TSM": "#f5a623", "NVDA": "#22c97a",
    "SOXL": "#a78bfa", "CRCL": "#60a5fa", "IONQ": "#fb923c",
}
# 기타 자산 (크립토)
OTHER_ASSETS = [
    {"ticker": "BTC", "name": "비트코인", "label": "0.1503 BTC", "qty": 0.15031317, "avg": 18990000, "accent": "#f7931a"},
    {"ticker": "XRP", "name": "리플", "label": "360.63 XRP", "qty": 360.63, "avg": 999500, "accent": "#25a3d9"},
]

# ── 검색용 추가 티커 (S&P500 외 주요 종목/ETF) ──
EXTRA_TICKERS = {
    "SPY":{"n":"SPDR S&P 500 ETF","s":"ETF"}, "QQQ":{"n":"Invesco QQQ (NASDAQ-100)","s":"ETF"},
    "DIA":{"n":"SPDR Dow Jones ETF","s":"ETF"}, "IWM":{"n":"iShares Russell 2000","s":"ETF"},
    "VOO":{"n":"Vanguard S&P 500","s":"ETF"}, "VTI":{"n":"Vanguard Total Market","s":"ETF"},
    "VGT":{"n":"Vanguard Info Tech ETF","s":"ETF"}, "VUG":{"n":"Vanguard Growth ETF","s":"ETF"},
    "TQQQ":{"n":"ProShares UltraPro QQQ 3×","s":"레버리지ETF"}, "SQQQ":{"n":"ProShares Short QQQ 3×","s":"인버스ETF"},
    "UPRO":{"n":"ProShares UltraPro S&P500 3×","s":"레버리지ETF"}, "SPXU":{"n":"ProShares Short S&P500 3×","s":"인버스ETF"},
    "SOXS":{"n":"Direxion Semi Bear 3×","s":"인버스ETF"}, "LABU":{"n":"Direxion Biotech Bull 3×","s":"레버리지ETF"},
    "TNA":{"n":"Direxion Small Cap Bull 3×","s":"레버리지ETF"}, "FNGU":{"n":"MicroSectors FANG+ Bull 3×","s":"레버리지ETF"},
    "ARKK":{"n":"ARK Innovation ETF","s":"테마ETF"}, "ARKG":{"n":"ARK Genomic Revolution","s":"테마ETF"},
    "SMH":{"n":"VanEck Semiconductor ETF","s":"반도체ETF"}, "SOXX":{"n":"iShares Semiconductor ETF","s":"반도체ETF"},
    "XBI":{"n":"SPDR S&P Biotech ETF","s":"바이오ETF"}, "KWEB":{"n":"KraneShares China Internet","s":"중국ETF"},
    "BOTZ":{"n":"Global X Robotics & AI ETF","s":"AI/로봇ETF"}, "LIT":{"n":"Global X Lithium & Battery","s":"배터리ETF"},
    "TAN":{"n":"Invesco Solar ETF","s":"신재생ETF"}, "ICLN":{"n":"iShares Clean Energy","s":"신재생ETF"},
    "GLD":{"n":"SPDR Gold Trust","s":"원자재ETF"}, "SLV":{"n":"iShares Silver Trust","s":"원자재ETF"},
    "USO":{"n":"United States Oil Fund","s":"원자재ETF"}, "TLT":{"n":"iShares 20+ Year Treasury","s":"채권ETF"},
    "HYG":{"n":"iShares High Yield Bond","s":"채권ETF"}, "LQD":{"n":"iShares Inv Grade Bond","s":"채권ETF"},
    "BND":{"n":"Vanguard Total Bond Market","s":"채권ETF"}, "EEM":{"n":"iShares MSCI Emerging Mkts","s":"신흥국ETF"},
    "VNQ":{"n":"Vanguard Real Estate ETF","s":"부동산ETF"},
    "XLE":{"n":"Energy Select SPDR","s":"에너지"}, "XLF":{"n":"Financial Select SPDR","s":"금융"},
    "XLK":{"n":"Technology Select SPDR","s":"정보기술"}, "XLV":{"n":"Health Care Select SPDR","s":"헬스케어"},
    "XLY":{"n":"Consumer Discret SPDR","s":"경기소비재"}, "XLC":{"n":"Communication Svc SPDR","s":"커뮤니케이션"},
    "XLI":{"n":"Industrial Select SPDR","s":"산업재"}, "XLP":{"n":"Consumer Staples SPDR","s":"필수소비재"},
    "XLU":{"n":"Utilities Select SPDR","s":"유틸리티"}, "XLRE":{"n":"Real Estate Select SPDR","s":"부동산"},
    "XLB":{"n":"Materials Select SPDR","s":"소재"},
    "IONQ":{"n":"IonQ","s":"정보기술"}, "CRCL":{"n":"Circle Internet Group","s":"금융"},
    "PLTR":{"n":"Palantir Technologies","s":"정보기술"}, "MSTR":{"n":"MicroStrategy","s":"정보기술"},
    "COIN":{"n":"Coinbase Global","s":"금융"}, "RKLB":{"n":"Rocket Lab USA","s":"산업재"},
    "SOFI":{"n":"SoFi Technologies","s":"금융"}, "RIVN":{"n":"Rivian Automotive","s":"경기소비재"},
    "LCID":{"n":"Lucid Group","s":"경기소비재"}, "NIO":{"n":"NIO Inc ADR","s":"경기소비재"},
    "BABA":{"n":"Alibaba Group ADR","s":"경기소비재"}, "PDD":{"n":"PDD Holdings ADR","s":"경기소비재"},
    "SE":{"n":"Sea Limited ADR","s":"커뮤니케이션"}, "SNAP":{"n":"Snap Inc","s":"커뮤니케이션"},
    "PINS":{"n":"Pinterest","s":"커뮤니케이션"}, "U":{"n":"Unity Software","s":"정보기술"},
    "DKNG":{"n":"DraftKings","s":"경기소비재"}, "HOOD":{"n":"Robinhood Markets","s":"금융"},
    "AFRM":{"n":"Affirm Holdings","s":"정보기술"}, "UPST":{"n":"Upstart Holdings","s":"금융"},
    "AI":{"n":"C3.ai","s":"정보기술"}, "RGTI":{"n":"Rigetti Computing","s":"정보기술"},
    "QUBT":{"n":"Quantum Computing","s":"정보기술"}, "MARA":{"n":"Marathon Digital","s":"정보기술"},
    "RIOT":{"n":"Riot Platforms","s":"정보기술"}, "CLSK":{"n":"CleanSpark","s":"정보기술"},
}


def build_all_tickers_info(df):
    """S&P500 + 포트폴리오 + 섹터ETF + EXTRA_TICKERS 통합 사전 생성."""
    info = {}
    for _, row in df.iterrows():
        tk = str(row["Ticker"]).strip()
        sec_kr = SECTOR_KR.get(str(row["Sector"]), str(row["Sector"])) if pd.notna(row.get("Sector")) else "기타"
        nm = str(row["Company"]).strip() if pd.notna(row.get("Company")) else tk
        info[tk] = {"n": nm, "s": sec_kr}
    for s in SECTORS:
        if s["etf"] not in info:
            info[s["etf"]] = {"n": s["nm"] + " ETF", "s": s["nm"]}
    for tk in MYPORT:
        if tk not in info:
            info[tk] = {"n": NAMES.get(tk, tk), "s": PORT_SEC.get(tk, "기타")}
    for tk, v in EXTRA_TICKERS.items():
        if tk not in info:
            info[tk] = v
    log.info(f"ALL_TICKERS_INFO: {len(info)} tickers total")
    return info


def build_all_tickers_js(info_dict):
    """ALL_TICKERS_INFO를 JS const 문자열로 변환."""
    pairs = []
    for tk in sorted(info_dict.keys()):
        v = info_dict[tk]
        n = v["n"].replace("'", "\\'")
        s = v["s"].replace("'", "\\'")
        pairs.append(f"'{tk}':{{n:'{n}',s:'{s}'}}")
    return "const ALL_TICKERS_INFO={" + ",".join(pairs) + "};"


# ── data.js / tickers.js 생성 ─────────────────────────────

def write_data_js(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas, ut_init):
    """SA 모듈 핵심 데이터를 data.js (const SA_DATA=...)로 저장 — file:// CORS 회피."""
    # 나머지 SP5 티커 (SMAP에 없는 것)
    smap_tickers = set()
    for s in smap:
        smap_tickers.update(s["st"])
    remaining = [t for t in all_sp5_tickers
                 if t not in smap_tickers and t not in MYPORT and t not in SECTOR_ETFS]

    data = {
        "TD": td,
        "SP5_INFO": sp5_info,
        "BETA_MAP": betas or {},
        "SMAP": smap,
        "FC": fc,
        "EV": ev,
        "MYPORT": MYPORT,
        "TC": TC,
        "NAMES": NAMES,
        "PORT_SEC": PORT_SEC,
        "ALLOC_W": ALLOC_W,
        "EXTRA_SP5": remaining,
        "ut_init": ut_init,
        "data_date": data_date,
    }
    out_path = SCRIPT_DIR / "data.js"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("const SA_DATA = ")
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")
    size_kb = out_path.stat().st_size / 1024
    log.info(f"Saved {out_path.name} ({size_kb:.0f} KB)")


def fetch_all_us_tickers():
    """나스닥 API에서 NYSE/NASDAQ/AMEX 전체 상장 종목 가져오기."""
    import urllib.request
    tickers = {}
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    for exchange in ["nasdaq", "nyse", "amex"]:
        url = f"https://api.nasdaq.com/api/screener/stocks?tableType=traded&exchange={exchange}&download=true"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                raw = json.loads(resp.read().decode("utf-8"))
            rows = raw.get("data", {}).get("rows", [])
            for row in rows:
                sym = str(row.get("symbol", "")).strip()
                name = str(row.get("name", "")).strip()
                sector = str(row.get("sector", "")).strip()
                if not sym or len(sym) > 5 or "." in sym or "^" in sym or "/" in sym:
                    continue  # 우선주·워런트 등 특수 종목 제외
                tickers[sym] = {"n": name, "s": sector if sector else "기타"}
            log.info(f"  {exchange.upper()}: {len(rows)} rows → {len([r for r in rows if str(r.get('symbol','')).strip() in tickers])} valid")
        except Exception as e:
            log.warning(f"  {exchange.upper()} fetch failed: {e}")
    log.info(f"Total US tickers fetched: {len(tickers)}")
    return tickers


def write_tickers_js(df):
    """S&P500(엑셀) + 전체 미국주식(나스닥 API) 통합 티커 사전 → tickers.js."""
    # 1) 기존 S&P500 + EXTRA_TICKERS (섹터 한국어명 포함)
    base = build_all_tickers_info(df)
    base_count = len(base)

    # 2) 나스닥 API에서 전체 미국주식 머지
    try:
        all_us = fetch_all_us_tickers()
        for tk, v in all_us.items():
            if tk not in base:  # 기존(S&P500 한국어 섹터) 데이터 우선
                base[tk] = v
        log.info(f"Tickers merged: {base_count} base + {len(all_us)} US → {len(base)} total")
    except Exception as e:
        log.warning(f"US ticker fetch failed: {e} (using {base_count} base tickers only)")

    # 3) JS 파일 저장 — file:// CORS 회피
    out_path = SCRIPT_DIR / "tickers.js"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("const ALL_TICKERS_INFO = ")
        json.dump(base, f, ensure_ascii=False, separators=(",", ":"))
        f.write(";\n")
    size_kb = out_path.stat().st_size / 1024
    log.info(f"Saved {out_path.name} ({size_kb:.0f} KB, {len(base)} tickers)")


def build_asset_update(df, fc, analyst, fx_rate):
    """mainData / otherData 의 최신 JS 초기값을 생성."""
    # mainData: 포트폴리오 종목 순서 = 평가액 내림차순
    items = []
    for ticker in MYPORT:
        d = PORT_DETAIL.get(ticker, {})
        qty = d.get("qty", 0)
        avg = d.get("avg", 0)
        # 현재가 — 애널리스트 파일 → 엑셀 Close 순
        at = analyst.get(ticker, {})
        cp_at = at.get("currentPrice")
        rows = df[df["Ticker"] == ticker]
        if cp_at:
            cp = float(cp_at)
        elif len(rows):
            cp = float(rows.iloc[0]["Close"])
        else:
            cp = avg  # fallback
        perf = ((cp - avg) / avg * 100) if avg > 0 else 0
        # 목표가 — fc에서 추출
        target = 0
        for f_ in fc:
            if f_["t"] == ticker:
                target = f_.get("at_avg") or f_.get("hi") or 0
                break
        items.append({
            "ticker": ticker, "name": NAMES_KR.get(ticker, ticker),
            "qty": qty, "cp": round(cp, 2), "avg": round(avg, 4),
            "target": round(target), "perf": round(perf, 2),
            "accent": ACCENT.get(ticker, "#888"),
            "eval": qty * cp,  # 정렬용
        })
    # 평가액 내림차순 정렬
    items.sort(key=lambda x: x["eval"], reverse=True)

    lines = []
    for a in items:
        lines.append(
            f"  {{ ticker:'{a['ticker']}', name:'{a['name']}', "
            f"qty:{a['qty']}, currentPrice:{a['cp']}, avgPrice:{a['avg']}, "
            f"targetPrice:{a['target']}, performance:{a['perf']}, "
            f"changePercent:0, changeAbs:0, prevClose:0, "
            f"preMarketPrice:null, preMarketChg:null, preMarketChgAbs:null, "
            f"postMarketPrice:null, postMarketChg:null, marketState:'CLOSED', "
            f"accent:'{a['accent']}' }},")

    main_js = "var mainData = [\n" + "\n".join(lines) + "\n];"

    # otherData: 크립토 (BTC, XRP) — 최신 KRW 가격 가져오기
    other_lines = []
    for oa in OTHER_ASSETS:
        tk = oa["ticker"]
        # 크립토 현재 KRW 단가 — fetch 시도
        unit_price = _fetch_crypto_krw(tk)
        if unit_price <= 0:
            unit_price = oa["avg"]  # fallback (총액으로 쓰이던 값)
        # otherData.currentPrice = 총 평가액 (calcStats에서 isUSD:false → currentPrice 그대로 합산)
        total_val = round(unit_price * oa["qty"])
        avg_total = round(oa["avg"] * oa["qty"]) if oa["avg"] < 100000 else oa["avg"]
        # 기존 avgPrice가 총액 형태(18990000)이면 그대로, 단가 형태면 qty 곱산
        avg_val = oa["avg"]
        perf = ((total_val - avg_val) / avg_val * 100) if avg_val > 0 else 0
        other_lines.append(
            f"  {{ ticker:'{tk}', name:'{oa['name']}', label:'{oa['label']}', "
            f"qty:{oa['qty']}, currentPrice:{total_val}, "
            f"avgPrice:{oa['avg']}, performance:{round(perf, 2)}, "
            f"changePercent:0, isUSD:false, accent:'{oa['accent']}' }},")

    other_js = "var otherData = [\n" + "\n".join(other_lines) + "\n];"
    return main_js, other_js


def _fetch_crypto_krw(symbol):
    """CoinGecko/Upbit 등에서 크립토 KRW 가격 조회."""
    import urllib.request
    coin_map = {"BTC": "bitcoin", "XRP": "ripple", "ETH": "ethereum"}
    cg_id = coin_map.get(symbol, symbol.lower())
    # CoinGecko
    try:
        url = f"https://api.coingecko.com/api/v3/simple/price?ids={cg_id}&vs_currencies=krw"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())
            price = data.get(cg_id, {}).get("krw", 0)
            if price > 0:
                log.info(f"Crypto {symbol} = ₩{price:,.0f} (CoinGecko)")
                return price
    except Exception as e:
        log.warning(f"CoinGecko {symbol} failed: {e}")
    # Upbit fallback
    try:
        url = f"https://api.upbit.com/v1/ticker?markets=KRW-{symbol}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())
            if data and len(data):
                price = data[0].get("trade_price", 0)
                if price > 0:
                    log.info(f"Crypto {symbol} = ₩{price:,.0f} (Upbit)")
                    return price
    except Exception as e:
        log.warning(f"Upbit {symbol} failed: {e}")
    return 0


def compute_betas(chart, benchmark="SPY", min_days=10):
    """SPY 대비 각 종목의 Beta를 계산 (공분산/분산 방식)."""
    betas = {}
    if benchmark not in chart or len(chart[benchmark]) < min_days:
        log.warning(f"Benchmark {benchmark} not in chart data or insufficient days")
        return betas
    spy_entries = chart[benchmark]
    spy_returns = []
    for i in range(1, len(spy_entries)):
        if spy_entries[i - 1]["c"] > 0:
            spy_returns.append((spy_entries[i]["c"] - spy_entries[i - 1]["c"]) / spy_entries[i - 1]["c"])
    if len(spy_returns) < 5:
        return betas
    spy_avg = sum(spy_returns) / len(spy_returns)
    spy_var = sum((r - spy_avg) ** 2 for r in spy_returns) / len(spy_returns)
    if spy_var == 0:
        return betas
    for ticker, entries in chart.items():
        if ticker == benchmark or len(entries) < min_days:
            continue
        # 날짜 매칭용 딕셔너리
        tk_map = {e["d"]: e["c"] for e in entries}
        tk_returns = []
        matched_spy = []
        for i in range(1, len(spy_entries)):
            d = spy_entries[i]["d"]
            d_prev = spy_entries[i - 1]["d"]
            if d in tk_map and d_prev in tk_map and tk_map[d_prev] > 0:
                tk_returns.append((tk_map[d] - tk_map[d_prev]) / tk_map[d_prev])
                matched_spy.append(spy_returns[i - 1] if i - 1 < len(spy_returns) else 0)
        if len(tk_returns) < 5:
            continue
        tk_avg = sum(tk_returns) / len(tk_returns)
        ms_avg = sum(matched_spy) / len(matched_spy)
        ms_var = sum((r - ms_avg) ** 2 for r in matched_spy) / len(matched_spy)
        if ms_var == 0:
            continue
        cov = sum((tk_returns[i] - tk_avg) * (matched_spy[i] - ms_avg) for i in range(len(tk_returns))) / len(tk_returns)
        beta = cov / ms_var
        betas[ticker] = round(beta, 2)
    log.info(f"Beta computed for {len(betas)} tickers")
    return betas


def compute_supply(entries):
    """차트 데이터에서 수급 지표를 계산."""
    if len(entries) < 5:
        return {"vc": 0, "co": 0, "j": "중립"}
    volumes = [e["v"] for e in entries]
    closes = [e["c"] for e in entries]
    r5 = sum(volumes[-5:]) / 5
    prev = volumes[:-5] if len(volumes) > 5 else volumes[:1]
    p_avg = sum(prev) / max(len(prev), 1)
    vc = (r5 - p_avg) / p_avg if p_avg > 0 else 0
    # 가격-거래량 상관관계
    if len(closes) >= 3:
        pch = [(closes[i] - closes[i - 1]) / closes[i - 1] for i in range(1, len(closes))]
        vols = volumes[1:]
        ap = sum(pch) / len(pch)
        av = sum(vols) / len(vols)
        cv = sum((pch[i] - ap) * (vols[i] - av) for i in range(len(pch)))
        vp = sum((p - ap) ** 2 for p in pch)
        vv = sum((v - av) ** 2 for v in vols)
        co = cv / math.sqrt(vp * vv) if vp > 0 and vv > 0 else 0
    else:
        co = 0
    # 판정
    if vc > 0.3 and co > 0.2:
        j = "강한 매수세"
    elif vc > 0.1 and co > 0:
        j = "매수 우위"
    elif vc > 0.2 and co < -0.1:
        j = "강한 매도세"
    elif vc > 0 and co < 0:
        j = "매도 우위"
    else:
        j = "중립"
    return {"vc": round(vc, 3), "co": round(co, 3), "j": j}


def build_td(df, chart):
    """모든 티커의 TD 데이터 구축."""
    td = {}
    ts = int(datetime.now().timestamp() * 1000)
    # 코어 티커 (포트폴리오 + ETF) → 풀 차트, 나머지 → 5엔트리
    core = set(MYPORT + SECTOR_ETFS)
    for sec in SECTORS:
        if sec["gics"] in df["Sector"].values:
            top5 = df[df["Sector"] == sec["gics"]].nlargest(5, "Composite_Score")["Ticker"].tolist()
            core.update(top5)

    for ticker, entries in chart.items():
        is_core = ticker in core
        n_entries = 20 if is_core else 5
        chart_slice = entries[-n_entries:] if len(entries) >= n_entries else entries

        # 차트 배열 (최소 3개 보장)
        daily = [{"d": e["d"], "c": e["c"], "a": e["a"], "b": e["b"]} for e in chart_slice]
        if len(daily) < 3:
            while len(daily) < 3:
                daily.insert(0, daily[0].copy())

        # 1개월 수익률
        r = (entries[-1]["c"] - entries[0]["c"]) / entries[0]["c"] if len(entries) >= 2 else 0

        # 수급 지표
        supply = compute_supply(entries)
        supply["t"] = ticker

        # 종목 정보
        row_match = df[df["Ticker"] == ticker]
        if len(row_match) > 0:
            row = row_match.iloc[0]
            name = str(row["Company"]) if pd.notna(row.get("Company")) else ticker
            sector = SECTOR_KR.get(str(row["Sector"]), str(row["Sector"])) if pd.notna(row.get("Sector")) else "기타"
            close = float(row["Close"])
            date_str = str(row["Date"])
        else:
            name = NAMES.get(ticker, ticker)
            sector = PORT_SEC.get(ticker, "기타")
            # ETF인 경우 이름 설정
            for s in SECTORS:
                if s["etf"] == ticker:
                    name = s["nm"] + " ETF"
                    sector = s["nm"]
                    break
            close = entries[-1]["c"] if entries else 0
            date_str = datetime.now().strftime("%Y-%m-%d")

        td[ticker] = {
            "daily": daily,
            "summary": {"t": ticker, "n": name, "s": sector, "p": close, "r": round(r, 4), "lastDate": date_str},
            "supply": supply,
            "ts": ts,
        }
    return td


def build_smap(df):
    """SMAP (섹터별 대표 종목) 구축."""
    smap = []
    for sec in SECTORS:
        sec_stocks = df[df["Sector"] == sec["gics"]]
        if len(sec_stocks) > 0:
            top5 = sec_stocks.nlargest(5, "Composite_Score")["Ticker"].tolist()
        else:
            top5 = []
        smap.append({"id": sec["id"], "nm": sec["nm"], "etf": sec["etf"], "co": sec["co"], "st": top5})
    return smap


def build_sp5_info(df):
    """SP5_INFO (티커→이름/섹터 맵) 구축."""
    info = {}
    for _, row in df.iterrows():
        tk = row["Ticker"]
        sec_kr = SECTOR_KR.get(str(row["Sector"]), str(row["Sector"])) if pd.notna(row.get("Sector")) else "기타"
        nm = str(row["Company"]) if pd.notna(row.get("Company")) else tk
        info[tk] = {"n": nm, "sec": sec_kr}
    # ETF
    for s in SECTORS:
        info[s["etf"]] = {"n": s["nm"] + " ETF", "sec": s["nm"]}
    # 포트폴리오 (S&P500에 없는 것)
    for tk in MYPORT:
        if tk not in info:
            info[tk] = {"n": NAMES.get(tk, tk), "sec": PORT_SEC.get(tk, "기타")}
    return info


def build_fc(df, analyst=None):
    """FC (포트폴리오 전망 카드) 구축 - 애널리스트 목표가 반영."""
    fc = []
    analyst = analyst or {}
    for ticker in MYPORT:
        rows = df[df["Ticker"] == ticker]
        at = analyst.get(ticker, {})

        if len(rows) == 0:
            cp = at.get("currentPrice") or 0
            at_lo = at.get("targetLowPrice") or 0
            at_hi = at.get("targetHighPrice") or 0
            at_avg = at.get("targetMeanPrice") or at.get("targetMedianPrice") or 0
            at_n = at.get("numberOfAnalystOpinions") or 0
            rec = at.get("recommendationKey", "")
            # 애널리스트 데이터가 있으면 활용
            if at_avg > 0:
                lo = round(at_lo) if at_lo > 0 else round(at_avg * 0.9)
                hi = round(at_hi) if at_hi > 0 else round(at_avg * 1.1)
                ret_pct = ((at_avg - cp) / cp * 100) if cp > 0 else 0
                rec_kr = {"buy": "매수", "strong_buy": "강력매수", "hold": "보유",
                           "sell": "매도", "strong_sell": "강력매도"}.get(rec, "")
                cat = f"애널리스트 {at_n}명 · 평균목표 ${at_avg:.0f} ({ret_pct:+.1f}%)"
                if rec_kr:
                    cat += f" · {rec_kr}"
                g = "Speculative" if ticker in ("IONQ",) else ("High Risk" if ticker in ("SOXL",) else "N/A")
            else:
                lo, hi = 0, 0
                cat = "데이터 없음"
                g = "High Risk" if ticker == "SOXL" else ("Speculative" if ticker == "IONQ" else "N/A")
            fc.append({
                "t": ticker, "lo": lo, "hi": hi, "g": g, "cat": cat,
                "cp": round(cp, 2) if cp else 0,
                "at_lo": round(at_lo, 2) if at_lo else 0,
                "at_avg": round(at_avg, 2) if at_avg else 0,
                "at_hi": round(at_hi, 2) if at_hi else 0,
                "at_n": at_n,
            })
            log.info(f"  FC [{ticker}] cp=${cp:.2f} target_avg=${at_avg:.0f} lo=${lo} hi=${hi} ({at_n} analysts) [non-SP500]")
            continue

        row = rows.iloc[0]

        # 현재가: 애널리스트 데이터 currentPrice 우선 (최신), fallback은 엑셀 Close
        cp_analyst = at.get("currentPrice")
        close = float(cp_analyst) if cp_analyst else float(row["Close"])

        rsi = float(row["RSI_14"]) if pd.notna(row.get("RSI_14")) else 50
        cross = str(row.get("Cross_Signal", ""))

        # 애널리스트 목표가
        at_lo = at.get("targetLowPrice") or 0
        at_hi = at.get("targetHighPrice") or 0
        at_avg = at.get("targetMeanPrice") or at.get("targetMedianPrice") or 0
        at_n = at.get("numberOfAnalystOpinions") or 0
        rec = at.get("recommendationKey", "")

        # lo/hi: 애널리스트 데이터 우선, 없으면 기존 20일 고저 방식
        if at_avg > 0:
            lo = round(at_lo) if at_lo > 0 else round(at_avg * 0.9)
            hi = round(at_hi) if at_hi > 0 else round(at_avg * 1.1)
        else:
            high_20d = float(row["High_20D"]) if pd.notna(row.get("High_20D")) else close * 1.1
            low_20d = float(row["Low_20D"]) if pd.notna(row.get("Low_20D")) else close * 0.9
            lo = round(low_20d * 0.95)
            hi = round(high_20d * 1.05)

        # 등급 판정
        if ticker == "SOXL":
            g = "High Risk"
        elif ticker == "IONQ":
            g = "Speculative"
        elif cross == "GOLDEN_CROSS":
            g = "Recovery"
        elif rsi < 30:
            g = "Oversold"
        elif rsi < 50:
            g = "Bullish"
        elif rsi < 65:
            g = "Moderate"
        elif rsi < 80:
            g = "Cautious"
        else:
            g = "Overbought"

        chg = float(row["Daily_Chg_Pct"]) if pd.notna(row.get("Daily_Chg_Pct")) else 0

        # 카테고리 설명: 애널리스트 데이터 포함
        if at_avg > 0 and at_n > 0:
            ret_pct = ((at_avg - close) / close * 100) if close > 0 else 0
            rec_kr = {"buy": "매수", "strong_buy": "강력매수", "hold": "보유",
                       "sell": "매도", "strong_sell": "강력매도"}.get(rec, "")
            cat = f"애널리스트 {at_n}명 · 평균목표 ${at_avg:.0f} ({ret_pct:+.1f}%)"
            if rec_kr:
                cat += f" · {rec_kr}"
        else:
            cat = f"RSI {rsi:.0f} · 일간 {chg:+.1f}%"
            if cross:
                cat += f" · {cross}"

        fc.append({
            "t": ticker, "lo": lo, "hi": hi, "g": g, "cat": cat,
            "cp": round(close, 2),
            "at_lo": round(at_lo, 2) if at_lo else 0,
            "at_avg": round(at_avg, 2) if at_avg else 0,
            "at_hi": round(at_hi, 2) if at_hi else 0,
            "at_n": at_n,
        })
        log.info(f"  FC [{ticker}] cp=${close:.2f} target_avg=${at_avg:.0f} lo=${lo} hi=${hi} ({at_n} analysts)")
    return fc


def build_ev(df):
    """EV (이벤트 로그) 구축."""
    ev = []
    date_str = str(df.iloc[0]["Date"])
    mm_dd = date_str[5:].replace("-", "/")

    # 골든크로스
    golden = df[df["Cross_Signal"] == "GOLDEN_CROSS"]["Ticker"].tolist()
    for t in golden[:10]:
        ev.append({"d": mm_dd, "t": t, "e": "골든크로스 (MA5↑MA20)"})

    # 데드크로스
    dead = df[df["Cross_Signal"] == "DEAD_CROSS"]["Ticker"].tolist()
    for t in dead[:10]:
        ev.append({"d": mm_dd, "t": t, "e": "데드크로스 (MA5↓MA20)"})

    # 거래량 급증
    vol_spike = df[df["Vol_Ratio"] > 3].nlargest(5, "Vol_Ratio")
    for _, row in vol_spike.iterrows():
        ev.append({"d": mm_dd, "t": row["Ticker"],
                    "e": f"거래량 {row['Vol_Ratio']:.1f}x ({row['Daily_Chg_Pct']:+.1f}%)"})

    # RSI 극단
    if "RSI_14" in df.columns:
        rsi_hi = df[df["RSI_14"] > 80].nlargest(3, "RSI_14")
        for _, row in rsi_hi.iterrows():
            ev.append({"d": mm_dd, "t": row["Ticker"], "e": f"과매수 RSI {row['RSI_14']:.0f}"})
        rsi_lo = df[df["RSI_14"] < 20].nsmallest(3, "RSI_14")
        for _, row in rsi_lo.iterrows():
            ev.append({"d": mm_dd, "t": row["Ticker"], "e": f"과매도 RSI {row['RSI_14']:.0f}"})

    return ev


def build_ai_briefing(df, fc, ev, data_date):
    """AI 전략 브리핑 HTML 생성 — 전일 증시 이슈 + 포트폴리오 분석.

    generate_dashboard.py가 매일 7시에 실행될 때 이 함수가 호출되어
    대시보드 우측 패널의 AI 브리핑 영역에 주입됨.
    """
    today_str = datetime.now().strftime("%m/%d")

    # ── ① 주목 종목 (전일 이슈 기반) ──
    highlights = []
    for ticker in MYPORT:
        rows = df[df["Ticker"] == ticker]
        if len(rows) == 0:
            continue
        row = rows.iloc[0]
        chg = float(row["Daily_Chg_Pct"]) if pd.notna(row.get("Daily_Chg_Pct")) else 0
        vol_r = float(row["Vol_Ratio"]) if pd.notna(row.get("Vol_Ratio")) else 1
        rsi = float(row["RSI_14"]) if pd.notna(row.get("RSI_14")) else 50
        cross = str(row.get("Cross_Signal", ""))
        close = float(row["Close"]) if pd.notna(row.get("Close")) else 0
        highlights.append({
            "t": ticker, "chg": chg, "vol": vol_r, "rsi": rsi,
            "cross": cross, "close": close
        })

    # 가장 주목할 종목 선정 (변동률 + 거래량 급증)
    highlights.sort(key=lambda x: abs(x["chg"]) + (x["vol"] - 1) * 2, reverse=True)
    top_stocks = highlights[:2]

    focus_parts = []
    for s in top_stocks:
        parts = []
        name = NAMES.get(s["t"], s["t"])
        parts.append(f'{s["t"]}({name})')
        if abs(s["chg"]) > 0.5:
            direction = "상승" if s["chg"] > 0 else "하락"
            parts.append(f'전일 {s["chg"]:+.1f}% {direction}')
        if s["vol"] > 2:
            parts.append(f'거래량 {s["vol"]:.1f}배 급증')
        if s["cross"] == "GOLDEN_CROSS":
            parts.append("골든크로스 발생")
        elif s["cross"] == "DEAD_CROSS":
            parts.append("데드크로스 주의")
        if s["rsi"] > 70:
            parts.append(f'RSI {s["rsi"]:.0f} 과매수 구간')
        elif s["rsi"] < 30:
            parts.append(f'RSI {s["rsi"]:.0f} 과매도 구간')
        # FC에서 애널리스트 목표가 찾기
        for f_item in fc:
            if f_item["t"] == s["t"] and f_item.get("at_avg", 0) > 0:
                upside = ((f_item["at_avg"] - s["close"]) / s["close"] * 100) if s["close"] > 0 else 0
                parts.append(f'애널리스트 목표 ${f_item["at_avg"]:.0f}({upside:+.0f}%)')
                break
        focus_parts.append(". ".join(parts))

    focus_text = ". ".join(focus_parts) if focus_parts else "포트폴리오 종목 전반 안정적 흐름"

    # ── ② 위험 요인 ──
    risk_items = []
    for s in highlights:
        if s["chg"] < -3:
            risk_items.append(f'{s["t"]} {s["chg"]:+.1f}% 급락')
        if s["rsi"] > 75:
            risk_items.append(f'{s["t"]} RSI {s["rsi"]:.0f} 과매수')
        if s["cross"] == "DEAD_CROSS":
            risk_items.append(f'{s["t"]} 데드크로스')

    # 섹터 집중도 리스크 (IT 비중 73%)
    risk_items.append("포트폴리오 IT섹터 집중도 73% — 섹터 조정 시 동반 하락 위험")

    # 레버리지 리스크
    risk_items.append("SOXL(3x 레버리지) 단기 변동성 확대 가능")

    risk_text = risk_items[0] if risk_items else "현재 포트폴리오에 특별한 리스크 시그널 없음"

    # ── ③ 오늘의 액션 ──
    action_items = []
    for s in highlights:
        if s["cross"] == "GOLDEN_CROSS":
            action_items.append(f'{s["t"]} 골든크로스 — 추가 매수 검토')
        elif s["rsi"] < 30:
            action_items.append(f'{s["t"]} RSI 과매도 — 분할 매수 기회')
        elif s["rsi"] > 80:
            action_items.append(f'{s["t"]} 과매수 — 부분 차익실현 고려')
        elif s["chg"] > 3 and s["vol"] > 2:
            action_items.append(f'{s["t"]} 급등+거래량 — 모멘텀 매수 검토')

    if not action_items:
        action_items.append("전 종목 홀딩 유지. 급격한 시그널 없음")

    action_text = ". ".join(action_items[:2])

    # ── 이벤트 요약 ──
    ev_summary_items = []
    for e in ev[:5]:
        ev_summary_items.append(f'{e["t"]}: {e["e"]}')
    ev_summary = " | ".join(ev_summary_items) if ev_summary_items else ""

    # ── 포트폴리오 전체 수익률 요약 ──
    port_chg_sum = sum(s["chg"] for s in highlights)
    port_chg_avg = port_chg_sum / len(highlights) if highlights else 0
    port_up = sum(1 for s in highlights if s["chg"] > 0)
    port_dn = sum(1 for s in highlights if s["chg"] < 0)

    # ── HTML 빌드 (세련된 카드 디자인) ──
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    h = []

    # 상단 헤더: 날짜 + 포트 요약
    trend_color = "var(--up)" if port_chg_avg >= 0 else "var(--dn)"
    trend_icon = "▲" if port_chg_avg >= 0 else "▼"
    h.append(
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'padding:5px 8px;margin-bottom:6px;background:linear-gradient(135deg,rgba(91,141,238,.08),rgba(157,110,255,.06));'
        f'border:1px solid rgba(91,141,238,.12);border-radius:6px">'
        f'<span style="font-size:8px;color:var(--t2);font-family:var(--font-mono)">{data_date}</span>'
        f'<span style="font-size:8px;font-weight:700;color:{trend_color}">'
        f'{trend_icon} {port_chg_avg:+.1f}% avg · {port_up}↑ {port_dn}↓</span></div>'
    )

    # ① 주목 종목 — 개별 종목 미니카드
    h.append(
        '<div style="margin-bottom:5px">'
        '<div style="display:flex;align-items:center;gap:4px;margin-bottom:4px;padding:0 2px">'
        '<span style="font-size:9px;font-weight:800;color:var(--blue)">①</span>'
        '<span style="font-size:7.5px;font-weight:700;color:var(--t2);letter-spacing:.5px;text-transform:uppercase">Focus</span>'
        '</div>'
    )
    for s in top_stocks:
        chg_c = "var(--up)" if s["chg"] >= 0 else "var(--dn)"
        chg_icon = "▲" if s["chg"] >= 0 else "▼"
        # 애널리스트 목표가 찾기
        at_info = ""
        for f_item in fc:
            if f_item["t"] == s["t"] and f_item.get("at_avg", 0) > 0:
                upside = ((f_item["at_avg"] - s["close"]) / s["close"] * 100) if s["close"] > 0 else 0
                up_c = "var(--up)" if upside >= 0 else "var(--dn)"
                at_info = (
                    f'<span style="font-size:7px;color:var(--t3);margin-left:auto;white-space:nowrap">'
                    f'🎯 ${f_item["at_avg"]:.0f}'
                    f'<span style="color:{up_c};font-weight:700"> {upside:+.0f}%</span></span>'
                )
                break
        # RSI 바
        rsi_pct = min(s["rsi"], 100)
        rsi_c = "var(--dn)" if s["rsi"] > 70 else ("var(--up)" if s["rsi"] < 30 else "var(--oj)")
        # 시그널 태그
        tags = []
        if s["cross"] == "GOLDEN_CROSS":
            tags.append('<span style="font-size:6px;padding:1px 3px;background:rgba(48,217,136,.15);color:var(--up);border-radius:3px">GC</span>')
        elif s["cross"] == "DEAD_CROSS":
            tags.append('<span style="font-size:6px;padding:1px 3px;background:rgba(240,82,82,.15);color:var(--dn);border-radius:3px">DC</span>')
        if s["vol"] > 2:
            tags.append(f'<span style="font-size:6px;padding:1px 3px;background:rgba(245,166,35,.15);color:var(--oj);border-radius:3px">VOL {s["vol"]:.0f}x</span>')
        tag_html = " ".join(tags)

        h.append(
            f'<div style="padding:5px 7px;margin-bottom:3px;background:var(--s2);border:1px solid var(--b2);border-radius:5px">'
            f'<div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap">'
            f'<span style="font-size:9px;font-weight:800;color:{TC.get(s["t"], "var(--blue)")}">{s["t"]}</span>'
            f'<span style="font-size:8px;font-weight:700;color:{chg_c}">{chg_icon}{abs(s["chg"]):.1f}%</span>'
            f'{tag_html}{at_info}</div>'
            f'<div style="margin-top:3px;height:2px;background:var(--b2);border-radius:1px;overflow:hidden">'
            f'<div style="width:{rsi_pct}%;height:100%;background:{rsi_c};border-radius:1px"></div></div>'
            f'<div style="display:flex;justify-content:space-between;margin-top:2px">'
            f'<span style="font-size:6px;color:var(--t4)">RSI {s["rsi"]:.0f}</span>'
            f'<span style="font-size:6px;color:var(--t4)">${s["close"]:.1f}</span></div></div>'
        )
    h.append('</div>')

    # ② 위험 요인
    h.append(
        f'<div style="padding:5px 7px;margin-bottom:5px;background:rgba(240,82,82,.04);'
        f'border:1px solid rgba(240,82,82,.1);border-left:2px solid var(--dn);border-radius:5px">'
        f'<div style="display:flex;align-items:center;gap:4px;margin-bottom:2px">'
        f'<span style="font-size:9px;font-weight:800;color:var(--dn)">②</span>'
        f'<span style="font-size:7.5px;font-weight:700;color:var(--t2);letter-spacing:.5px;text-transform:uppercase">Risk</span></div>'
        f'<div style="font-size:8px;line-height:1.55;color:var(--t2)">{risk_text}</div></div>'
    )

    # ③ 오늘의 액션
    h.append(
        f'<div style="padding:5px 7px;margin-bottom:5px;background:rgba(48,217,136,.04);'
        f'border:1px solid rgba(48,217,136,.1);border-left:2px solid var(--up);border-radius:5px">'
        f'<div style="display:flex;align-items:center;gap:4px;margin-bottom:2px">'
        f'<span style="font-size:9px;font-weight:800;color:var(--up)">③</span>'
        f'<span style="font-size:7.5px;font-weight:700;color:var(--t2);letter-spacing:.5px;text-transform:uppercase">Action</span></div>'
        f'<div style="font-size:8px;line-height:1.55;color:var(--t2)">{action_text}</div></div>'
    )

    # 하단: 시그널 + 타임스탬프
    ev_tags = []
    for e in ev[:4]:
        ev_tags.append(
            f'<span style="font-size:6.5px;padding:1px 4px;background:var(--s2);border:1px solid var(--b2);'
            f'border-radius:3px;color:var(--t3);white-space:nowrap">{e["t"]} {e["e"][:12]}</span>'
        )
    ev_html = " ".join(ev_tags) if ev_tags else ""
    if ev_html:
        h.append(
            f'<div style="display:flex;flex-wrap:wrap;gap:2px;padding:3px 2px">{ev_html}</div>'
        )

    h.append(
        f'<div style="display:flex;justify-content:space-between;align-items:center;padding:2px 2px 0;margin-top:2px">'
        f'<span style="font-size:6px;color:var(--t4);font-family:var(--font-mono)">DAILY BRIEFING</span>'
        f'<span style="font-size:6px;color:var(--t4);font-family:var(--font-mono)">{now_str}</span></div>'
    )

    return "\n".join(h)


def build_momentum_notes(df, fc, data_date):
    """모멘텀 노트 HTML 생성 — 포트폴리오 종목별 현황 + 크립토.

    generate_dashboard.py가 매일 7시에 실행될 때 최신 데이터로 갱신.
    """
    date_short = data_date[5:].replace("-", "/")
    h = []

    # 종목별 색상 매핑
    accent = {
        "NVDA": ("var(--up)", "48,217,136"),
        "TSM":  ("var(--pu)", "157,110,255"),
        "MU":   ("var(--blue)", "91,141,238"),
        "CEG":  ("var(--oj)", "245,166,35"),
        "SOXL": ("#FF6B35", "255,107,53"),
        "IONQ": ("var(--dn)", "240,82,82"),
        "CRCL": ("var(--cy)", "34,211,238"),
    }

    for ticker in MYPORT:
        rows = df[df["Ticker"] == ticker]
        detail = PORT_DETAIL.get(ticker, {})
        avg_price = detail.get("avg", 0)
        qty = detail.get("qty", 0)
        color, rgb = accent.get(ticker, ("var(--blue)", "91,141,238"))

        if len(rows) > 0:
            row = rows.iloc[0]
            close = float(row["Close"]) if pd.notna(row.get("Close")) else 0
            chg = float(row["Daily_Chg_Pct"]) if pd.notna(row.get("Daily_Chg_Pct")) else 0
            rsi = float(row["RSI_14"]) if pd.notna(row.get("RSI_14")) else 50
            cross = str(row.get("Cross_Signal", ""))
            vol_r = float(row["Vol_Ratio"]) if pd.notna(row.get("Vol_Ratio")) else 1
        else:
            close = 0
            chg = 0
            rsi = 50
            cross = ""
            vol_r = 1

        # 수익률 계산
        perf = ((close - avg_price) / avg_price * 100) if avg_price > 0 and close > 0 else 0

        # 태그 결정
        if cross == "GOLDEN_CROSS":
            tag_text, tag_bg = "▲ 골든크로스", f"rgba({rgb},0.12)"
        elif cross == "DEAD_CROSS":
            tag_text, tag_bg = "▼ 데드크로스", f"rgba({rgb},0.12)"
        elif rsi > 70:
            tag_text, tag_bg = f"RSI {rsi:.0f}", f"rgba({rgb},0.12)"
        elif chg > 2:
            tag_text, tag_bg = f"▲ +{chg:.1f}%", f"rgba({rgb},0.12)"
        elif chg < -2:
            tag_text, tag_bg = f"▼ {chg:.1f}%", f"rgba({rgb},0.12)"
        elif perf > 20:
            tag_text, tag_bg = "▲ 고수익", f"rgba({rgb},0.12)"
        elif perf < -10:
            tag_text, tag_bg = "▼ 손실중", f"rgba({rgb},0.12)"
        else:
            tag_text, tag_bg = "◆ 홀딩", f"rgba({rgb},0.12)"

        # 애널리스트 목표가
        target_str = ""
        for f_item in fc:
            if f_item["t"] == ticker and f_item.get("at_avg", 0) > 0:
                target_str = f'목표 ${f_item["at_avg"]:.0f}'
                break

        # 이슈 텍스트 (RSI + 거래량 + 가격변동 기반)
        issues = []
        if abs(chg) > 1:
            direction = "상승" if chg > 0 else "하락"
            issues.append(f'전일 <span style="color:#fff;font-weight:600">{chg:+.1f}% {direction}</span>')
        if vol_r > 2:
            issues.append(f'거래량 {vol_r:.1f}배 급증')
        if rsi > 70:
            issues.append(f'RSI {rsi:.0f} 과매수 주의')
        elif rsi < 30:
            issues.append(f'RSI {rsi:.0f} 과매도 영역')
        if cross == "GOLDEN_CROSS":
            issues.append("MA5↑MA20 골든크로스 발생")
        elif cross == "DEAD_CROSS":
            issues.append("MA5↓MA20 데드크로스 경고")
        if not issues:
            issues.append(f'${close:.1f} 횡보 중. RSI {rsi:.0f} 중립 구간')

        issue_text = ". ".join(issues)
        sub_text = f'{target_str} · 현재 {perf:+.1f}% / avg ${avg_price:.2f}' if target_str else f'현재 {perf:+.1f}% / avg ${avg_price:.2f}'

        h.append(
            f'<div class="mom-row" style="border-left:3px solid {color}">'
            f'<div class="mom-left"><div class="mom-dot" style="background:{color}"></div></div>'
            f'<div class="mom-body">'
            f'<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">'
            f'<b style="color:{color}">{ticker}</b>'
            f'<span class="mom-tag" style="background:{tag_bg};color:{color}">{tag_text}</span>'
            f'<span style="font-size:9px;color:var(--t3);font-family:var(--font-mono);margin-left:auto">{date_short}</span>'
            f'</div>'
            f'<div class="mom-issue">{issue_text}</div>'
            f'<div class="mom-sub">{sub_text}</div>'
            f'</div></div>'
        )

    # BTC — 크립토 (고정 메모)
    h.append(
        '<div class="mom-row" style="border-left:3px solid #f7931a">'
        '<div class="mom-left"><div class="mom-dot" style="background:#f7931a"></div></div>'
        '<div class="mom-body">'
        '<div style="display:flex;align-items:center;gap:6px;margin-bottom:5px">'
        '<b style="color:#f7931a">BTC</b>'
        '<span class="mom-tag" style="background:rgba(247,147,26,0.12);color:#f7931a">₿ 크립토</span>'
        f'<span style="font-size:9px;color:var(--t3);font-family:var(--font-mono);margin-left:auto">{date_short}</span>'
        '</div>'
        '<div class="mom-issue">0.1503 BTC 보유. 장기 홀딩 전략 유지. 가격은 실시간 API로 갱신.</div>'
        '<div class="mom-sub">장기투자 · 손익은 실시간 반영</div>'
        '</div></div>'
    )

    return "\n".join(h)


def build_sa_html(data_date, ticker_count):
    """SA (Supply Analysis) 패널 HTML 조각 생성 — 센터 탭2 콘텐츠."""
    return f"""<div class="sa-w">
  <div class="sa-hd">
    <div><h1>SUPPLY ANALYSIS</h1><div class="sub" id="sa-hSub">7 Holdings &middot; S&amp;P 500 &middot; {data_date}</div></div>
    <div class="rv"><div class="lb">PORTFOLIO 1M AVG</div><div class="vl" id="sa-hAvg">&mdash;</div></div>
  </div>
  <div class="sa-st" id="sa-strip"></div>
  <div class="sa-tb" id="sa-tabs"></div>
  <div id="sa-pSignal"></div>
  <div id="sa-pTarget" class="hide"></div>
  <div id="sa-pPort" class="hide"></div>
  <div id="sa-pRec" class="hide"></div>
  <div id="sa-pSector" class="hide"></div>
  <div class="ft">&olarr; 수급 분석 &middot; S&amp;P 500 {ticker_count}종목 &middot; {data_date}</div>
</div>"""


def build_sa_right_html():
    """SA 우측 패널 하이라이트 HTML."""
    return """<div class="panel-hd si" style="background:var(--s1)">
  <div class="panel-hd-title">
    <div class="panel-hd-ico" style="background:rgba(91,141,238,.12)">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="var(--blue)" stroke-width="2.5"><path d="M22 12h-4l-3 9L9 3l-3 9H2"/></svg>
    </div>
    <div class="panel-hd-text">수급 하이라이트</div>
  </div>
  <div class="panel-hd-badge" style="background:rgba(91,141,238,.15);color:var(--blue)">SUPPLY</div>
</div>
<div class="rp-section" id="sa-highlights" style="padding:6px 0">
  <div style="padding:10px 14px;font-size:10px;color:var(--t3)">수급 분석 탭을 클릭하세요</div>
</div>"""


def build_sa_script(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas=None):
    """SA 모듈 전체 스크립트 생성 — IIFE로 SA 네임스페이스 격리."""
    j = lambda obj: json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

    # ── 원본 JS 추출 (UTILS ~ INIT 이전) ──
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        html = f.read()
    script_start = html.find("<script>", html.find("chart.js")) + len("<script>")
    script_end = html.find("</script>", script_start)
    orig_js = html[script_start:script_end]

    sep = "// ═══════════════════════════════════════════════"
    utils_idx = orig_js.find(sep + "\n// UTILS")
    init_idx = orig_js.find(sep + "\n// INIT")
    if utils_idx < 0 or init_idx < 0:
        log.error("원본 HTML에서 JS 섹션 마커를 찾을 수 없습니다.")
        sys.exit(1)
    kept_js = orig_js[utils_idx:init_idx]

    # ALPHA VANTAGE FETCH / SMART SCAN ENGINE 제거
    fetch_start = kept_js.find(sep + "\n// ALPHA VANTAGE FETCH")
    scoring_start = kept_js.find(sep + "\n// SCORING ENGINE")
    if fetch_start >= 0 and scoring_start >= 0:
        kept_js = kept_js[:fetch_start] + "\n// (Offline mode - fetch/scan removed)\n" + kept_js[scoring_start:]

    # calcAllSP5Sigs 하드코딩 수정
    kept_js = kept_js.replace(
        'var sp500tickers=["NVDA","TSM","MU","CEG"];',
        'var sp500tickers=[];for(var pi=0;pi<MYPORT.length;pi++){if(SP5_INFO[MYPORT[pi]])sp500tickers.push(MYPORT[pi])}'
    )

    # noDataMsg 수정
    old_nodata = 'function noDataMsg(msg){return\'<div class="cd" style="text-align:center;padding:30px"><div style="font-size:28px;margin-bottom:10px">📡</div><div style="font-size:12px;font-weight:700;margin-bottom:6px">\'+msg+\'</div><div style="font-size:9px;color:var(--dim);line-height:1.7">① API 키 입력 → "🔄 전체 스캔" 클릭<br>② 포트폴리오(7) → 섹터ETF(11) → S&P500 순차 스캔<br><a href="https://www.alphavantage.co/support/#api-key" target="_blank" style="color:var(--b)">▸ 무료 API 키 발급</a></div></div>'
    new_nodata = 'function noDataMsg(msg){return\'<div class="cd" style="text-align:center;padding:30px"><div style="font-size:28px;margin-bottom:10px">📡</div><div style="font-size:12px;font-weight:700;margin-bottom:6px">\'+msg+\'</div><div style="font-size:9px;color:var(--dim)">오프라인 데이터가 로드되지 않았습니다.</div></div>'
    kept_js = kept_js.replace(old_nodata, new_nodata)

    # ── $ 함수를 sa- 접두사 우선 탐색으로 교체 ──
    kept_js = kept_js.replace(
        'function $(id){return document.getElementById(id)}',
        'function $(id){return document.getElementById("sa-"+id)||document.getElementById(id)}'
    )

    # ── API 관련 유틸리티 함수 제거 (no-op 대체) ──

    # showErr: $("errArea") 참조 — 단일행 함수
    kept_js = re.sub(
        r'function showErr\(msg\)\{[^\n]+\}',
        'function showErr(msg){}',
        kept_js
    )

    # getQuota: 다중행 함수 — \n}로 종료
    kept_js = re.sub(
        r'function getQuota\(\)\{[\s\S]*?\n\}',
        'function getQuota(){return 0}',
        kept_js
    )

    # addQuota: 단일행 함수
    kept_js = re.sub(
        r'function addQuota\(n\)\{[^\n]+\}',
        'function addQuota(n){}',
        kept_js
    )

    # updateQuotaUI: $("quotaArea") 참조 — 단일행 함수
    kept_js = re.sub(
        r'function updateQuotaUI\(\)\{[^\n]+\}',
        'function updateQuotaUI(){}',
        kept_js
    )

    # isCachedToday: 다중행 함수 — 오프라인 데이터는 항상 true
    kept_js = re.sub(
        r'function isCachedToday\(tk\)\{[\s\S]*?\n\}',
        'function isCachedToday(tk){return true}',
        kept_js
    )

    # cacheAge: 다중행 함수 — 오프라인 데이터는 고정 라벨
    kept_js = re.sub(
        r'function cacheAge\(tk\)\{[\s\S]*?\n\}',
        'function cacheAge(tk){return "오프라인"}',
        kept_js
    )

    # saveTD / loadTD: localStorage 캐시 불필요 — 다중행
    kept_js = re.sub(
        r'// Save/Load TD to localStorage\n',
        '',
        kept_js
    )
    kept_js = re.sub(
        r'function saveTD\(\)\{[\s\S]*?\n\}',
        'function saveTD(){}',
        kept_js
    )
    kept_js = re.sub(
        r'function loadTD\(\)\{[\s\S]*?\n\}',
        'function loadTD(){}',
        kept_js
    )

    # Free tier API 힌트 메시지 제거 (문자열 연결이 포함된 복합 라인)
    kept_js = re.sub(
        r"if\(scanned<total\)h\+='[^\n]+",
        "if(scanned<total)h+='<div style=\"font-size:8px;color:var(--a);margin-bottom:6px\">📊 데이터가 없는 종목은 표시되지 않습니다.</div>';",
        kept_js
    )

    # ═══════════════════════════════════════════════════════════════
    # rRec() 전면 재작성: 시그널별 Top10 그리드 UI
    # ═══════════════════════════════════════════════════════════════
    new_rRec = r'''function rRec(){
  var all=calcAllSP5Sigs();
  var scanned=all.length,total=SP5.length+4;
  if(scanned===0){$("pRec").innerHTML=noDataMsg("S\uFF06P 500 스캔이 필요합니다");return}

  /* 시그널별 분류 */
  var grp={"강력매수":[],"매수":[],"관망":[],"매도":[],"강력매도":[]};
  var port=[];
  for(var c=0;c<all.length;c++){
    var sig=all[c].sig;
    if(grp[sig])grp[sig].push(all[c]);
    if(all[c].inPort)port.push(all[c]);
  }

  /* 카드 빌더 헬퍼 */
  function mkCard(x,rank){
    var portBadge=x.inPort?'<span style="background:var(--a);color:#000;font-size:6px;padding:1px 3px;border-radius:2px;margin-left:3px;font-weight:700">MY</span>':"";
    var betaH=typeof BETA_MAP!=="undefined"&&BETA_MAP[x.t]?'<span style="font-size:7px;color:var(--dim)">β '+BETA_MAP[x.t].toFixed(2)+'</span>':"";
    return '<div class="rec-card"><div class="rec-card-head"><span class="rec-rank" style="background:'+x.sigC+'33;color:'+x.sigC+'">'+(rank+1)+'</span><div style="flex:1;min-width:0"><div style="font-weight:700;font-size:11px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+x.t+portBadge+'</div><div style="font-size:7px;color:var(--dim);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+x.n+'</div></div><div style="text-align:right"><div style="font-weight:700;font-size:11px">$'+pf(x.p)+'</div><div class="'+pc(x.r)+'" style="font-size:8px;font-weight:600">'+pp(x.r)+'</div></div></div><div class="rec-card-body"><div class="rec-stat"><span class="rec-stat-lb">종합</span><span style="font-weight:800;color:'+x.sigC+';font-size:11px">'+(x.comp>0?"+":"")+x.comp+'</span></div><div class="rec-stat"><span class="rec-stat-lb">기술</span><span class="'+pc(x.tech/100)+'">'+(x.tech>0?"+":"")+x.tech+'</span></div><div class="rec-stat"><span class="rec-stat-lb">수급</span><span class="'+pc(x.sup/100)+'">'+(x.sup>0?"+":"")+x.sup+'</span></div><div class="rec-stat"><span class="rec-stat-lb">MA</span><span style="color:'+x.maC+';font-size:8px">'+x.maSt+'</span></div></div><div style="display:flex;justify-content:space-between;align-items:center;margin-top:4px"><span class="pl" style="background:'+x.sigC+'22;color:'+x.sigC+';border:1px solid '+x.sigC+'44;font-size:8px">'+x.sig+'</span>'+betaH+'</div></div>';
  }

  /* 그룹 섹션 빌더 */
  function mkSection(icon,title,color,items,limit,asc){
    if(!items||items.length===0)return '';
    var sorted=items.slice().sort(function(a,b){return asc?a.comp-b.comp:b.comp-a.comp});
    var show=sorted.slice(0,limit||10);
    var h='<div class="rec-section"><div class="rec-sec-title" style="border-left:3px solid '+color+';padding-left:8px"><span>'+icon+' '+title+'</span><span class="rec-sec-cnt">'+items.length+'종목</span></div><div class="rec-grid">';
    for(var i=0;i<show.length;i++){h+=mkCard(show[i],i)}
    h+='</div></div>';
    return h;
  }

  /* ── HTML 조합 ── */
  var h='';

  /* 1) 요약 통계 카드 */
  var strong=grp["강력매수"].length,buy=grp["매수"].length,hold=grp["관망"].length;
  var sell=grp["매도"].length,strongSell=grp["강력매도"].length;
  h+='<div class="cd"><div class="ct">📊 S&P 500 종합 시그널 분석</div>';
  h+='<div class="rec-summary">';
  h+='<div class="rec-sum-item" style="border-bottom:2px solid #00ff88"><div class="rec-sum-n">'+strong+'</div><div class="rec-sum-lb">강력매수</div></div>';
  h+='<div class="rec-sum-item" style="border-bottom:2px solid var(--g)"><div class="rec-sum-n">'+buy+'</div><div class="rec-sum-lb">매수</div></div>';
  h+='<div class="rec-sum-item" style="border-bottom:2px solid var(--a)"><div class="rec-sum-n">'+hold+'</div><div class="rec-sum-lb">관망</div></div>';
  h+='<div class="rec-sum-item" style="border-bottom:2px solid var(--r)"><div class="rec-sum-n">'+(sell+strongSell)+'</div><div class="rec-sum-lb">매도</div></div>';
  h+='<div class="rec-sum-item" style="border-bottom:2px solid var(--b)"><div class="rec-sum-n">'+scanned+'</div><div class="rec-sum-lb">분석완료</div></div>';
  h+='</div></div>';

  /* 2) 내 보유 종목 (전체 표시) */
  if(port.length>0){
    h+=mkSection('\uD83D\uDCBC','내 보유 종목','var(--a)',port,port.length,false);
  }

  /* 3) 강력매수 Top 10 */
  h+=mkSection('\uD83D\uDD25','강력매수 Top 10','#00ff88',grp["\uAC15\uB825\uB9E4\uC218"],10,false);

  /* 4) 매수 Top 10 */
  h+=mkSection('\uD83D\uDFE2','매수 Top 10','var(--g)',grp["\uB9E4\uC218"],10,false);

  /* 5) 관망 Top 10 */
  h+=mkSection('\uD83D\uDFE1','관망 Top 10','var(--a)',grp["\uAD00\uB9DD"],10,false);

  /* 6) 매도 하위 10 */
  h+=mkSection('\uD83D\uDD34','매도 하위 10','var(--r)',grp["\uB9E4\uB3C4"],10,true);

  /* 7) 강력매도 하위 10 */
  h+=mkSection('\u2744\uFE0F','강력매도 하위 10','#ff1744',grp["\uAC15\uB825\uB9E4\uB3C4"],10,true);

  $("pRec").innerHTML=h;
}'''

    kept_js = re.sub(
        r'function rRec\(\)\{[\s\S]*?\$\("pRec"\)\.innerHTML=h;\s*\}',
        lambda m: new_rRec,
        kept_js,
        count=1
    )
    log.info("rRec() replaced with signal-based Top10 grid UI")

    # ── 전체 SP5 티커 리스트 ──
    smap_tickers = set()
    for s in smap:
        smap_tickers.update(s["st"])
    remaining = [t for t in all_sp5_tickers if t not in smap_tickers and t not in MYPORT and t not in SECTOR_ETFS]

    # ── 목표가 초기값 ──
    ut_init = {}
    for item in fc:
        if item["lo"] > 0 and item["hi"] > 0:
            ut_init[item["t"]] = round((item["lo"] + item["hi"]) / 2)
        else:
            ut_init[item["t"]] = 0

    # ── SA 데이터 스텁 블록 (실제 데이터는 data.js → <script> 태그로 로드) ──
    sa_data = f"""
// ═══════════════════════════════════════════════
// SA DATA STUB (Auto-generated: {data_date})
// 실제 데이터는 data.js → SA.loadData(SA_DATA)로 로드
// ═══════════════════════════════════════════════
var SA = (function() {{
"use strict";

var MYPORT = [];
var TC = {{}};
var NAMES = {{}};
var PORT_SEC = {{}};
var ALLOC_W = {{}};
var SMAP = [];
var SETFS = [];
var SP5 = [];
var SP5_INFO = {{}};
var FN = {{}};
var FC = [];
var EV = [];
var BETA_MAP = {{}};

var sel = "NVDA", tab = "signal", ca = 5;
var ut = {{}};
var charts = {{}};
var TD = {{}};
var _dataLoaded = false;

function loadData(d) {{
  MYPORT = d.MYPORT || [];
  TC = d.TC || {{}};
  NAMES = d.NAMES || {{}};
  PORT_SEC = d.PORT_SEC || {{}};
  ALLOC_W = d.ALLOC_W || {{}};
  SMAP = d.SMAP || [];
  SP5_INFO = d.SP5_INFO || {{}};
  FC = d.FC || [];
  EV = d.EV || [];
  BETA_MAP = d.BETA_MAP || {{}};
  ut = d.ut_init || {{}};

  /* TD는 객체 identity 보존 (SA.TD 참조 유지) — 머지 방식 */
  var k;
  for (k in TD) {{ if (TD.hasOwnProperty(k)) delete TD[k]; }}
  var tdNew = d.TD || {{}};
  for (k in tdNew) {{ if (tdNew.hasOwnProperty(k)) TD[k] = tdNew[k]; }}

  /* SETFS, SP5 배열 재구축 */
  SETFS = []; SP5 = [];
  for (var si = 0; si < SMAP.length; si++) {{
    SETFS.push(SMAP[si].etf);
    for (var sj = 0; sj < SMAP[si].st.length; sj++) {{
      var tk = SMAP[si].st[sj];
      if (MYPORT.indexOf(tk) < 0) SP5.push(tk);
    }}
  }}
  var extra = d.EXTRA_SP5 || [];
  for (var ei = 0; ei < extra.length; ei++) {{
    if (SP5.indexOf(extra[ei]) < 0) SP5.push(extra[ei]);
  }}

  _dataLoaded = true;
  console.log('[SA] data loaded — TD:', Object.keys(TD).length, 'tickers');
}}
"""

    # ── SA 로직 블록 (placeholder: __SA_LOGIC_PLACEHOLDER__) ──
    sa_logic = f"""
{kept_js}

// ═══════════════════════════════════════════════
// SA INIT + PUBLIC API
// ═══════════════════════════════════════════════
function saInit() {{
  try {{
    /* mainData 목표가 → SA ut 연동 */
    if (window.mainData) {{
      for (var mi = 0; mi < window.mainData.length; mi++) {{
        var md = window.mainData[mi];
        if (md.targetPrice && md.targetPrice > 0) {{
          ut[md.ticker] = md.targetPrice;
        }}
      }}
    }}
    render(); renderHighlights();

    /* SA 목표가 input 변경 → 포커스 유지하며 실시간 UI 업데이트 */
    var saPanel = document.getElementById("sa-pTarget");
    if (saPanel) {{
      saPanel.addEventListener("input", function(e) {{
        if (e.target && e.target.classList.contains("tgt-input")) {{
          var tk = e.target.getAttribute("data-tk");
          var val = parseFloat(e.target.value) || 0;

          /* mainData 역동기화 */
          if (window.mainData) {{
            var item = window.mainData.find(function(a) {{ return a.ticker === tk; }});
            if (item) item.targetPrice = val;
          }}

          /* 실시간 UI 업데이트: 리턴률 + 막대 채우기 */
          var row = e.target.closest('.tr');
          if (row && val > 0) {{
            var cpSpan = row.querySelector('span[style*="font-weight:600"]');
            var cpMatch = cpSpan ? cpSpan.textContent.match(/\$([0-9.]+)/) : null;
            var cp = cpMatch ? parseFloat(cpMatch[1]) : 0;
            var retEl = row.querySelector('.tgt-ret');
            var fillEl = row.querySelector('.tgt-fill');

            if (cp > 0) {{
              var up = (val - cp) / cp;
              var pct = (up * 100).toFixed(1);
              if (retEl) {{
                retEl.textContent = (up >= 0 ? '+' : '') + pct + '%';
                retEl.style.color = up >= 0 ? 'var(--sa-g)' : 'var(--sa-r)';
              }}
              if (fillEl) {{
                fillEl.style.width = Math.min(Math.abs(up) * 200, 100) + '%';
                fillEl.style.background = up >= 0 ? 'var(--sa-g)' : 'var(--sa-r)';
                fillEl.style.left = up >= 0 ? '50%' : 'auto';
                fillEl.style.right = up >= 0 ? 'auto' : '50%';
              }}
            }}
          }}
        }}
      }});
    }}
  }} catch(e) {{
    console.error("SA init error:", e);
  }}
}}

function renderHighlights() {{
  var el = document.getElementById("sa-highlights");
  if (!el) return;
  var sigs = calcPortSigs();
  var h = "";
  // Top 3 strongest signals
  var sorted = sigs.filter(function(s){{ return !s.noData; }}).sort(function(a,b){{ return Math.abs(b.comp) - Math.abs(a.comp); }});
  for (var i = 0; i < Math.min(sorted.length, 4); i++) {{
    var s = sorted[i];
    var sjC = s.supJ.indexOf("매수") >= 0 ? "var(--up)" : s.supJ.indexOf("매도") >= 0 ? "var(--dn)" : "var(--oj)";
    h += '<div class="sa-highlight-card">';
    h += '<div style="display:flex;justify-content:space-between;align-items:center">';
    h += '<span style="font-weight:800;color:' + (TC[s.t]||"var(--blue)") + '">' + s.t + '</span>';
    h += '<span style="font-size:11px;font-weight:900;color:' + s.sigC + '">' + s.sig + '</span>';
    h += '</div>';
    h += '<div style="display:flex;gap:8px;margin-top:5px;font-size:9px">';
    h += '<span style="color:var(--t2)">기술 <b style="color:' + (s.tech>=0?"var(--up)":"var(--dn)") + '">' + (s.tech>0?"+":"") + s.tech + '</b></span>';
    h += '<span style="color:var(--t2)">수급 <b style="color:' + sjC + '">' + (s.sup>0?"+":"") + s.sup + '</b></span>';
    h += '<span style="color:var(--t2)">종합 <b style="color:' + s.sigC + '">' + (s.comp>0?"+":"") + s.comp + '</b></span>';
    h += '</div></div>';
  }}
  if (h === "") h = '<div style="padding:10px 14px;font-size:10px;color:var(--t3)">데이터 로딩 중...</div>';
  el.innerHTML = h;
}}

/* ── AI 리밸런싱 시뮬레이터 ── */
function rRebalanceSim() {{
  var container = $("pPort");
  if (!container) return;
  /* 기존 리밸런싱 블록 제거 (중복 방지) */
  var old = container.querySelector('.reb-sim-block');
  if (old) old.remove();

  var stats = (typeof calcStats === 'function') ? calcStats() : null;
  if (!stats) return;
  var totalVal = stats.total;
  if (totalVal <= 0) return;

  var h = '<div class="reb-sim-block cd" style="margin-top:10px; border-top: 2px solid var(--sa-b)">';
  h += '<div class="ct" style="display:flex;align-items:center;gap:6px">⚖️ AI 리밸런싱 시뮬레이터</div>';
  h += '<table class="tbl"><thead><tr><th>종목</th><th>현재비중</th><th>목표비중</th><th>추천 액션</th></tr></thead><tbody>';

  if (typeof MYPORT !== 'undefined' && typeof ALLOC_W !== 'undefined' && window.mainData) {{
    MYPORT.forEach(function(tk) {{
      var item = window.mainData.find(function(a) {{ return a.ticker === tk; }});
      if (!item) return;
      var currPct = (item.qty * item.currentPrice * FX / totalVal * 100);
      var targetPct = ALLOC_W[tk] || 0;
      var diffWon = (totalVal * ((targetPct - currPct) / 100));
      var actionQty = Math.floor(diffWon / (item.currentPrice * FX));

      var actionColor = actionQty > 0 ? 'var(--sa-g)' : (actionQty < 0 ? 'var(--sa-r)' : 'var(--dim)');
      var actionText = actionQty !== 0 ? Math.abs(actionQty) + '주 ' + (actionQty > 0 ? '매수' : '매도') : '적정';
      h += '<tr><td style="font-weight:700">' + tk + '</td>';
      h += '<td>' + currPct.toFixed(1) + '%</td>';
      h += '<td style="color:var(--sa-b)">' + targetPct.toFixed(1) + '%</td>';
      h += '<td><span class="pl" style="background:' + actionColor + '22; color:' + actionColor + '">' + actionText + '</span></td></tr>';
    }});
  }}
  h += '</tbody></table></div>';
  container.insertAdjacentHTML('beforeend', h);
}}

return {{
  init: saInit,
  render: render,
  renderHighlights: renderHighlights,
  rRebalanceSim: rRebalanceSim,
  charts: charts,
  TD: TD,
  getTD: function() {{ return _dataLoaded && TD ? TD : null; }},
  loadData: loadData,
  isLoaded: function() {{ return _dataLoaded; }},
  updateLivePrice: function(ticker, price, chgPct) {{
    if (TD && TD[ticker] && TD[ticker].summary) {{
      TD[ticker].summary.p = price;
      TD[ticker].summary.r = chgPct / 100;
    }}
  }}
}};

}})();
"""
    return sa_data, sa_logic


# ═══════════════════════════════════════════════════════════════
# ANALYST TAB (AI 프롬프트 생성기)
# ═══════════════════════════════════════════════════════════════

def build_analyst_html():
    """애널리스트 탭 HTML 골격."""
    return '<div id="an-root"></div>'


def build_analyst_script():
    """애널리스트 탭 전체 IIFE 반환."""
    return r"""
var ANALYST = (function(){
"use strict";
var _inited = false;

/* ── 데이터 ── */
var KR_STOCKS = ['삼성전자','SK하이닉스','NAVER','카카오','LG에너지솔루션','현대차','셀트리온','삼성SDI','기아','POSCO홀딩스'];
var US_STOCKS = ['Apple','Microsoft','NVIDIA','Tesla','Google','Amazon','Meta','Netflix','AMD','Intel'];
var ETF_LIST = [
  {n:'SPY',d:'S&P500 추종',c:'🇺🇸',r:'미국'},{n:'QQQ',d:'나스닥100 기술주',c:'🇺🇸',r:'미국'},
  {n:'VOO',d:'S&P500 뱅가드 저비용',c:'🇺🇸',r:'미국'},{n:'VTI',d:'미국 전체시장',c:'🇺🇸',r:'미국'},
  {n:'SCHD',d:'배당성장 10년연속',c:'🇺🇸',r:'미국'},{n:'JEPI',d:'월배당 커버드콜',c:'🇺🇸',r:'미국'},
  {n:'VWO',d:'신흥국 전체',c:'🇺🇸',r:'글로벌'},{n:'VEA',d:'선진국(미국제외)',c:'🇺🇸',r:'글로벌'},
  {n:'KODEX 200',d:'KOSPI200',c:'🇰🇷',r:'한국'},{n:'KODEX 코스닥150',d:'코스닥150',c:'🇰🇷',r:'한국'},
  {n:'TIGER 미국S&P500',d:'S&P500 원화',c:'🇰🇷',r:'미국'},{n:'TIGER 미국나스닥100',d:'나스닥100 원화',c:'🇰🇷',r:'미국'},
  {n:'TIGER 미국배당다우존스',d:'SCHD 한국판',c:'🇰🇷',r:'미국'},{n:'KODEX 미국배당프리미엄액티브',d:'JEPI 한국판',c:'🇰🇷',r:'미국'},
  {n:'ARIRANG 고배당주',d:'국내 고배당',c:'🇰🇷',r:'한국'}
];
var SECTORS = [
  {id:'ai',nm:'AI',kw:'AI, GPU, LLM'},{id:'semiconductor',nm:'반도체',kw:'HBM, 메모리'},
  {id:'battery',nm:'2차전지',kw:'전기차 배터리'},{id:'bio',nm:'바이오',kw:'신약, CDMO'},
  {id:'platform',nm:'IT플랫폼',kw:'포털, 핀테크'},{id:'defense',nm:'방산',kw:'K-방산'},
  {id:'energy',nm:'에너지',kw:'수소, 원자력'},{id:'finance',nm:'금융',kw:'은행, 증권'},
  {id:'consumer',nm:'소비재',kw:'화장품, 식품'},{id:'auto',nm:'자동차',kw:'전기차, 자율주행'},
  {id:'shipbuilding',nm:'조선/해운',kw:'LNG선'},{id:'entertainment',nm:'엔터/게임',kw:'K-POP, 게임'},
  {id:'robot',nm:'로봇',kw:'휴머노이드'},{id:'travel',nm:'항공/여행',kw:'항공사'},
  {id:'construction',nm:'건설',kw:'SOC, 플랜트'},{id:'telecom',nm:'통신',kw:'5G, 배당'}
];

/* 용어 사전 (카테고리별) */
var TERMS = {
basic:[
{t:'PER',f:'주가수익비율',ds:'주가÷EPS. 10이하 저평가, 20이상 고평가 가능',dt:'업종마다 기준 다름. 은행 5~8배, IT 30배 이상도 흔함',ex:'삼성전자 PER 12배 vs 네이버 25배'},
{t:'PBR',f:'주가순자산비율',ds:'주가÷BPS. 1미만이면 청산가치보다 주가 낮음',dt:'가치투자자들이 PBR 1 미만을 "안전마진이 있다"고 봄',ex:'현대차 PBR 0.5배 → 자산 대비 저평가'},
{t:'ROE',f:'자기자본이익률',ds:'자기자본 대비 순이익 비율. 15% 이상 우수',dt:'워런 버핏: ROE 15% 이상 꾸준히 유지하는 기업에 투자하라',ex:'삼성전자 ROE 8% vs 애플 150%'},
{t:'EPS',f:'주당순이익',ds:'순이익÷발행주식수. 매년 증가→성장 기업',dt:'PER 계산의 기본. 배당 여력 판단에도 활용',ex:'삼성전자 EPS 6,000원, 주가 72,000원 → PER 12배'},
{t:'시가총액',f:'Market Cap',ds:'주가×발행주식수. 기업의 총 가치',dt:'대형주 10조+, 중형주 1~10조, 소형주 1조 미만',ex:'삼성전자 시총 400조원 = 코스피의 약 20%'},
{t:'배당수익률',f:'Dividend Yield',ds:'주당배당금÷주가×100',dt:'3~5% 적정, 5%+ 고배당. 너무 높으면 주가 급락한 경우 주의',ex:'SK텔레콤 배당수익률 6% → 1억 투자 시 연 600만원'},
{t:'거래량',f:'Trading Volume',ds:'하루 거래된 주식 수량',dt:'거래량 급증+상승=강한 매수세, 급증+하락=대량 매도',ex:'평소 10만주→오늘 100만주 → 무슨 일인지 확인!'},
{t:'52주 신고가/신저가',f:'52-Week High/Low',ds:'최근 1년간 최고/최저 주가',dt:'신고가 돌파=강한 상승신호, 신저가=약세 지속',ex:'삼성전자 52주 최고 88,000 최저 59,000'}
],
financial:[
{t:'부채비율',f:'Debt to Equity',ds:'총부채÷자기자본. 100% 이하 양호',dt:'금융업은 1000%+도 정상(예금이 부채). 업종별 기준 다름',ex:'현대차 부채비율 180% → 설비투자 때문'},
{t:'영업이익률',f:'Operating Margin',ds:'영업이익÷매출. 본업의 수익성',dt:'유통 2~5%, 제조 8~15%, IT/제약 20~40%',ex:'네이버 15% vs 이마트 3%'},
{t:'FCF',f:'잉여현금흐름',ds:'영업현금흐름-CAPEX. 진짜 남는 현금',dt:'워런 버핏: FCF가 진정한 기업가치',ex:'테슬라 순이익 100억, FCF 50억(공장건설 투자)'},
{t:'ROA',f:'총자산이익률',ds:'순이익÷총자산. 자산 활용 효율성',dt:'ROE vs ROA 격차가 크면 레버리지 사용 중',ex:'은행 ROA 0.5~1% → 자산이 커서 낮음'},
{t:'듀퐁분석',f:'DuPont Analysis',ds:'ROE를 순이익률×자산회전율×레버리지로 분해',dt:'같은 ROE 20%도 속이 다를 수 있음. 빚으로 만든 ROE 주의!',ex:'삼성전자: 마진12%×회전0.6×레버리지1.3 → 건전한 ROE'},
{t:'SWOT분석',f:'강점/약점/기회/위협',ds:'기업을 4가지로 나눠 투자 판단',dt:'강점+기회 크면 매수, 약점+위협 심하면 매도 고려',ex:'엔비디아: S=AI칩80%, W=게임의존, O=AI폭발, T=AMD추격'}
],
valuation:[
{t:'DCF',f:'현금흐름할인법',ds:'미래 수익을 현재가치로 할인해 기업가치 산출',dt:'가장 이론적으로 정확하나 가정에 따라 결과 큰 차이',ex:'5년간 매년 100억, 할인율10% → 약380억'},
{t:'EV/EBITDA',f:'기업가치/EBITDA',ds:'M&A에 많이 사용. 6배 이하 저평가',dt:'감가상각, 세금, 자본구조 차이 제거해 비교 용이',ex:'EV 1조, EBITDA 1000억 → 10배'},
{t:'PEG',f:'PER÷이익성장률',ds:'1 미만이면 성장 대비 저평가',dt:'피터 린치: PEG 1 이하 성장주를 찾아라',ex:'엔비디아 PER60, 성장80% → PEG 0.75'},
{t:'안전마진',f:'Margin of Safety',ds:'적정가 대비 할인율. 실수에 대비한 완충장치',dt:'벤저민 그레이엄/워런 버핏의 투자 초석',ex:'적정가 10만원, 안전마진25% → 7.5만원에 매수'},
{t:'경제적 해자',f:'Economic Moat',ds:'경쟁자가 넘볼 수 없는 방어막',dt:'브랜드, 전환비용, 네트워크효과, 원가우위, 무형자산',ex:'엔비디아: CUDA생태계 종속 = 강력한 해자'}
],
etf:[
{t:'ETF',f:'상장지수펀드',ds:'주식처럼 거래되는 펀드. 1주로 수백종목 분산투자',dt:'지수ETF, 섹터ETF, 채권ETF, 원자재ETF 등',ex:'SPY 1주 = S&P500 500개 기업 분산투자'},
{t:'운용보수',f:'Expense Ratio',ds:'ETF 연간 비용. 수익에서 자동 차감',dt:'0.03~0.10% 매우 저렴(VOO), 0.50%+ 비쌈',ex:'VOO 0.03% vs 국내펀드 1.5% → 50배 차이'},
{t:'추적오차',f:'Tracking Error',ds:'ETF수익률과 추종지수 차이',dt:'0.1%이하 우수, 0.5%+ 불량',ex:'S&P500 +10%, ETF +9.7% → 추적오차 0.3%'},
{t:'환헤지',f:'Currency Hedging',ds:'환율 변동 제거. (H)표시 ETF',dt:'단기→환헤지 고려, 장기→환노출(달러자산 분산)',ex:'S&P500+10%, 환율-5% → 환노출5%, 환헤지10%'},
{t:'분배금',f:'Distribution',ds:'ETF 보유종목 배당금을 투자자에게 분배',dt:'월배당(JEPI), 분기배당(SPY), 세금 15.4%',ex:'JEPI 월0.5% → 1억 투자시 월50만원'}
],
trading:[
{t:'손절',f:'Stop Loss',ds:'손실 제한 위해 일정 하락 시 매도',dt:'-50% 손실 → 원금회복에 +100% 필요!',ex:'100만원 매수, -10% 손절 → 최대 손실 10만원'},
{t:'분할매수',f:'DCA',ds:'정기적 일정 금액 매수. 타이밍 리스크 감소',dt:'주가 높을때 적게, 낮을때 많이 → 평단 자연 하락',ex:'매월 100만원 SPY 매수 12개월 → 평균가 매수'},
{t:'리밸런싱',f:'Rebalancing',ds:'목표 비중에서 벗어난 포트폴리오 재조정',dt:'자동으로 비싸면 팔고 싸면 사는 효과',ex:'주식60:채권40 → 주식70으로 변동 → 재조정'},
{t:'골든크로스',f:'Golden/Dead Cross',ds:'이동평균선 교차 매수/매도 신호',dt:'단기선이 장기선 상향돌파=매수, 하향=매도',ex:'삼성전자 5일선이 20일선 상향돌파 → 상승신호'},
{t:'RSI',f:'상대강도지수',ds:'0~100 사이. 70↑ 과매수, 30↓ 과매도',dt:'다이버전스: 주가↓+RSI↑=매수신호',ex:'삼성전자 RSI 25 → 과매도 → 반등 가능'},
{t:'볼린저밴드',f:'Bollinger Bands',ds:'이평선 중심 상하 표준편차 밴드',dt:'밴드 수축=변동성 감소→큰 움직임 예고',ex:'밴드 수축 후 상단 돌파+거래량↑ → 상승 시작'},
{t:'블랙스완',f:'Black Swan',ds:'극히 드물지만 엄청난 충격 사건',dt:'분산투자, 현금 유지, 레버리지 최소화로 대비',ex:'2020년 코로나 S&P500 한달 -34% → 대비한 투자자 저점매수'},
{t:'시나리오분석',f:'Scenario Analysis',ds:'Bull/Base/Bear 3가지 상황 가정',dt:'(확률×수익률) 합산 = 기대수익률',ex:'Bull20%×40 + Base50%×15 + Bear30%×(-30) = 6.5%'}
],
technical:[
{t:'이동평균선',f:'Moving Average',ds:'일정기간 평균 주가. 추세 파악 기본',dt:'정배열(5>20>60>120)=상승, 역배열=하락',ex:'60일선 위 올라옴 → 중기 상승 전환 신호'},
{t:'MACD',f:'이동평균수렴확산',ds:'단기-장기 이평 차이로 추세전환 포착',dt:'MACD>시그널=매수, MACD<시그널=매도',ex:'MACD 골든크로스+히스토그램 양수 → 강매수'},
{t:'OBV',f:'거래량누적지표',ds:'상승일 거래량 더하고 하락일 빼서 누적',dt:'OBV↑+주가횡보=세력 매집(상승 예고)',ex:'주가 3주 횡보+OBV 상승 → 곧 돌파 기대'},
{t:'ESG',f:'환경·사회·지배구조',ds:'비재무적 성과 평가. 장기 리스크 관리',dt:'MSCI AAA~CCC 등급. 기관투자자 기준',ex:'삼성전자 ESG A등급'},
{t:'인덱스펀드',f:'Index Fund',ds:'시장지수 그대로 추종. 워런 버핏 추천',dt:'10년 기준 80%+ 액티브펀드가 인덱스 못이김',ex:'유산의 90%는 S&P500 인덱스에 투자하라-버핏'}
]
};

/* ── 상태 변수 ── */
var curMenu = 'recommend';
var selStocks = [];
var prompt = '';
var history = [];
var expandedTerm = -1;
var opts = {
  recommend: {sector:'ai',type:'성장주',period:'6개월',growth:'30%',count:'5개'},
  morning: {newsType:'정치테마',market:'한국'},
  fundamental: {analysis:'종합분석',focus:'ROE'},
  valuation: {method:'상대가치',margin:'20%'},
  technical: {timeframe:'일봉',indicator:'이동평균',pattern:'추세'},
  qualitative: {focus:'경쟁우위'},
  risk: {scenario:'3가지',stopLoss:'15%'},
  portfolio: {amount:'1000만원',strategy:'균등분산'},
  etf: {goal:'장기투자',style:'패시브',region:'글로벌'}
};
var MENUS = [
  {id:'morning',nm:'🌅 모닝브리핑'},{id:'recommend',nm:'💡 종목발굴'},{id:'discovery',nm:'🔍 내종목'},
  {id:'fundamental',nm:'📊 재무'},{id:'valuation',nm:'💰 밸류'},{id:'technical',nm:'📈 기술적'},
  {id:'qualitative',nm:'🧠 질적'},{id:'risk',nm:'⚠️ 리스크'},{id:'portfolio',nm:'💼 포트폴리오'},
  {id:'etf',nm:'🏦 ETF'},{id:'terms',nm:'📖 용어'}
];

function $(id){return document.getElementById(id);}

function ts(){
  var d=new Date(),M=d.getMonth()+1,D=d.getDate(),h=d.getHours(),m=String(d.getMinutes()).padStart(2,'0');
  return d.getFullYear()+'년 '+M+'월 '+D+'일 '+h+':'+m+' (KST)';
}
function timeInst(){
  var t=ts();
  return '📅 분석 요청 시점: '+t+'\n⚠️ 중요: 반드시 '+t+' 기준 최신 데이터를 웹 검색으로 조회하세요.\n학습된 과거 데이터 사용 금지! 실시간 주가, 뉴스, 공시 정보를 검색해주세요.\n---\n';
}

/* ── 렌더링 ── */
function init(){
  if(_inited){ render(); return; }
  _inited = true;
  render();
}

function render(){
  var root = $('an-root'); if(!root) return;
  var h = '<div class="an-hd"><h2>🤖 AI 투자분석 Pro</h2><span class="an-sub">RICE 프롬프트 생성기</span></div>';
  /* 서브탭 */
  h += '<div class="an-tabs">';
  MENUS.forEach(function(m){
    h += '<div class="an-tab'+(curMenu===m.id?' on':'')+'" onclick="ANALYST.setMenu(\''+m.id+'\')">'+m.nm+'</div>';
  });
  h += '</div>';
  /* 바디 */
  h += '<div class="an-body">';
  h += renderMenu();
  h += '</div>';
  /* 프롬프트 결과 */
  if(prompt){
    h += '<div class="an-result"><div class="an-result-hd"><h3>✨ RICE 프롬프트</h3>';
    h += '<button class="an-copy-btn" onclick="ANALYST.copyPrompt()">📋 복사</button></div>';
    h += '<textarea class="an-textarea" id="an-prompt-ta" readonly>'+esc(prompt)+'</textarea>';
    h += '<div class="an-tip">💡 복사 후 → ChatGPT / Claude / Gemini에 붙여넣기!</div></div>';
  }
  /* 히스토리 */
  if(history.length>0){
    h += '<div class="an-hist"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">';
    h += '<span style="font-size:10px;font-weight:700;color:var(--t1)">📜 히스토리</span>';
    h += '<span style="font-size:8px;color:var(--dn);cursor:pointer" onclick="ANALYST.clearHist()">삭제</span></div>';
    history.forEach(function(e,i){
      h += '<div class="an-hist-item" onclick="ANALYST.loadHist('+i+')">';
      h += '<span class="an-hist-name">'+esc(e.menu)+'</span>';
      h += '<span class="an-hist-time">'+e.time+'</span>';
      h += '<span class="an-hist-load">불러오기→</span></div>';
    });
    h += '</div>';
  }
  root.innerHTML = h;
}

function esc(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

function setMenu(id){ curMenu=id; prompt=''; render(); }

/* ── 종목 선택 공통 ── */
function stockSelector(isETF){
  var h = '<div class="an-stock-wrap">';
  h += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">';
  h += '<span class="an-lbl">\uD83D\uDCCC '+(isETF?'ETF':'\uC885\uBAA9')+' \uC120\uD0DD</span>';
  h += '<div style="display:flex;gap:10px;align-items:center">';
  h += '<span style="font-size:9px;color:var(--blue);cursor:pointer;font-weight:700;background:rgba(91,141,238,0.1);padding:3px 6px;border-radius:4px" onclick="ANALYST.loadMyPortfolio()">\uD83D\uDCBC \uB0B4 \uBCF4\uC720\uC885\uBAA9 \uC804\uCCB4 \uB2F4\uAE30</span>';
  if(selStocks.length>0) h += '<span style="font-size:9px;color:var(--dn);cursor:pointer;font-weight:700;padding:3px 0" onclick="ANALYST.clearStocks()">\uCD08\uAE30\uD654</span>';
  h += '</div></div>';
  /* 직접 입력 */
  h += '<div class="an-input-row"><input class="an-input" id="an-custom-input" placeholder="'+(isETF?'ETF명':'종목명')+' 직접 입력" onkeydown="if(event.key===\'Enter\')ANALYST.addCustom()">';
  h += '<button class="an-btn-sm" onclick="ANALYST.addCustom()">추가</button></div>';
  /* 선택된 종목 */
  if(selStocks.length>0){
    h += '<div class="an-stock-chosen">';
    selStocks.forEach(function(s){
      h += '<span class="an-stock-tag">'+esc(s)+'<span class="x" onclick="ANALYST.toggle(\''+esc(s)+'\')">×</span></span>';
    });
    h += '</div>';
  }
  if(isETF){
    h += '<div class="an-lbl" style="margin-top:6px">🇺🇸 미국 ETF</div>';
    ETF_LIST.filter(function(e){return e.c==='🇺🇸';}).forEach(function(e){
      h += '<div class="an-etf-item'+(selStocks.indexOf(e.n)>=0?' on':'')+'" onclick="ANALYST.toggle(\''+e.n+'\')">';
      h += '<span class="an-etf-name">'+e.n+'</span><span class="an-etf-desc">'+e.d+'</span></div>';
    });
    h += '<div class="an-lbl" style="margin-top:6px">🇰🇷 한국 ETF</div>';
    ETF_LIST.filter(function(e){return e.c==='🇰🇷';}).forEach(function(e){
      h += '<div class="an-etf-item'+(selStocks.indexOf(e.n)>=0?' on':'')+'" onclick="ANALYST.toggle(\''+e.n+'\')">';
      h += '<span class="an-etf-name">'+e.n+'</span><span class="an-etf-desc">'+e.d+'</span></div>';
    });
  } else {
    h += '<div class="an-lbl" style="margin-top:6px">🇰🇷 한국</div><div class="an-stock-sel">';
    KR_STOCKS.forEach(function(s){ h += '<span class="an-stock-chip'+(selStocks.indexOf(s)>=0?' on':'')+'" onclick="ANALYST.toggle(\''+s+'\')">'+s+'</span>'; });
    h += '</div><div class="an-lbl" style="margin-top:4px">🇺🇸 미국</div><div class="an-stock-sel">';
    US_STOCKS.forEach(function(s){ h += '<span class="an-stock-chip'+(selStocks.indexOf(s)>=0?' on':'')+'" onclick="ANALYST.toggle(\''+s+'\')">'+s+'</span>'; });
    h += '</div>';
  }
  h += '</div>';
  return h;
}

function optGrid(label,options,curVal,optKey,cols){
  var h='<div class="an-card-title">'+label+'</div><div class="an-opt-grid c'+(cols||3)+'">';
  options.forEach(function(o){
    h+='<div class="an-opt'+(curVal===o?' on':'')+'" onclick="ANALYST.setOpt(\''+optKey+'\',\''+o+'\')">'+o+'</div>';
  });
  h+='</div>';
  return h;
}

function selectBox(label,options,curVal,optKey){
  var h='<div class="an-card-title">'+label+'</div><select class="an-select" onchange="ANALYST.setOpt(\''+optKey+'\',this.value)">';
  options.forEach(function(o){ h+='<option'+(curVal===o?' selected':'')+'>'+o+'</option>'; });
  h+='</select>';
  return h;
}

/* ── 메뉴별 렌더 ── */
function renderMenu(){
  var h='';
  switch(curMenu){
  case 'morning':
    h+='<div class="an-card">'+optGrid('📰 뉴스 유형',['정치테마','기업공시','글로벌이슈','테마급등'],opts.morning.newsType,'morning.newsType',2);
    h+=optGrid('🌍 시장',['한국','미국','전체'],opts.morning.market,'morning.market',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()">🌅 모닝브리핑 프롬프트 생성</button>';
    break;
  case 'recommend':
    h+='<div class="an-card"><div class="an-card-title">🎯 섹터</div><div class="an-opt-grid c5">';
    SECTORS.forEach(function(s){
      h+='<div class="an-opt'+(opts.recommend.sector===s.id?' on':'')+'" onclick="ANALYST.setOpt(\'recommend.sector\',\''+s.id+'\')">'+s.nm+'</div>';
    });
    h+='</div>'+optGrid('📊 투자 스타일',['성장주','가치주','배당주'],opts.recommend.type,'recommend.type',3);
    h+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">'+selectBox('📅 기간',['1개월','3개월','6개월','1년','2년','3년','5년'],opts.recommend.period,'recommend.period');
    h+=selectBox('📈 목표',['5%','10%','15%','20%','30%','50%','100%'],opts.recommend.growth,'recommend.growth')+'</div>';
    h+=optGrid('🔢 개수',['3개','5개','7개','10개'],opts.recommend.count,'recommend.count',4)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()">✨ RICE 프롬프트 생성</button>';
    break;
  case 'discovery':
    h+=stockSelector(false);
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 내종목 RICE 프롬프트 생성</button>';
    break;
  case 'fundamental':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('분석 유형',['종합분석','비교분석','심층분석'],opts.fundamental.analysis,'fundamental.analysis',3);
    h+=optGrid('중점 영역',['ROE','안정성','성장성'],opts.fundamental.focus,'fundamental.focus',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 재무분석 프롬프트 생성</button>';
    break;
  case 'valuation':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('평가 방법',['상대가치','절대가치','종합'],opts.valuation.method,'valuation.method',3);
    h+=optGrid('안전마진',['10%','20%','30%','40%'],opts.valuation.margin,'valuation.margin',4)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 밸류에이션 프롬프트 생성</button>';
    break;
  case 'technical':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('타임프레임',['일봉','주봉','월봉'],opts.technical.timeframe,'technical.timeframe',3);
    h+=optGrid('주요 지표',['이동평균','모멘텀','거래량'],opts.technical.indicator,'technical.indicator',3);
    h+=optGrid('패턴',['추세','패턴','캔들'],opts.technical.pattern,'technical.pattern',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 기술적분석 프롬프트 생성</button>';
    break;
  case 'qualitative':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('중점 영역',['경쟁우위','경영진','ESG'],opts.qualitative.focus,'qualitative.focus',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 질적분석 프롬프트 생성</button>';
    break;
  case 'risk':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('시나리오 수',['3가지','5가지','7가지'],opts.risk.scenario,'risk.scenario',3);
    h+=optGrid('손절 기준',['10%','15%','20%','25%'],opts.risk.stopLoss,'risk.stopLoss',4)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 리스크분석 프롬프트 생성</button>';
    break;
  case 'portfolio':
    h+=stockSelector(false);
    h+='<div class="an-card">'+optGrid('투자금액',['1000만원','3000만원','5000만원','1억원'],opts.portfolio.amount,'portfolio.amount',2);
    h+=optGrid('배분 전략',['균등분산','차등배분','집중투자'],opts.portfolio.strategy,'portfolio.strategy',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ 포트폴리오 프롬프트 생성</button>';
    break;
  case 'etf':
    h+=stockSelector(true);
    h+='<div class="an-card">'+optGrid('투자 목적',['장기투자','배당수익','단기트레이딩'],opts.etf.goal,'etf.goal',3);
    h+=optGrid('투자 스타일',['패시브','액티브','테마'],opts.etf.style,'etf.style',3);
    h+=optGrid('투자 지역',['미국','한국','글로벌'],opts.etf.region,'etf.region',3)+'</div>';
    h+='<button class="an-gen-btn" onclick="ANALYST.gen()"'+(selStocks.length===0?' disabled':'')+'>✨ ETF분석 프롬프트 생성</button>';
    break;
  case 'terms':
    h+=renderTerms();
    break;
  }
  return h;
}

/* ── 용어 사전 ── */
var termCat = 'basic';
function renderTerms(){
  var cats = [{id:'basic',nm:'기본'},{id:'financial',nm:'재무'},{id:'valuation',nm:'밸류'},{id:'etf',nm:'ETF'},{id:'trading',nm:'매매'},{id:'technical',nm:'기술적'}];
  var h = '<div class="an-tabs" style="margin-bottom:8px">';
  cats.forEach(function(c){
    h+='<div class="an-tab'+(termCat===c.id?' on':'')+'" onclick="ANALYST.setTermCat(\''+c.id+'\')">'+c.nm+'</div>';
  });
  h+='</div>';
  var items = TERMS[termCat] || [];
  items.forEach(function(it,i){
    h+='<div class="an-term-item">';
    h+='<div class="an-term-hd" onclick="ANALYST.toggleTerm('+i+')">';
    h+='<div><span class="an-term-name">'+it.t+'</span><span class="an-term-full">'+it.f+'</span>';
    h+='<div class="an-term-desc">'+it.ds+'</div></div>';
    h+='<span class="an-term-arrow">'+(expandedTerm===i?'▲':'▼')+'</span></div>';
    if(expandedTerm===i){
      h+='<div class="an-term-detail"><div class="detail-box">'+esc(it.dt)+'</div>';
      if(it.ex) h+='<div class="example-box"><b>💡 예시:</b> '+esc(it.ex)+'</div>';
      h+='</div>';
    }
    h+='</div>';
  });
  return h;
}

/* ── 옵션 세팅 ── */
function setOpt(key,val){
  var parts=key.split('.'); opts[parts[0]][parts[1]]=val; render();
}
function toggle(name){
  var i=selStocks.indexOf(name);
  if(i>=0) selStocks.splice(i,1); else selStocks.push(name);
  render();
}
function addCustom(){
  var inp=$('an-custom-input'); if(!inp) return;
  var v=inp.value.trim(); if(v && selStocks.indexOf(v)<0){ selStocks.push(v); }
  inp.value=''; render();
}
function clearStocks(){ selStocks=[]; render(); }
function loadMyPortfolio(){
  var myTickers=[];
  if(window.mainData) window.mainData.forEach(function(a){if(a.ticker&&myTickers.indexOf(a.ticker)<0) myTickers.push(a.ticker);});
  if(window.otherData) window.otherData.forEach(function(a){if(a.ticker&&myTickers.indexOf(a.ticker)<0) myTickers.push(a.ticker);});
  if(myTickers.length===0){alert('\uBCF4\uC720 \uC911\uC778 \uC790\uC0B0\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.');return;}
  selStocks=myTickers; render();
}
function setTermCat(c){ termCat=c; expandedTerm=-1; render(); }
function toggleTerm(i){ expandedTerm = expandedTerm===i ? -1 : i; render(); }

/* ── 프롬프트 생성 ── */
function gen(){
  var needStocks = ['discovery','fundamental','valuation','technical','qualitative','risk','portfolio','etf'].indexOf(curMenu)>=0;
  if(needStocks && selStocks.length===0){ alert('종목을 선택해주세요!'); return; }
  var list = selStocks.join(', ');
  var T = timeInst();

  switch(curMenu){
  case 'morning': prompt = genMorning(T); break;
  case 'recommend': prompt = genRecommend(T); break;
  case 'discovery': prompt = genDiscovery(T,list); break;
  case 'fundamental': prompt = genFundamental(T,list); break;
  case 'valuation': prompt = genValuation(T,list); break;
  case 'technical': prompt = genTechnical(T,list); break;
  case 'qualitative': prompt = genQualitative(T,list); break;
  case 'risk': prompt = genRisk(T,list); break;
  case 'portfolio': prompt = genPortfolio(T,list); break;
  case 'etf': prompt = genETF(T,list); break;
  }

  /* 히스토리 저장 */
  var menuNm = MENUS.find(function(m){return m.id===curMenu;});
  history.unshift({menu:menuNm?menuNm.nm:'',prompt:prompt,time:new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'})});
  if(history.length>10) history.pop();
  render();
  setTimeout(function(){
    var ta=$('an-prompt-ta');
    var panel=document.getElementById('analyst-panel');
    if(ta && panel){ panel.scrollTop=ta.offsetTop-20; }
  },120);
}

function copyPrompt(){
  var ta=$('an-prompt-ta'); if(!ta) return;
  ta.focus(); ta.select();
  document.execCommand('copy');
  alert('✅ 복사 완료!\nChatGPT, Claude, Gemini에 붙여넣기 하세요.');
}
function clearHist(){ history=[]; render(); }
function loadHist(i){
  var e=history[i]; if(!e) return;
  prompt=e.prompt; render();
  setTimeout(function(){
    var ta=$('an-prompt-ta');
    var panel=document.getElementById('analyst-panel');
    if(ta && panel){ panel.scrollTop=ta.offsetTop-20; }
  },120);
}

/* ═══ 프롬프트 생성 함수들 ═══ */

function genMorning(T){
  var o=opts.morning;
  return T+'🌅 오늘의 단타 모닝브리핑\n\nR (Role)\n10년 경력 단기 트레이딩 전문가. 테마주/이슈 선점 투자 전문.\n\nI (Instruction)\n오늘 주식장 시작 전 단타 매매 핵심 정보를 분석해주세요.\n🔴 필수: 최근 24시간 뉴스를 Web Search로 검색. 현재가 실시간 조회.\n\n분석 조건:\n- 뉴스 유형: '+o.newsType+'\n- 시장: '+o.market+'\n\n찾아야 할 정보:\n1. 정치인 발언/행동 → 수혜주\n2. 새벽 공시 (계약, 실적, M&A)\n3. 해외 시장 뉴스 → 국내 영향\n4. SNS/커뮤니티 화제 테마\n5. 시간외 급등/급락 종목\n\nC (Context)\n장 시작 전 빠른 의사결정. 단타(당일~2-3일) 관점.\n\nE (Example) - 출력 형식\n📰 오늘의 핵심 뉴스 & 수혜주\n---\n🔥 1순위: [뉴스 헤드라인]\n| 항목 | 내용 |\n|------|------|\n| 뉴스 요약 | [1-2줄] |\n| 수혜 종목 | [종목명] (코드) |\n| 연결 고리 | [왜 이 종목?] |\n| 현재가 | [실시간 조회] |\n| 매매 전략 | 시초가매수/눌림목/관망 |\n| 목표가 | [단기 목표] |\n| 손절가 | [리스크 관리] |\n---\n(2순위, 3순위 동일 형식)\n\n📊 테마별 정리 | ⚠️ 주의 종목 | 🎯 전략 요약 | 📅 오늘 일정';
}

function genRecommend(T){
  var o=opts.recommend;
  var sec=SECTORS.find(function(s){return s.id===o.sector;});
  var sn=sec?sec.nm:o.sector, kw=sec?sec.kw:'';
  return T+sn+' '+o.type+' 발굴\n\nR (Role)\n20년 경력 '+sn+' 섹터 전문 애널리스트. CFA 보유.\n\nI (Instruction)\n다음 조건에 맞는 '+o.count+'을 발굴해주세요:\n- 섹터: '+sn+' ('+kw+')\n- 스타일: '+o.type+'\n- 기간: '+o.period+'\n- 목표: '+o.growth+' 이상\n- 제외: 2년 연속 적자, 감사의견 비적정, 시총 1000억 미만\n\nC (Context)\n현재 시장 상황과 '+sn+' 업황 고려. 구조적 성장 중시.\n\nE (Example)\n🏆 '+sn+' '+o.type+' TOP '+o.count+'\n---\n1순위: [종목명]\n| 항목 | 내용 |\n|------|------|\n| 기본 정보 | 현재가, 시총, 52주 고/저 |\n| 밸류에이션 | PER, PBR, PEG |\n| 수익성 | ROE, 영업이익률 |\n| 성장성 | 매출 CAGR, 이익 CAGR |\n\n투자포인트 / 리스크 / 매매전략(적정가,목표가,손절가)\n---\n(2순위~'+o.count+' 동일)\n\n📊 종합 비교표\n🎯 최종 추천: 가장 확신 높은 종목 + 이유';
}

function genDiscovery(T,list){
  var tickers=list.split(',').map(function(s){return s.trim().toUpperCase();});
  var quantContext='';
  tickers.forEach(function(tk){
    if(typeof SCORING_DATA!=='undefined'&&SCORING_DATA[tk]&&SCORING_DATA[tk].total!==undefined){
      var d=SCORING_DATA[tk];
      quantContext+='  \u25A0 '+tk+' : 총점 '+d.total+'점 (기술: '+d.tech+'/30, 가치: '+d.fund+'/30, 수급: '+d.supply+'/20, 심리: '+d.sent+'/20)\n';
    }else{
      quantContext+='  \u25A0 '+tk+' : 퀀트 점수 데이터 없음\n';
    }
  });
  return T+'관심종목 퀀트 융합 심층 분석\n\n'
    +'R (Role) - 역할\n'
    +'당신은 글로벌 투자은행의 수석 리서치 애널리스트이자 퀀트(Quant) 포트폴리오 매니저입니다.\n'
    +'- 10년 이상 기업 분석 및 알고리즘 트레이딩 경력\n'
    +'- 다수의 베스트 애널리스트 수상 경력\n'
    +'- 정량적 데이터(퀀트 스코어)와 정성적 분석(비즈니스 모델, 매크로)을 결합하여 통찰력 있는 결론을 도출하는 데 전문\n\n'
    +'C (Context) - 맥락\n'
    +'- 투자 기간: 6개월 ~ 1년\n'
    +'- 투자자 성향: 중위험 중수익 추구\n'
    +'- 시장 대비 초과 수익 목표\n'
    +'- [나의 대시보드 퀀트 스코어 정보]\n'
    +'  현재 내가 자체 구축한 분석 시스템에서 산출된 종목별 퀀트 스코어는 다음과 같습니다. 이 데이터를 반드시 분석의 핵심 근거로 활용하세요.\n'
    +quantContext+'\n'
    +'I (Instruction) - 지시사항\n'
    +'다음 종목들을 위 퀀트 점수와 실시간 데이터를 결합하여 심층 분석해주세요.\n'
    +'분석 대상: '+list+'\n\n'
    +'\uD83D\uDD34 필수 - 실시간 데이터 조회:\n'
    +'각 종목의 현재 주가, 시가총액, 52주 최고/최저가를 웹에서 실시간 조회하여 분석에 포함해주세요.\n\n'
    +'분석 항목:\n'
    +'1. 기업 개요 및 사업 모델 요약\n'
    +'2. 실시간 주가 및 밸류에이션 점검\n'
    +'3. 퀀트 스코어 해석 및 펀더멘털 대조 (제공된 퀀트 점수의 강/약점 원인 분석)\n'
    +'4. SWOT 분석 (정성적 모멘텀 포함)\n'
    +'5. 종합 투자 매력도 및 매수/매도 타이밍 제안\n\n'
    +'E (Example) - 출력 형식\n\n'
    +'\uD83D\uDCCB 종목별 분석\n\n'
    +'---\n'
    +'1. [종목명]\n\n'
    +'기업 개요 (2줄)\n'
    +'[사업 모델, 시장 위치, 핵심 경쟁력]\n\n'
    +'핵심 지표 및 퀀트 진단\n'
    +'| 지표 | 수치 | 업종 평균 | 평가 |\n'
    +'|------|------|-----------|------|\n'
    +'| PER | | | |\n'
    +'| PBR | | | |\n'
    +'| ROE | | | |\n'
    +'| 퀀트 총점 | [입력된 총점]/100 | - | [강세/중립/약세] |\n'
    +'| 퀀트 상세 | 기술 [ ], 가치 [ ], 수급 [ ], 심리 [ ] | - | [특이사항 진단] |\n\n'
    +'SWOT 분석\n'
    +'- S (강점): [자체 경쟁력 및 퀀트 점수 상의 강점 반영]\n'
    +'- W (약점): [리스크 및 퀀트 점수 상의 약점 반영]\n'
    +'- O (기회): [최신 뉴스 기반 모멘텀]\n'
    +'- T (위협): [매크로/경쟁사 위협]\n\n'
    +'투자 판단\n'
    +'| 항목 | 평가 |\n'
    +'|------|------|\n'
    +'| 현재 밸류에이션 | 저평가 / 적정 / 고평가 |\n'
    +'| 적정 주가 | 원 |\n'
    +'| 목표 주가 (1년) | 원 (+__%) |\n'
    +'| 투자 매력도 | \u2B50\u2B50\u2B50\u2B50\u2606 (4/5) |\n'
    +'| 투자 의견 | 강력매수 / 매수 / 중립 / 매도 |\n'
    +'| 매매 타이밍 | [퀀트 기술/수급 점수를 바탕으로 한 단기 진입/청산 시점 조언] |\n\n'
    +'---\n'
    +'(나머지 종목 동일 형식)\n\n'
    +'---\n'
    +'\uD83D\uDCCA 종합 비교표\n\n'
    +'| 종목 | 퀀트점수 | 매력도 | 목표수익 | 리스크 | 추천순위 |\n'
    +'|------|----------|--------|----------|--------|----------|\n'
    +'| [종목명] | [점수] | \u2B50\u2B50\u2B50\u2606\u2606 | +__% | 상/중/하 | |\n\n'
    +'\uD83C\uDFC6 투자 우선순위\n'
    +'1. 1순위: [종목] - [퀀트 점수와 실시간 모멘텀을 종합한 핵심 이유]\n'
    +'2. 2순위: [종목] - [핵심 이유]\n'
    +'3. 3순위: [종목] - [핵심 이유]';
}

function genFundamental(T,list){
  var o=opts.fundamental;
  var focusDetail={ROE:'ROE 듀퐁분석: 순이익률×자산회전율×레버리지 분해',안정성:'부채비율, 이자보상배율, 유동비율, 순부채/EBITDA 집중',성장성:'매출/이익 CAGR(3~5년), 신사업 매출 비중'};
  return T+'재무 분석: '+o.analysis+'\n\nR (Role)\nCFA 보유 재무분석 전문가. Big4 회계법인 10년 경력.\n\nI (Instruction)\n분석 대상: '+list+'\n⚠️ 현재 주가/재무지표 웹 실시간 조회.\n분석 유형: '+o.analysis+'\n중점: '+o.focus+' — '+(focusDetail[o.focus]||'')+'\n\nC (Context)\n최근 3~5년 재무제표 기반. 업종 평균 대비 상대 평가.\n\nE (Example)\n📊 재무 분석 결과\n---\n[종목명]\n1. 수익성: ROE, ROA, 영업이익률, 순이익률 (최근/전년/3년평균/업종평균)\n'+(o.focus==='ROE'?'   → ROE 듀퐁분해: 순이익률, 자산회전율, 레버리지\n':'')+
'2. 안정성: 부채비율, 유동비율, 이자보상배율\n3. 성장성: 매출/영업이익/순이익 CAGR, YoY\n---\n📈 종합 재무 점수 (수익성/30 + 안정성/30 + 성장성/40 = /100)\n💡 투자 시사점';
}

function genValuation(T,list){
  var o=opts.valuation;
  return T+'밸류에이션 분석: '+o.method+'\n\nR (Role)\n밸류에이션 전문 퀀트 애널리스트. M&A 자문 경력.\n\nI (Instruction)\n분석 대상: '+list+'\n⚠️ 현재 PER, PBR, EPS, BPS 등 웹 실시간 조회.\n평가 방법: '+o.method+'\n안전마진: '+o.margin+'\n\nC (Context)\n보수적 가정. 업종 특성/성장 단계/시장 상황 반영.\n\nE (Example)\n💰 밸류에이션\n---\n[종목명]\n'+
(o.method!=='절대가치'?'PER 기반: 적정PER × 예상EPS\nPBR 기반: 적정PBR × BPS\n':'')+(o.method!=='상대가치'?'DCF: 5년 FCF 추정, WACC 9%, 영구성장률 2%\n':'')+
'종합 적정가 (가중평균)\n---\n📊 최종 요약\n| 종목 | 현재가 | 적정가 | 안전마진(-'+o.margin+') | 목표가 | 상승여력 | 의견 |\n🎯 가장 저평가 / 매수 적기 / 관망 종목';
}

function genTechnical(T,list){
  var o=opts.technical;
  var indGuide={이동평균:'이평선 배열(정/역), 골든크로스/데드크로스, 그랜빌 8법칙, 이격도',모멘텀:'RSI(14), MACD(12,26,9), 스토캐스틱(14,3,3), 다이버전스',거래량:'거래량-가격 패턴 4가지, OBV 분석, 세력 수급'};
  return T+'기술적 분석: '+o.indicator+'\n\nR (Role)\n'+o.indicator+' 분석 전문가. 20년 매매전략 연구.\n\nI (Instruction)\n분석 대상: '+list+'\n⚠️ 현재 주가 및 '+o.indicator+' 관련 지표 웹 실시간 조회.\n🔴 모든 분석에 구체적 가격 필수 표기!\n\n분석 포인트:\n'+(indGuide[o.indicator]||'')+'\n\nC (Context)\n타임프레임: '+o.timeframe+', 패턴: '+o.pattern+'\n\nE (Example)\n📊 '+o.indicator+' 분석 결과\n---\n[종목명] ('+o.timeframe+')\n📍 현재가: ₩______ (필수 조회)\n\n지표 현황 테이블 + 신호 분석 + 매매 판단\n---\n🎯 종합: 종목별 현재가/주요지표/매매판단/목표가/손절가';
}

function genQualitative(T,list){
  var o=opts.qualitative;
  var focusGuide={경쟁우위:'5가지 해자: 브랜드, 전환비용, 네트워크효과, 원가우위, 무형자산',경영진:'실적달성력, 자본배분, 주주친화, 지분율, 커뮤니케이션',ESG:'E(환경): 탄소/재생에너지, S(사회): 근로환경/다양성, G(지배구조): 이사회 독립성'};
  return T+'질적 분석: '+o.focus+'\n\nR (Role)\n'+o.focus+' 전문 분석가. 20년 경력.\n\nI (Instruction)\n분석 대상: '+list+'\n⚠️ 시장점유율, 경쟁사, 경영진 정보 웹 조회.\n\n분석 포인트:\n'+(focusGuide[o.focus]||'')+'\n\nC (Context)\n재무제표에 안 나타나는 질적 가치. 장기 경쟁력/지속가능성.\n\nE (Example)\n🏰 '+o.focus+' 분석 결과\n---\n[종목명]\n유형별 상세 평가 테이블\n종합 점수 + 등급\n---\n💡 핵심 인사이트: 가장 강력한 종목 + 주의 필요 종목';
}

function genRisk(T,list){
  var o=opts.risk;
  return T+'리스크 분석 & 시나리오\n\nR (Role)\n글로벌 자산운용사 CRO. 시나리오/스트레스테스트 전문.\n\nI (Instruction)\n분석 대상: '+list+'\n🔴 현재 주가, 52주 변동폭, 베타, 최근 뉴스 웹 조회.\n시나리오: '+o.scenario+'\n손절: -'+o.stopLoss+'\n\n리스크 유형: 기업고유(실적/경쟁/규제), 섹터, 매크로\n\nC (Context)\n최악까지 고려. Risk/Reward 관점.\n\nE (Example)\n⚠️ 리스크 분석\n---\n[종목명]\n1. 변동성: 52주 범위, 베타, MDD\n2. 기업 리스크: 실적/경쟁/규제/재무/키맨\n3. 시나리오 ('+o.scenario+'):\n| 시나리오 | 확률 | 목표가 | 수익률 | 조건 |\n기대수익률 = Σ(확률×수익률)\n4. 손절/익절 전략\n---\n📊 종합: Risk/Reward 비율 → 투자 판단';
}

function genPortfolio(T,list){
  var o=opts.portfolio;
  return T+'포트폴리오 구성 전략\n\nR (Role)\nPB 출신 포트폴리오 매니저. 자산배분 전문.\n\nI (Instruction)\n구성 종목: '+list+'\n🔴 현재 주가 웹 조회 → 실제 매수 가능 수량 계산.\n투자금: '+o.amount+'\n전략: '+o.strategy+'\n\nC (Context)\n분산투자 리스크 관리. 섹터/스타일 균형.\n\nE (Example)\n💼 포트폴리오 구성안\n1. 종목별 점수 (매력도/리스크/확신도)\n2. 비중 배분: 역할(핵심/주력/위성), 비중%, 투자금, 매수가, 목표가\n3. 분할 매수: 1차 50%(즉시), 2차 30%(-5%), 3차 20%(-10%)\n4. 리밸런싱: 분기 1회, ±5%p 이탈시\n---\n📈 기대 수익 시나리오 (낙관/기본/비관)\n✅ 실행 체크리스트';
}

function genETF(T,list){
  var o=opts.etf;
  return T+'ETF 분석 & 포트폴리오\n\nR (Role)\nETF 전문 투자 자문가. 패시브/자산배분 전문.\n\nI (Instruction)\n분석 대상: '+list+'\n🔴 현재가, 운용보수, 배당수익률, 1년 수익률 웹 조회.\n목적: '+o.goal+' / 스타일: '+o.style+' / 지역: '+o.region+'\n\nC (Context)\n장기 복리 극대화. 운용보수/추적오차 중요.\n\nE (Example)\n🏦 ETF 분석\n---\n[ETF명]\n기본정보: 추종지수, 운용사, 보수, 거래량\n수익률: 1M/3M/1Y/3Y/5Y\n배당: 수익률, 주기, 최근 분배금\nTOP10 종목, 품질(추적오차/괴리율/MDD)\n---\n📊 비교표 | 💼 포트폴리오 제안 (Core/Satellite)\n🎯 추천: 목적별 최적 + 비용최고 + 배당최고';
}

return {
  init:init, render:render, setMenu:setMenu, gen:gen,
  copyPrompt:copyPrompt, clearHist:clearHist, loadHist:loadHist,
  toggle:toggle, addCustom:addCustom, clearStocks:clearStocks,
  loadMyPortfolio:loadMyPortfolio,
  setOpt:setOpt, setTermCat:setTermCat, toggleTerm:toggleTerm
};
})();
"""


def generate_merged_html(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas=None, df=None, fx_rate=None, analyst=None):
    """Dashboard B + SA 통합 HTML 생성."""

    # 1. SA 콘텐츠 생성
    sa_html = build_sa_html(data_date, len(td))
    sa_right = build_sa_right_html()
    sa_data, sa_logic = build_sa_script(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas)

    # 2. AI 브리핑 생성
    if df is not None:
        ai_briefing = build_ai_briefing(df, fc, ev, data_date)
    else:
        ai_briefing = '<span style="color:var(--t3);font-size:9px">브리핑 데이터 없음</span>'

    # 3. 모멘텀 노트 생성
    if df is not None:
        momentum_html = build_momentum_notes(df, fc, data_date)
    else:
        momentum_html = '<div style="padding:10px;font-size:9px;color:var(--t3)">데이터 로딩 중...</div>'

    # 4. 머지 템플릿 읽기
    with open(MERGED_TEMPLATE, encoding="utf-8") as f:
        merged = f.read()

    # 5. 플레이스홀더 치환
    merged = merged.replace("<!-- __SA_HTML__ -->", sa_html)
    merged = merged.replace("<!-- __SA_RIGHT_HTML__ -->", sa_right)
    merged = merged.replace("/* __SA_DATA_PLACEHOLDER__ */", sa_data)
    merged = merged.replace("/* __SA_LOGIC_PLACEHOLDER__ */", sa_logic)
    merged = merged.replace("<!-- __AI_BRIEFING_PLACEHOLDER__ -->", ai_briefing)
    merged = merged.replace("<!-- __MOMENTUM_PLACEHOLDER__ -->", momentum_html)

    # 5-b. 애널리스트 탭 콘텐츠 주입
    analyst_html = build_analyst_html()
    analyst_script = build_analyst_script()
    merged = merged.replace("<!-- __ANALYST_HTML__ -->", analyst_html)
    merged = merged.replace("/* __ANALYST_SCRIPT_PLACEHOLDER__ */", analyst_script)
    log.info("Analyst tab injected (HTML + IIFE script)")

    # 6. 환율 초기값 업데이트
    if fx_rate and fx_rate > 0:
        merged = re.sub(
            r'let\s+FX\s*=\s*[\d.]+;',
            f'let   FX   = {fx_rate};',
            merged, count=1
        )
        log.info(f"FX initial value updated to {fx_rate}")

    # 7. mainData / otherData 자동 업데이트
    if df is not None and analyst is not None:
        main_js, other_js = build_asset_update(df, fc, analyst, fx_rate or 1440)
        merged = re.sub(
            r'var mainData\s*=\s*\[[\s\S]*?\];',
            main_js,
            merged, count=1
        )
        merged = re.sub(
            r'var otherData\s*=\s*\[[\s\S]*?\];',
            other_js,
            merged, count=1
        )
        log.info("mainData / otherData updated with latest prices")

    # 8. ALL_TICKERS_INFO — tickers.js로 분리됨 (플레이스홀더 제거)
    merged = merged.replace("/* __ALL_TICKERS_PLACEHOLDER__ */", "/* ALL_TICKERS_INFO loaded from tickers.js via <script> tag */")
    log.info("ALL_TICKERS_INFO → tickers.js (loaded via <script> tag)")

    # 9. 데이터 날짜 표시 — dateLabel 요소를 topbar에 추가
    from datetime import datetime as _dt
    try:
        dt_obj = _dt.strptime(data_date[:10], "%Y-%m-%d")
        weekdays = ["월", "화", "수", "목", "금", "토", "일"]
        date_display = f'{dt_obj.month}/{dt_obj.day}({weekdays[dt_obj.weekday()]})'
    except Exception:
        date_display = data_date[:10]
    date_badge = (
        f'<div class="tb-stat">'
        f'<div class="tb-stat-lbl">데이터 기준</div>'
        f'<div class="tb-stat-val" id="dateLabel" style="color:var(--t2);font-size:10px">{date_display}</div>'
        f'</div>'
    )
    merged = merged.replace(
        '<!-- 동기화 표시 제거됨 -->',
        f'<!-- 데이터 날짜 -->\n  <div class="tb-divider"></div>\n  {date_badge}'
    )

    return merged


def generate_html(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas=None):
    """완성된 대시보드 HTML 생성 (standalone SA — 모바일용)."""

    # 원본 HTML 읽기
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        html = f.read()

    # CSS 섹션 추출 (</style> 까지)
    style_end = html.find("</style>") + len("</style>")
    css = html[:style_end]

    # 원본 JS에서 유지할 부분 추출 (UTILS ~ INIT 이전)
    script_start = html.find("<script>", html.find("chart.js")) + len("<script>")
    script_end = html.find("</script>", script_start)
    orig_js = html[script_start:script_end]

    # UTILS 시작 ~ INIT 시작 추출
    sep = "// ═══════════════════════════════════════════════"
    utils_idx = orig_js.find(sep + "\n// UTILS")
    init_idx = orig_js.find(sep + "\n// INIT")
    if utils_idx < 0 or init_idx < 0:
        log.error("원본 HTML에서 JS 섹션 마커를 찾을 수 없습니다.")
        sys.exit(1)

    kept_js = orig_js[utils_idx:init_idx]

    # ALPHA VANTAGE FETCH / SMART SCAN ENGINE 제거
    # 이 섹션들을 no-op으로 교체
    fetch_start = kept_js.find(sep + "\n// ALPHA VANTAGE FETCH")
    scoring_start = kept_js.find(sep + "\n// SCORING ENGINE")
    if fetch_start >= 0 and scoring_start >= 0:
        kept_js = kept_js[:fetch_start] + "\n// (Offline mode - fetch/scan removed)\n" + kept_js[scoring_start:]

    # calcAllSP5Sigs에서 하드코딩된 sp500tickers 수정
    kept_js = kept_js.replace(
        'var sp500tickers=["NVDA","TSM","MU","CEG"];',
        'var sp500tickers=[];for(var pi=0;pi<MYPORT.length;pi++){if(SP5_INFO[MYPORT[pi]])sp500tickers.push(MYPORT[pi])}'
    )

    # noDataMsg 수정 (API 안내 제거)
    old_nodata = 'function noDataMsg(msg){return\'<div class="cd" style="text-align:center;padding:30px"><div style="font-size:28px;margin-bottom:10px">📡</div><div style="font-size:12px;font-weight:700;margin-bottom:6px">\'+msg+\'</div><div style="font-size:9px;color:var(--dim);line-height:1.7">① API 키 입력 → "🔄 전체 스캔" 클릭<br>② 포트폴리오(7) → 섹터ETF(11) → S&P500 순차 스캔<br><a href="https://www.alphavantage.co/support/#api-key" target="_blank" style="color:var(--b)">▸ 무료 API 키 발급</a></div></div>'
    new_nodata = 'function noDataMsg(msg){return\'<div class="cd" style="text-align:center;padding:30px"><div style="font-size:28px;margin-bottom:10px">📡</div><div style="font-size:12px;font-weight:700;margin-bottom:6px">\'+msg+\'</div><div style="font-size:9px;color:var(--dim)">오프라인 데이터가 로드되지 않았습니다.</div></div>'
    kept_js = kept_js.replace(old_nodata, new_nodata)

    # JSON 직렬화
    j = lambda obj: json.dumps(obj, ensure_ascii=False, separators=(",", ":"))

    # 전체 SP5 티커 리스트 (SMAP에 없는 나머지)
    smap_tickers = set()
    for s in smap:
        smap_tickers.update(s["st"])
    remaining = [t for t in all_sp5_tickers if t not in smap_tickers and t not in MYPORT and t not in SECTOR_ETFS]

    # 스크립트 구축
    script = f"""
// ═══════════════════════════════════════════════
// DATA (Auto-generated: {data_date})
// ═══════════════════════════════════════════════
var MYPORT = {j(MYPORT)};
var TC = {j(TC)};
var NAMES = {j(NAMES)};
var PORT_SEC = {j(PORT_SEC)};
var ALLOC_W = {j(ALLOC_W)};
var SMAP = {j(smap)};
var SETFS = [];
var SP5 = [];
var SP5_INFO = {j(sp5_info)};
for(var si=0;si<SMAP.length;si++){{
  SETFS.push(SMAP[si].etf);
  for(var sj=0;sj<SMAP[si].st.length;sj++){{
    var tk=SMAP[si].st[sj];
    if(MYPORT.indexOf(tk)<0)SP5.push(tk);
  }}
}}
var EXTRA_SP5={j(remaining)};
for(var ei=0;ei<EXTRA_SP5.length;ei++){{if(SP5.indexOf(EXTRA_SP5[ei])<0)SP5.push(EXTRA_SP5[ei])}}
var FN = {{}};
var FC = {j(fc)};
var EV = {j(ev)};
var BETA_MAP = {j(betas or {})};

// ═══════════════════════════════════════════════
// STATE
// ═══════════════════════════════════════════════
var sel = "NVDA", tab = "signal", ca = 5;
var ut = {j({t: 0 for t in MYPORT})};
var charts = {{}};
var TD = {j(td)};
var isFetching = false;
var API_DELAY = 0;
var DAILY_LIMIT = 999;

{kept_js}

// ═══════════════════════════════════════════════
// INIT (Offline Mode)
// ═══════════════════════════════════════════════
function init(){{
  try{{render()}}catch(e){{
    document.body.innerHTML='<div style="padding:20px;color:#e2e8f0;font-family:-apple-system,system-ui,sans-serif"><h2 style="color:#ff4757">⚠ 렌더링 오류</h2><p style="font-size:13px">'+e.message+'</p><p style="color:#64748b;font-size:11px;margin-top:10px">Safari 브라우저에서 열어주세요.</p></div>';
  }}
}}
if(document.readyState==="loading"){{document.addEventListener("DOMContentLoaded",init)}}else{{init()}}
"""

    # 목표가 초기값 설정 (FC 기반)
    ut_init = {}
    for item in fc:
        if item["lo"] > 0 and item["hi"] > 0:
            ut_init[item["t"]] = round((item["lo"] + item["hi"]) / 2)
        else:
            ut_init[item["t"]] = 0
    script = script.replace(
        f'var ut = {j({t: 0 for t in MYPORT})};',
        f'var ut = {j(ut_init)};'
    )

    ticker_count = len(td)

    # HTML 본문
    body = f"""</head>
<body>
<div class="w">
  <div class="topbar">
    <div class="row1">
      <div class="status" id="statusArea"><span class="dot dot-live"></span>오프라인 데이터 &middot; {data_date}</div>
      <div class="quota" id="quotaArea">{ticker_count}종목 로드</div>
    </div>
  </div>
  <div class="prog hide" id="progArea">
    <div class="bar"><div class="fill" id="progFill" style="width:100%"></div></div>
    <div class="info"><span id="progText">완료</span><span id="progPct">100%</span></div>
  </div>
  <div id="errArea"></div>
  <div class="hd">
    <div><h1>PORTFOLIO COMMAND CENTER</h1><div class="sub" id="hSub">7 Holdings &middot; S&amp;P 500 &middot; {data_date}</div></div>
    <div class="rv"><div class="lb">PORTFOLIO 1M AVG</div><div class="vl" id="hAvg">&mdash;</div></div>
  </div>
  <div class="st" id="strip"></div>
  <div class="tb" id="tabs"></div>
  <div id="pSignal"></div>
  <div id="pTarget" class="hide"></div>
  <div id="pPort" class="hide"></div>
  <div id="pRec" class="hide"></div>
  <div id="pSector" class="hide"></div>
  <div class="ft">&olarr; 본 대시보드는 투자 조언이 아닙니다. 과거 데이터 기반 분석이며 미래 수익을 보장하지 않습니다.<br>S&amp;P 500 {ticker_count}종목 &middot; 자동 생성: {data_date}</div>
</div>
<script>{script}
</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js" async onload="try{{if(typeof render==='function')render()}}catch(e){{}}"></script>
</body>
</html>
"""

    return css + body


# ── 모바일 최적화 CSS ──────────────────────────────
MOBILE_CSS = """
:root{--bg:#0b0f19;--card:#111827;--card2:#0f172a;--bdr:#1e293b;--txt:#e2e8f0;--dim:#64748b;--dim2:#475569;--g:#00d4aa;--r:#ff4757;--a:#fbbf24;--b:#3b82f6;--p:#a855f7;--safe-b:env(safe-area-inset-bottom)}
@supports(padding:max(0px)){:root{--safe-b:max(0px,env(safe-area-inset-bottom))}}
*{margin:0;padding:0;box-sizing:border-box}
body{background:var(--bg);color:var(--txt);font-family:-apple-system,BlinkMacSystemFont,'Inter',system-ui,sans-serif;font-size:14px;-webkit-font-smoothing:antialiased;padding-bottom:constant(safe-area-inset-bottom);padding-bottom:env(safe-area-inset-bottom);padding-bottom:calc(68px + constant(safe-area-inset-bottom));padding-bottom:calc(68px + env(safe-area-inset-bottom));overflow-x:hidden;-webkit-text-size-adjust:100%;-webkit-overflow-scrolling:touch}
.w{max-width:100%;margin:0 auto;padding:10px 12px}
.hd{display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--bdr);margin-bottom:10px;-webkit-flex-wrap:wrap;flex-wrap:wrap;gap:8px}
.hd h1{font-size:14px;font-weight:900;letter-spacing:1.5px;background:linear-gradient(135deg,var(--g),var(--b));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hd .sub{font-size:9px;color:var(--dim);margin-top:2px}
.hd .rv{text-align:right}.hd .rv .lb{font-size:9px;color:var(--dim)}.hd .rv .vl{font-size:18px;font-weight:900}
.cg{color:var(--g)}.cr{color:var(--r)}.ca{color:var(--a)}.cb{color:var(--b)}.cp{color:var(--p)}
.topbar{position:-webkit-sticky;position:sticky;top:0;z-index:100;padding:10px 12px;background:rgba(17,24,39,.97);-webkit-backdrop-filter:blur(12px);backdrop-filter:blur(12px);border:1px solid var(--bdr);border-radius:0 0 12px 12px;margin:-10px -12px 10px;padding-top:calc(10px + constant(safe-area-inset-top));padding-top:calc(10px + env(safe-area-inset-top))}
.topbar .row1{display:flex;justify-content:space-between;align-items:center;-webkit-flex-wrap:wrap;flex-wrap:wrap;gap:6px}
.topbar .status{font-size:10px;color:var(--dim)}
.topbar .acts{display:flex;gap:6px;align-items:center;-webkit-flex-wrap:wrap;flex-wrap:wrap;margin-top:8px}
.topbar input{width:100%;padding:10px 12px;border-radius:8px;font-size:16px;font-family:inherit;background:var(--card2);color:var(--a);border:1px solid var(--bdr);outline:none;-webkit-appearance:none;appearance:none}
.topbar input:focus{border-color:var(--a)}.topbar input::placeholder{color:var(--dim2)}
.dot{display:inline-block;width:7px;height:7px;border-radius:50%;margin-right:4px;vertical-align:middle}
.dot-live{background:var(--g);box-shadow:0 0 8px var(--g)}.dot-off{background:var(--r)}.dot-load{background:var(--a);-webkit-animation:pulse 1s infinite;animation:pulse 1s infinite}
@-webkit-keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
.btn{padding:10px 16px;border-radius:8px;border:none;font-size:11px;font-weight:700;font-family:inherit;cursor:pointer;-webkit-tap-highlight-color:transparent;white-space:nowrap;min-height:44px;-webkit-appearance:none;appearance:none}
.btn-g{background:var(--g);color:#000}.btn-b{background:var(--b);color:#fff}.btn-a{background:var(--a);color:#000}.btn-p{background:var(--p);color:#fff}
.btn:disabled{opacity:.4;cursor:not-allowed}
.prog{position:-webkit-sticky;position:sticky;top:52px;z-index:99;background:var(--bg);padding:6px 0;margin-bottom:4px}
.prog .bar{height:5px;background:var(--bdr);border-radius:3px;overflow:hidden}
.prog .fill{height:100%;background:linear-gradient(90deg,var(--g),var(--b));border-radius:3px;-webkit-transition:width .3s;transition:width .3s}
.prog .info{font-size:9px;color:var(--dim);margin-top:4px;display:-webkit-flex;display:flex;-webkit-justify-content:space-between;justify-content:space-between}
.errlog{margin-bottom:8px;padding:10px 12px;background:rgba(255,71,87,.08);border:1px solid rgba(255,71,87,.2);border-radius:10px;font-size:11px;color:var(--r);line-height:1.5}
.st{display:-webkit-flex;display:flex;gap:6px;margin-bottom:10px;overflow-x:auto;padding:4px 0 8px;-webkit-overflow-scrolling:touch;scroll-snap-type:x mandatory;-ms-overflow-style:none;scrollbar-width:none}
.st::-webkit-scrollbar{height:0;display:none}
.st button{-webkit-flex-shrink:0;flex-shrink:0;padding:10px 14px;border-radius:12px;border:1px solid var(--bdr);background:transparent;color:var(--dim);font-family:inherit;font-size:11px;font-weight:700;cursor:pointer;-webkit-tap-highlight-color:transparent;min-height:50px;scroll-snap-align:start;-webkit-appearance:none;appearance:none}
.st button.on{border-color:var(--g);background:rgba(0,212,170,.08);color:#fff}
.st .dum{border-style:dashed;opacity:.45;font-size:9px}
.tb{position:fixed;bottom:0;left:0;right:0;z-index:1000;display:-webkit-flex;display:flex;-webkit-justify-content:space-around;justify-content:space-around;-webkit-align-items:center;align-items:center;gap:0;background:rgba(17,24,39,.97);-webkit-backdrop-filter:saturate(180%) blur(20px);backdrop-filter:saturate(180%) blur(20px);border-top:1px solid var(--bdr);padding:4px 2px;padding-bottom:constant(safe-area-inset-bottom);padding-bottom:env(safe-area-inset-bottom);padding-bottom:calc(4px + constant(safe-area-inset-bottom));padding-bottom:calc(4px + env(safe-area-inset-bottom))}
.tb::-webkit-scrollbar{height:0;display:none}
.tb button{-webkit-flex:1;flex:1;padding:8px 2px;border-radius:10px;border:none;font-size:9px;font-weight:700;font-family:inherit;cursor:pointer;background:transparent;color:var(--dim);-webkit-tap-highlight-color:transparent;white-space:nowrap;min-height:44px;text-align:center;-webkit-transition:background .15s;transition:background .15s;-webkit-appearance:none;appearance:none}
.tb button.on{background:var(--b);color:#fff}
.cd{background:var(--card);border:1px solid var(--bdr);border-radius:14px;padding:14px;margin-bottom:10px}
.cd.warn{border-color:rgba(255,71,87,.2)}
.ct{font-size:12px;font-weight:700;margin-bottom:10px}
.g2{display:-ms-grid;display:grid;grid-template-columns:1fr;gap:12px}
.g3{display:-ms-grid;display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px}
.ag{display:-ms-grid;display:grid;grid-template-columns:1fr 1fr;gap:8px}
.bx{padding:10px;background:var(--card2);border-radius:8px}.bx .lb{color:var(--dim);font-size:9px}.bx .vl{font-weight:700;font-size:13px;margin-top:3px}
.tbl{width:100%;border-collapse:collapse;font-size:10px}
.tbl th{padding:8px 3px;text-align:center;color:var(--dim);font-weight:600;font-size:8px;border-bottom:1px solid var(--bdr);white-space:nowrap;position:-webkit-sticky;position:sticky;top:0;background:var(--card);z-index:2}
.tbl td{padding:8px 3px;text-align:center;border-bottom:1px solid rgba(30,41,59,.15)}
.tbl tr.ck{cursor:pointer}.tbl tr.ck:active{background:rgba(30,41,59,.4)}
.pl{padding:4px 8px;border-radius:6px;font-size:8px;font-weight:800;display:inline-block}
.badge{font-size:8px;padding:2px 6px;border-radius:4px;font-weight:700}
.rank-num{display:-webkit-inline-flex;display:inline-flex;width:22px;height:22px;border-radius:50%;-webkit-align-items:center;align-items:center;-webkit-justify-content:center;justify-content:center;font-size:10px;font-weight:900;color:#000}
.sb{margin-bottom:6px}.sb .tp{display:-webkit-flex;display:flex;-webkit-justify-content:space-between;justify-content:space-between;font-size:10px;color:var(--dim);margin-bottom:3px}
.sb .br{height:6px;background:var(--bdr);border-radius:4px;position:relative;overflow:hidden}
.sb .fl{height:100%;border-radius:4px;position:absolute}.sb .cn{position:absolute;left:50%;top:0;width:1px;height:100%;background:var(--dim2)}
.sg{margin-top:12px;padding:16px;background:var(--card2);border-radius:14px;text-align:center}
.sg .nm{font-size:30px;font-weight:900;letter-spacing:3px}.sg .ds{font-size:10px;color:var(--dim);margin-top:5px}
.ai{padding:12px;border-radius:12px;background:var(--card2)}.ai .ab{height:5px;background:var(--bdr);border-radius:3px;margin-top:6px}.ai .af{height:100%;border-radius:3px}
.el{max-height:280px;overflow-y:auto;-webkit-overflow-scrolling:touch}
.er{display:-webkit-flex;display:flex;-webkit-align-items:flex-start;align-items:flex-start;gap:6px;padding:6px 0;border-bottom:1px solid rgba(30,41,59,.1);font-size:11px}
.tr{margin-bottom:14px;padding-bottom:14px;border-bottom:1px solid var(--bdr)}
.tgt-wrap{display:-webkit-flex;display:flex;-webkit-align-items:center;align-items:center;gap:10px;background:var(--card2);padding:10px 14px;border-radius:10px;border:1px solid var(--bdr);margin-top:6px;-webkit-flex-wrap:wrap;flex-wrap:wrap}
.tgt-wrap:focus-within{border-color:var(--a);box-shadow:0 0 12px rgba(251,191,36,.2)}
.tgt-input{width:90px;padding:8px;border-radius:6px;font-size:16px;font-weight:700;font-family:inherit;background:transparent;color:var(--a);border:none;text-align:right;outline:none;-webkit-appearance:none;appearance:none}
.tgt-ret{font-size:18px;font-weight:900;min-width:80px;text-align:right}
.tgt-bar{height:6px;border-radius:3px;margin-top:6px;background:var(--bdr);overflow:hidden;position:relative}
.tgt-bar .tgt-fill{position:absolute;height:100%;border-radius:3px;-webkit-transition:width .2s,left .2s;transition:width .2s,left .2s}
.ti{width:60px;padding:8px;border-radius:6px;font-size:16px;font-family:inherit;background:var(--card2);color:var(--a);border:1px solid rgba(251,191,36,.3);text-align:right;outline:none;min-height:36px;-webkit-appearance:none;appearance:none}
.ti:focus{border-color:var(--a)}
.cc{position:relative;width:100%;height:200px}.cc canvas{width:100%!important;height:100%!important}
.lg{display:-webkit-flex;display:flex;gap:10px;font-size:9px;color:var(--dim);-webkit-justify-content:center;justify-content:center;margin-top:5px}
.lg i{display:inline-block;width:12px;height:3px;border-radius:2px;margin-right:3px;vertical-align:middle}
input[type=range]{width:100%;-webkit-appearance:none;appearance:none;height:28px;background:transparent}
input[type=range]::-webkit-slider-runnable-track{height:6px;background:var(--bdr);border-radius:3px}
input[type=range]::-webkit-slider-thumb{-webkit-appearance:none;width:22px;height:22px;border-radius:50%;background:var(--a);margin-top:-8px}
.sec-card{background:var(--card2);border-radius:12px;padding:14px;border-left:4px solid var(--dim);margin-bottom:8px}
.sec-bar{height:8px;background:var(--bdr);border-radius:4px;overflow:hidden;margin-top:6px}
.sec-fill{height:100%;border-radius:4px;-webkit-transition:width .3s;transition:width .3s}
.metric-bar{display:-webkit-flex;display:flex;-webkit-align-items:center;align-items:center;gap:6px;margin:5px 0}
.metric-bar .mb-label{font-size:9px;color:var(--dim);width:32px;-webkit-flex-shrink:0;flex-shrink:0;text-align:right}
.metric-bar .mb-track{-webkit-flex:1;flex:1;height:10px;background:var(--bdr);border-radius:5px;overflow:hidden}
.metric-bar .mb-fill{height:100%;border-radius:5px;-webkit-transition:width .3s;transition:width .3s}
.metric-bar .mb-val{font-size:10px;font-weight:700;width:44px;text-align:right;-webkit-flex-shrink:0;flex-shrink:0}
.hm-cell{padding:14px 6px;border-radius:10px;text-align:center;cursor:default;-webkit-transform:translateZ(0);transform:translateZ(0)}
.hm-cell:active{-webkit-transform:scale(0.95);transform:scale(0.95)}
.hm-cell .hm-nm{font-size:10px;font-weight:700;color:#fff;text-shadow:0 1px 4px rgba(0,0,0,.7)}
.hm-cell .hm-val{font-size:16px;font-weight:900;color:#fff;text-shadow:0 1px 4px rgba(0,0,0,.7);margin-top:2px}
.hm-cell .hm-etf{font-size:8px;color:rgba(255,255,255,.8);margin-top:2px}
.beta-badge{font-size:9px;padding:3px 7px;border-radius:4px;font-weight:800;display:inline-block;margin-left:4px;letter-spacing:.3px}
/* ═══ rRec Signal Grid UI ═══ */
.rec-summary{display:grid;grid-template-columns:repeat(5,1fr);gap:6px;margin-bottom:6px}
.rec-sum-item{text-align:center;padding:8px 4px;background:var(--card2);border-radius:6px}
.rec-sum-n{font-size:20px;font-weight:900;color:#fff}
.rec-sum-lb{font-size:8px;color:var(--dim);margin-top:2px}
.rec-section{margin-bottom:10px}
.rec-sec-title{display:flex;justify-content:space-between;align-items:center;font-size:12px;font-weight:700;margin-bottom:8px;padding:6px 0}
.rec-sec-cnt{font-size:9px;color:var(--dim);font-weight:400}
.rec-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:6px}
@media(min-width:900px){.rec-grid{grid-template-columns:repeat(3,1fr)}}
@media(max-width:480px){.rec-grid{grid-template-columns:1fr}}
.rec-card{background:var(--card);border:1px solid var(--bdr);border-radius:8px;padding:10px;transition:border-color .2s}
.rec-card:hover{border-color:var(--b)}
.rec-card-head{display:flex;align-items:flex-start;gap:8px;margin-bottom:6px}
.rec-rank{width:22px;height:22px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-size:10px;font-weight:900;flex-shrink:0}
.rec-card-body{display:grid;grid-template-columns:repeat(4,1fr);gap:4px;padding:6px 0;border-top:1px solid var(--bdr);border-bottom:1px solid var(--bdr)}
.rec-stat{text-align:center}
.rec-stat-lb{display:block;font-size:7px;color:var(--dim);margin-bottom:1px}
.rec-stat span:last-child{font-size:10px;font-weight:700}
.quota{font-size:9px;color:var(--dim);padding:3px 8px;background:var(--card2);border-radius:6px}
.ft{text-align:center;padding:12px 0;padding-bottom:calc(80px + constant(safe-area-inset-bottom));padding-bottom:calc(80px + env(safe-area-inset-bottom));font-size:8px;color:var(--dim);border-top:1px solid var(--bdr);margin-top:12px}
.hide{display:none!important}
"""


def make_mobile_html(desktop_html):
    """데스크톱 HTML의 CSS만 교체하여 모바일 최적화 버전 생성."""
    html = desktop_html.replace(
        "<title>Portfolio Command Center v4</title>",
        "<title>Portfolio Mobile v4</title>",
    )
    style_start = html.find("<style>") + len("<style>")
    style_end = html.find("</style>")
    return html[:style_start] + MOBILE_CSS + html[style_end:]


def main():
    log.info("Dashboard Generator START")

    # 데이터 로드
    if not EXCEL_FILE.exists():
        log.error(f"Excel not found: {EXCEL_FILE}")
        sys.exit(1)
    if not CHART_FILE.exists():
        log.error(f"Chart data not found: {CHART_FILE}")
        sys.exit(1)
    if not TEMPLATE_FILE.exists():
        log.error(f"Template not found: {TEMPLATE_FILE}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_FILE, sheet_name="Summary")
    meta = pd.read_excel(EXCEL_FILE, sheet_name="Meta")
    with open(CHART_FILE, encoding="utf-8") as f:
        chart = json.load(f)

    data_date = str(meta.iloc[0]["Generated"])[:10]
    log.info(f"Data date: {data_date}, Tickers: {len(df)}, Chart tickers: {len(chart)}")

    # 애널리스트 목표가 로드
    analyst = {}
    if ANALYST_FILE.exists():
        with open(ANALYST_FILE, encoding="utf-8") as f:
            analyst = json.load(f)
        log.info(f"Analyst targets loaded: {len(analyst)} tickers")
    else:
        log.warning(f"Analyst file not found: {ANALYST_FILE} (using fallback)")

    # 데이터 구축
    smap = build_smap(df)
    sp5_info = build_sp5_info(df)
    td = build_td(df, chart)
    fc = build_fc(df, analyst)
    ev = build_ev(df)
    betas = compute_betas(chart)
    all_sp5_tickers = df["Ticker"].tolist()

    log.info(f"TD entries: {len(td)}, SMAP sectors: {len(smap)}, Events: {len(ev)}, Betas: {len(betas)}")

    # ── JSON 파일 분리 저장 (HTML 인라인 주입 대체) ──
    ut_init = {}
    for item in fc:
        if item.get("lo", 0) > 0 and item.get("hi", 0) > 0:
            ut_init[item["t"]] = round((item["lo"] + item["hi"]) / 2)
        else:
            ut_init[item["t"]] = 0
    write_data_js(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas, ut_init)
    write_tickers_js(df)

    # ── 환율 조회 ──
    fx_rate = fetch_fx_rate()

    # ── 통합 대시보드 생성 (Desktop: Dashboard B + SA) ──
    if MERGED_TEMPLATE.exists():
        merged_html = generate_merged_html(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas, df=df, fx_rate=fx_rate, analyst=analyst)
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(merged_html)
        size_kb = OUTPUT_FILE.stat().st_size / 1024
        log.info(f"Saved MERGED {OUTPUT_FILE} ({size_kb:.0f} KB)")
    else:
        log.warning(f"Merged template not found: {MERGED_TEMPLATE} → fallback to standalone")
        html = generate_html(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas)
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            f.write(html)
        size_kb = OUTPUT_FILE.stat().st_size / 1024
        log.info(f"Saved standalone {OUTPUT_FILE} ({size_kb:.0f} KB)")

    # ── 모바일 버전 생성 ──
    if MERGED_TEMPLATE.exists() and 'merged_html' in dir():
        mobile_html = make_mobile_html(merged_html)
    else:
        standalone_html = generate_html(td, sp5_info, smap, fc, ev, all_sp5_tickers, data_date, betas)
        mobile_html = make_mobile_html(standalone_html)
    with open(MOBILE_OUTPUT, "w", encoding="utf-8") as f:
        f.write(mobile_html)
    mobile_kb = MOBILE_OUTPUT.stat().st_size / 1024
    log.info(f"Saved {MOBILE_OUTPUT} ({mobile_kb:.0f} KB)")

    log.info("Dashboard Generator DONE")


if __name__ == "__main__":
    main()
