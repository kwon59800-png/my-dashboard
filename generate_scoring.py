#!/usr/bin/env python3
"""
generate_scoring.py  v3
───────────────────────
Russell 3000 + 포트폴리오 + 주요 ADR 종목을 대상으로
주가·펀더멘털 데이터를 yfinance로 수집하고
100점 만점 투자 스코어링을 계산하여 scoring.js로 출력한다.

출력 형식: const SCORING_DATA = { ... };  (file:// CORS 회피)

카테고리 배점:
  기술적(Tech)      30점  ─ 이평선 정배열 15 + RSI 15
  펀더멘털(Fund)    30점  ─ EPS Growth 15 + Forward P/E 15
  수급(Supply)      20점  ─ 거래량 급증 비율
  심리(Sentiment)   20점  ─ 목표가 상승 여력

v3: Russell 3000 확장, 배치 딜레이 + 재시도 방어, 주요 ADR 강제 포함
"""

import json
import logging
import math
import time
import urllib.request
import csv
import io
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import pandas as pd
import yfinance as yf

# ── 설정 ──────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = SCRIPT_DIR / "scoring.js"

# ── 1) 포트폴리오 보유 종목 (항상 최우선 포함) ──
PORTFOLIO_TICKERS = [
    "NVDA", "TSM", "MU", "CEG", "SOXL",
    "IONQ", "CRCL", "AAPL", "MSFT", "TSLA",
]

# ── 2) 주요 ADR / Russell 3000에 빠질 수 있는 종목 강제 포함 ──
FORCE_INCLUDE = [
    # 대형 ADR
    "TSM", "ASML", "ARM", "NVO", "SAP", "SHOP", "MELI", "SE", "GRAB",
    "BABA", "PDD", "JD", "BIDU", "NIO", "XPEV", "LI",
    # 인기 소형주/양자컴퓨팅/크립토
    "RGTI", "QBTS", "QUBT", "IREN", "MARA", "RIOT", "COIN", "HOOD",
    "SMCI", "CRCL", "SOXL",
    # 인기 ETF
    "QQQ", "SPY", "SOXX", "XLK", "XLE", "XLF", "XLV", "ARKK", "ARKG",
    # 주요 지수 (헤더 인덱스 바용)
    "^GSPC", "^IXIC",
]

# 인덱스 티커 목록 (fetch_ticker_data 특수 처리)
INDEX_TICKERS = {"^GSPC", "^IXIC"}

# ── 주요 지수 구성종목 (sector/index badge용) ──
_IDX_CACHE: dict[str, set[str]] = {}

def _load_index_constituents() -> None:
    """S&P500 / Nasdaq-100 구성종목을 위키에서 1회 로딩하여 캐시."""
    if _IDX_CACHE:
        return
    _IDX_CACHE["SP500"] = set()
    _IDX_CACHE["NDX100"] = set()
    try:
        html = _url_read("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies")
        tables = pd.read_html(io.StringIO(html), attrs={"id": "constituents"})
        if tables:
            for t in tables[0]["Symbol"].str.replace(".", "-", regex=False):
                _IDX_CACHE["SP500"].add(t.strip().upper())
        log.info(f"S&P 500 구성종목 로드: {len(_IDX_CACHE['SP500'])}개")
    except Exception as e:
        log.warning(f"S&P 500 목록 로드 실패: {e}")
    try:
        html = _url_read("https://en.wikipedia.org/wiki/Nasdaq-100")
        tables = pd.read_html(io.StringIO(html))
        for tbl in tables:
            if "Ticker" in tbl.columns:
                for t in tbl["Ticker"]:
                    _IDX_CACHE["NDX100"].add(str(t).strip().upper())
                break
            elif "Symbol" in tbl.columns:
                for t in tbl["Symbol"]:
                    _IDX_CACHE["NDX100"].add(str(t).strip().upper())
                break
        log.info(f"Nasdaq-100 구성종목 로드: {len(_IDX_CACHE['NDX100'])}개")
    except Exception as e:
        log.warning(f"Nasdaq-100 목록 로드 실패: {e}")

def _get_idx_group(ticker: str) -> str:
    """티커의 소속 지수 뱃지 문자열 반환."""
    parts = []
    if ticker in _IDX_CACHE.get("NDX100", set()):
        parts.append("NDX")
    if ticker in _IDX_CACHE.get("SP500", set()):
        parts.append("S&P")
    return ",".join(parts) if parts else ""

# ── 섹터 한글 매핑 ──
SECTOR_KR = {
    "Technology": "기술", "Healthcare": "헬스케어", "Financial Services": "금융",
    "Consumer Cyclical": "경기소비재", "Communication Services": "커뮤니케이션",
    "Industrials": "산업재", "Consumer Defensive": "필수소비재", "Energy": "에너지",
    "Basic Materials": "소재", "Real Estate": "부동산", "Utilities": "유틸리티",
}

# ── 병렬 처리 설정 ──
MAX_TICKERS = 1500       # 수집 대상 최대 종목 수 (시총 상위)
MAX_WORKERS = 8          # 동시 스레드 수
BATCH_SIZE = 40          # N개 처리마다 쉼
BATCH_SLEEP = 2.0        # 배치 간 대기 (초)
RETRY_COUNT = 1          # 실패 시 재시도 횟수
RETRY_SLEEP = 3.0        # 재시도 대기 (초)
FAIL_STREAK_ABORT = 3    # 연속 N배치 전원 실패 시 조기 종료

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════
# 0. 러셀 3000 티커 수집
# ══════════════════════════════════════════════════════════════

def _url_read(url: str, ua: str = "Mozilla/5.0 ScoringBot/3.0") -> str:
    """User-Agent 헤더 포함 URL 읽기."""
    req = urllib.request.Request(url, headers={"User-Agent": ua})
    with urllib.request.urlopen(req, timeout=20) as resp:
        return resp.read().decode("utf-8")


def _parse_market_cap(val_str: str) -> float:
    """Nasdaq screener 시총 문자열 → 숫자 (e.g. '$1.23B' → 1_230_000_000)."""
    if not val_str or val_str == "N/A":
        return 0
    s = val_str.replace("$", "").replace(",", "").strip()
    mult = 1
    if s.endswith("T"):
        mult = 1e12; s = s[:-1]
    elif s.endswith("B"):
        mult = 1e9; s = s[:-1]
    elif s.endswith("M"):
        mult = 1e6; s = s[:-1]
    elif s.endswith("K"):
        mult = 1e3; s = s[:-1]
    try:
        return float(s) * mult
    except ValueError:
        return 0


def fetch_russell3000_tickers() -> list[str]:
    """
    미국 상장 주식 중 시총 상위 ~3000개를 가져온다.
    Nasdaq Screener API에서 전체 수집 후 시총 기준 필터링.
    """
    all_stocks = []   # [(symbol, market_cap), ...]

    # ── Nasdaq Screener API (모든 거래소) ──
    try:
        log.info("Nasdaq Screener API에서 전체 미국 주식 수집 중...")
        for exchange in ["nasdaq", "nyse", "amex"]:
            url = (
                f"https://api.nasdaq.com/api/screener/stocks"
                f"?tableType=traded&exchange={exchange}&download=true"
            )
            req = urllib.request.Request(url, headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "application/json",
            })
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))

            rows = data.get("data", {}).get("rows", [])
            for row in rows:
                sym = row.get("symbol", "").strip().upper()
                if sym and re.match(r'^[A-Z]{1,5}$', sym):
                    mcap = _parse_market_cap(row.get("marketCap", ""))
                    all_stocks.append((sym, mcap))
            log.info(f"  {exchange.upper()}: {len(rows)}행 수집")

        log.info(f"Nasdaq Screener 총 수집: {len(all_stocks)}개")
    except Exception as e:
        log.warning(f"Nasdaq Screener 실패: {e}")

    # ── 위키피디아 S&P 500 보완 ──
    if len(all_stocks) < 500:
        try:
            html = _url_read("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies")
            tables = pd.read_html(io.StringIO(html), attrs={"id": "constituents"})
            if tables:
                sp5 = tables[0]["Symbol"].str.replace(".", "-", regex=False).tolist()
                existing = {s[0] for s in all_stocks}
                for t in sp5:
                    t = t.strip().upper()
                    if t and t not in existing:
                        all_stocks.append((t, 1e9))  # S&P500은 시총 큰 편
                        existing.add(t)
                log.info(f"S&P 500 보완 완료")
        except Exception as e:
            log.warning(f"S&P 500 보완 실패: {e}")

    # ── 시총 상위 MAX_TICKERS개로 제한 ──
    # 시총 정보가 있는 종목은 시총 내림차순, 없는 종목은 뒤에 배치
    all_stocks.sort(key=lambda x: x[1], reverse=True)
    seen = set()
    result = []
    for sym, mcap in all_stocks:
        if sym not in seen:
            result.append(sym)
            seen.add(sym)
        if len(result) >= MAX_TICKERS:
            break

    log.info(f"시총 기준 상위 {len(result)}개 선별 (전체 {len(all_stocks)}개 중)")
    return result


def build_ticker_universe() -> list[str]:
    """포트폴리오 + FORCE_INCLUDE + Russell3000을 합쳐 중복 없는 유니버스 구성."""
    seen = set()
    universe = []

    # 포트폴리오 최우선
    for tk in PORTFOLIO_TICKERS:
        t = tk.upper().strip()
        if t and t not in seen:
            universe.append(t)
            seen.add(t)

    # 강제 포함 종목 (ADR 등)
    for tk in FORCE_INCLUDE:
        t = tk.upper().strip()
        if t and t not in seen:
            universe.append(t)
            seen.add(t)

    # Russell 3000 (Nasdaq Screener)
    r3k = fetch_russell3000_tickers()
    for tk in r3k:
        if tk and tk not in seen:
            universe.append(tk)
            seen.add(tk)

    log.info(f"═══ 티커 유니버스: 총 {len(universe)}개 ═══")
    return universe


# ══════════════════════════════════════════════════════════════
# 1. 데이터 수집 함수
# ══════════════════════════════════════════════════════════════

def fetch_ticker_data(ticker: str) -> dict | None:
    """yfinance에서 단일 종목의 주가·펀더멘털·컨센서스 데이터를 수집."""
    try:
        tk = yf.Ticker(ticker)

        # ── 인덱스 티커 전용 (^GSPC, ^IXIC) ──
        if ticker in INDEX_TICKERS:
            hist = tk.history(period="5d")
            if hist.empty or len(hist) < 2:
                return None
            close = hist["Close"]
            cp = float(close.iloc[-1])
            pc = float(close.iloc[-2])
            chg = round((cp - pc) / pc * 100, 2) if pc > 0 else 0
            info = tk.info or {}
            return {
                "price": round(cp, 2), "prev_close": round(pc, 2),
                "change": chg, "change_abs": round(cp - pc, 2),
                "ma20": None, "ma50": None, "rsi": None, "vol_ratio": None,
                "fwd_pe": None, "eps_growth": None, "trailing_eps": None,
                "fwd_eps": None, "target_mean": None, "upside": None,
                "name": info.get("shortName", ticker),
                "sector": "", "idx_grp": "",
            }

        # ── 주가 히스토리 (6개월) ──
        hist = tk.history(period="6mo")
        if hist.empty or len(hist) < 20:
            return None

        close = hist["Close"]
        volume = hist["Volume"]
        current_price = float(close.iloc[-1])
        if current_price <= 0:
            return None

        # 전일 대비 등락률 (close[-1] vs close[-2])
        prev_close = float(close.iloc[-2]) if len(close) >= 2 else 0
        change_pct = round((current_price - prev_close) / prev_close * 100, 2) if prev_close > 0 else 0
        change_abs = round(current_price - prev_close, 2) if prev_close > 0 else 0

        # 이동평균선
        ma20 = float(close.rolling(20).mean().iloc[-1])
        ma50 = float(close.rolling(50).mean().iloc[-1]) if len(close) >= 50 else None

        # RSI (14일)
        rsi = _calc_rsi(close, 14)

        # 거래량 비율 (당일 / 최근 5일 평균)
        vol_today = float(volume.iloc[-1])
        vol_avg5 = float(volume.iloc[-6:-1].mean()) if len(volume) >= 6 else float(volume.mean())
        vol_ratio = round(vol_today / vol_avg5, 2) if vol_avg5 > 0 else 0

        # ── 펀더멘털 ──
        info = tk.info or {}
        fwd_pe = info.get("forwardPE")
        trailing_eps = info.get("trailingEps")
        fwd_eps = info.get("forwardEps")
        eps_growth = None
        if trailing_eps and fwd_eps and trailing_eps > 0:
            eps_growth = round((fwd_eps - trailing_eps) / abs(trailing_eps) * 100, 1)

        # ── 목표가 컨센서스 ──
        target_mean = info.get("targetMeanPrice")
        upside = None
        if target_mean and current_price > 0:
            upside = round((target_mean - current_price) / current_price * 100, 1)

        # ── 섹터 & 지수 소속 ──
        raw_sector = info.get("sector", "")
        sector_kr = SECTOR_KR.get(raw_sector, raw_sector)
        idx_grp = _get_idx_group(ticker)

        return {
            "price": round(current_price, 2),
            "prev_close": prev_close,
            "change": change_pct,
            "change_abs": change_abs,
            "ma20": round(ma20, 2),
            "ma50": round(ma50, 2) if ma50 is not None else None,
            "rsi": rsi,
            "vol_ratio": vol_ratio,
            "fwd_pe": round(fwd_pe, 1) if fwd_pe else None,
            "eps_growth": eps_growth,
            "trailing_eps": round(trailing_eps, 2) if trailing_eps else None,
            "fwd_eps": round(fwd_eps, 2) if fwd_eps else None,
            "target_mean": round(target_mean, 2) if target_mean else None,
            "upside": upside,
            "name": info.get("shortName", ticker),
            "sector": sector_kr,
            "idx_grp": idx_grp,
        }

    except Exception:
        return None


def fetch_with_retry(ticker: str) -> tuple[str, dict | None]:
    """재시도 포함 단일 종목 수집."""
    for attempt in range(RETRY_COUNT + 1):
        data = fetch_ticker_data(ticker)
        if data is not None:
            return ticker, data
        if attempt < RETRY_COUNT:
            time.sleep(RETRY_SLEEP)
    return ticker, None


def _calc_rsi(series: pd.Series, period: int = 14) -> float | None:
    """RSI — Wilder's smoothing."""
    try:
        delta = series.diff()
        gain = delta.where(delta > 0, 0.0)
        loss = (-delta).where(delta < 0, 0.0)

        avg_gain = gain.rolling(window=period, min_periods=period).mean()
        avg_loss = loss.rolling(window=period, min_periods=period).mean()

        for i in range(period, len(avg_gain)):
            if pd.notna(avg_gain.iloc[i - 1]):
                avg_gain.iloc[i] = (avg_gain.iloc[i - 1] * (period - 1) + gain.iloc[i]) / period
                avg_loss.iloc[i] = (avg_loss.iloc[i - 1] * (period - 1) + loss.iloc[i]) / period

        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))
        val = rsi.iloc[-1]
        return round(float(val), 1) if pd.notna(val) and not math.isinf(val) else None
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════
# 2. 스코어링 함수 (100점 만점)
# ══════════════════════════════════════════════════════════════

def score_tech(data: dict) -> tuple[int, dict]:
    ma_score = 0
    price, ma20, ma50 = data["price"], data.get("ma20"), data.get("ma50")
    if ma20 is not None:
        if ma50 is not None and price > ma20 > ma50:
            ma_score = 15
        elif price > ma20:
            ma_score = 10
        elif ma50 is not None and price > ma50:
            ma_score = 5

    rsi_score = 0
    rsi = data["rsi"]
    if rsi is not None:
        if 40 <= rsi <= 60:
            rsi_score = 15
        elif rsi < 40:
            rsi_score = 10
        elif 60 < rsi <= 70:
            rsi_score = 5

    return ma_score + rsi_score, {"ma": ma_score, "rsi": rsi_score}


def score_fund(data: dict) -> tuple[int, dict]:
    eps_score = 0
    eg = data["eps_growth"]
    if eg is not None:
        if eg > 20: eps_score = 15
        elif eg >= 0: eps_score = 10

    pe_score = 0
    fpe = data["fwd_pe"]
    if fpe is not None and fpe > 0:
        if fpe <= 15: pe_score = 15
        elif fpe <= 25: pe_score = 10
        elif fpe <= 40: pe_score = 5

    return eps_score + pe_score, {"eps": eps_score, "pe": pe_score}


def score_supply(data: dict) -> tuple[int, dict]:
    vr = data.get("vol_ratio")
    if vr is None: return 0, {"vol": 0}
    if vr > 1.5: sc = 20
    elif vr >= 1.0: sc = 10
    elif vr >= 0.8: sc = 5
    else: sc = 0
    return sc, {"vol": sc}


def score_sentiment(data: dict) -> tuple[int, dict]:
    up = data["upside"]
    if up is None: return 0, {"target": 0}
    if up > 20: sc = 20
    elif up >= 10: sc = 15
    elif up >= 0: sc = 10
    else: sc = 0
    return sc, {"target": sc}


def compute_score(data: dict) -> dict:
    tech_total, tech_detail = score_tech(data)
    fund_total, fund_detail = score_fund(data)
    supply_total, supply_detail = score_supply(data)
    sent_total, sent_detail = score_sentiment(data)
    total = tech_total + fund_total + supply_total + sent_total

    return {
        "total": total,
        "tech": tech_total, "fund": fund_total,
        "supply": supply_total, "sent": sent_total,
        "detail": {**tech_detail, **fund_detail, **supply_detail, **sent_detail},
        "raw": {
            "price": data["price"], "prev_close": data.get("prev_close", 0),
            "change": data.get("change", 0), "change_abs": data.get("change_abs", 0),
            "chg_pct": data.get("change", 0),  # updateHeaderIndices 호환
            "ma20": data.get("ma20"), "ma50": data.get("ma50"),
            "rsi": data.get("rsi"), "vol_ratio": data.get("vol_ratio"),
            "fwd_pe": data.get("fwd_pe"), "eps_growth": data.get("eps_growth"),
            "trailing_eps": data.get("trailing_eps"), "fwd_eps": data.get("fwd_eps"),
            "target_mean": data.get("target_mean"), "upside": data.get("upside"),
            "sector": data.get("sector", ""), "idx_grp": data.get("idx_grp", ""),
        },
        "name": data["name"],
    }


# ══════════════════════════════════════════════════════════════
# 3. 병렬 수집 + 스코어링 (배치 딜레이 포함)
# ══════════════════════════════════════════════════════════════

def process_single_ticker(ticker: str) -> tuple[str, dict | None]:
    """단일 종목 수집 + 스코어링 (재시도 포함)."""
    tk_str, data = fetch_with_retry(ticker)
    if data is None:
        return tk_str, None
    try:
        return tk_str, compute_score(data)
    except Exception:
        return tk_str, None


# ══════════════════════════════════════════════════════════════
# 4. 메인 실행
# ══════════════════════════════════════════════════════════════

def main():
    t_start = time.time()
    _load_index_constituents()   # S&P500, Nasdaq-100 구성종목 캐시
    tickers = build_ticker_universe()
    log.info(f"=== Scoring Generator v3 START ({len(tickers)} tickers, {MAX_WORKERS} threads) ===")

    results = {}
    success = 0
    fail = 0
    processed = 0
    consecutive_zero_batches = 0   # 연속 전원 실패 배치 카운트

    # ── 배치 단위 멀티스레드 수집 (Rate-limit 방어) ──
    for batch_start in range(0, len(tickers), BATCH_SIZE):
        batch = tickers[batch_start:batch_start + BATCH_SIZE]
        batch_num = batch_start // BATCH_SIZE + 1
        total_batches = (len(tickers) + BATCH_SIZE - 1) // BATCH_SIZE
        batch_success = 0

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(process_single_ticker, tk): tk for tk in batch}

            for future in as_completed(futures):
                try:
                    tk, scored = future.result()
                    processed += 1
                    if scored:
                        results[tk] = scored
                        success += 1
                        batch_success += 1
                    else:
                        fail += 1
                except Exception:
                    processed += 1
                    fail += 1

        log.info(
            f"  배치 {batch_num}/{total_batches} 완료 | "
            f"이번 +{batch_success} | 누적 성공 {success} / 실패 {fail} / 처리 {processed}/{len(tickers)}"
        )

        # ── 레이트 리밋 감지: 연속 N배치 전원 실패 시 조기 종료 ──
        if batch_success == 0:
            consecutive_zero_batches += 1
            if consecutive_zero_batches >= FAIL_STREAK_ABORT:
                log.warning(
                    f"⚠ 연속 {FAIL_STREAK_ABORT}배치 전원 실패 — 레이트 리밋 감지!"
                    f" (성공 {success}개 저장 후 조기 종료)"
                )
                break
        else:
            consecutive_zero_batches = 0

        # 배치 간 딜레이 (마지막 배치 제외)
        if batch_start + BATCH_SIZE < len(tickers):
            time.sleep(BATCH_SLEEP)

    elapsed = time.time() - t_start
    log.info(f"수집 완료: 성공 {success}개 / 실패 {fail}개 / 전체 {len(tickers)}개 ({elapsed:.0f}초)")

    if not results:
        log.error("스코어링 결과 없음 — scoring.js 미생성")
        return

    # ── 메타 정보 ──
    output = {
        "_meta": {
            "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "tickers": len(results),
            "max_score": 100,
            "categories": {"tech": 30, "fund": 30, "supply": 20, "sent": 20},
        },
        **results,
    }

    # ── scoring.js 저장 ──
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write("const SCORING_DATA = ")
        json.dump(output, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    size_kb = OUTPUT_PATH.stat().st_size / 1024
    log.info(f"Saved {OUTPUT_PATH.name} ({size_kb:.1f} KB, {len(results)} tickers)")

    # ── 포트폴리오 종목 요약 ──
    log.info("─── 포트폴리오 스코어링 요약 ───")
    for tk in PORTFOLIO_TICKERS:
        sc = results.get(tk)
        if sc:
            bar = "█" * (sc["total"] // 5) + "░" * (20 - sc["total"] // 5)
            log.info(f"  {tk:5s} {sc['total']:3d}/100  [T{sc['tech']:2d} F{sc['fund']:2d} S{sc['supply']:2d} C{sc['sent']:2d}]  {bar}")
        else:
            log.warning(f"  {tk:5s} — 데이터 없음")

    # ── 전체 TOP 20 ──
    log.info("─── 전체 TOP 20 ───")
    ranked = sorted(results.items(), key=lambda x: x[1]["total"], reverse=True)
    for rank, (tk, sc) in enumerate(ranked[:20], 1):
        log.info(
            f"  #{rank:2d} {tk:5s} {sc['total']:3d}/100  "
            f"[T{sc['tech']:2d} F{sc['fund']:2d} S{sc['supply']:2d} C{sc['sent']:2d}]  "
            f"${sc['raw']['price']:.2f}  {sc['name']}"
        )

    log.info(f"=== Scoring Generator v3 DONE ({elapsed:.0f}s) ===")


if __name__ == "__main__":
    main()
