"""
S&P 500 Daily Data Fetcher
==========================
yfinance로 S&P 500 전체 종목 데이터를 수집하여 daily_sp500.xlsx로 저장.
Windows 작업 스케줄러에 등록하여 매일 자동 실행 가능.

사용법:
  1) pip install yfinance openpyxl pandas
  2) python daily_sp500_fetcher.py
  3) 같은 폴더에 daily_sp500.xlsx 생성됨

작업 스케줄러 등록:
  - 프로그램: pythonw.exe (또는 python.exe)
  - 인수: "C:\\...\\daily_sp500_fetcher.py"
  - 시작 위치: 스크립트가 있는 폴더
  - 트리거: 매일 오전 7:00 (미국장 마감 후 한국시간 기준)
"""

import sys
import json
import logging
from pathlib import Path
from datetime import datetime, timedelta

import io
import pandas as pd
import requests
import yfinance as yf

# ── 설정 ──────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent
OUTPUT_FILE = SCRIPT_DIR / "daily_sp500.xlsx"
CHART_FILE = SCRIPT_DIR / "chart_data.json"
ANALYST_FILE = SCRIPT_DIR / "analyst_targets.json"
LOG_FILE = SCRIPT_DIR / "fetcher_log.txt"
PERIOD = "1mo"  # 이평선 계산에 필요한 최소 기간

# 포트폴리오 종목 (애널리스트 목표가 수집 대상)
MYPORT_TICKERS = ["NVDA", "TSM", "MU", "CEG", "SOXL", "IONQ", "CRCL"]

# 대시보드에 필요한 추가 티커 (포트폴리오 + 섹터 ETF)
EXTRA_TICKERS = [
    "SOXL", "IONQ", "CRCL", "TSM", "SPY",
    "XLK", "XLV", "XLF", "XLY", "XLC", "XLI", "XLP", "XLE", "XLU", "XLRE", "XLB",
]

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


def get_sp500_tickers() -> pd.DataFrame:
    """위키피디아에서 S&P 500 종목 리스트를 가져온다."""
    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    tables = pd.read_html(io.StringIO(resp.text))
    df = tables[0][["Symbol", "Security", "GICS Sector", "GICS Sub-Industry"]].copy()
    # BRK.B -> BRK-B (yfinance 형식)
    df["Symbol"] = df["Symbol"].str.replace(".", "-", regex=False)
    df.rename(columns={
        "Symbol": "Ticker",
        "Security": "Company",
        "GICS Sector": "Sector",
        "GICS Sub-Industry": "SubIndustry",
    }, inplace=True)
    return df


def fetch_market_data(tickers: list[str]) -> dict:
    """yfinance로 한 번에 다운로드 (배치)."""
    ticker_str = " ".join(tickers)
    log.info(f"Downloading {len(tickers)} tickers...")
    data = yf.download(ticker_str, period=PERIOD, group_by="ticker", threads=True)
    return data


def compute_indicators(raw_data, tickers: list[str]) -> list[dict]:
    """종목별 기술 지표를 계산한다."""
    rows = []
    today = datetime.now().strftime("%Y-%m-%d")

    for ticker in tickers:
        try:
            # 단일 종목일 때와 다중 종목일 때 컬럼 구조가 다름
            if len(tickers) == 1:
                df = raw_data.copy()
            else:
                if ticker not in raw_data.columns.get_level_values(0):
                    continue
                df = raw_data[ticker].copy()

            df = df.dropna(subset=["Close"])
            if len(df) < 2:
                continue

            # 기본 가격 데이터
            close = df["Close"].iloc[-1]
            prev_close = df["Close"].iloc[-2]
            daily_chg = (close - prev_close) / prev_close if prev_close else 0

            # 이동평균선
            ma5 = df["Close"].rolling(5).mean().iloc[-1] if len(df) >= 5 else None
            ma20 = df["Close"].rolling(20).mean().iloc[-1] if len(df) >= 20 else None

            # 골든크로스/데드크로스 판정
            cross_signal = ""
            if ma5 is not None and ma20 is not None:
                if len(df) >= 6:
                    prev_ma5 = df["Close"].rolling(5).mean().iloc[-2]
                    prev_ma20 = df["Close"].rolling(20).mean().iloc[-2]
                    if prev_ma5 is not None and prev_ma20 is not None:
                        if prev_ma5 < prev_ma20 and ma5 >= ma20:
                            cross_signal = "GOLDEN_CROSS"
                        elif prev_ma5 > prev_ma20 and ma5 <= ma20:
                            cross_signal = "DEAD_CROSS"

            # 거래량 분석
            vol = df["Volume"].iloc[-1]
            vol_ma5 = df["Volume"].rolling(5).mean().iloc[-1] if len(df) >= 5 else vol
            vol_ratio = vol / vol_ma5 if vol_ma5 else 0

            # 최근 5일 수익률
            chg_5d = None
            if len(df) >= 6:
                chg_5d = (close - df["Close"].iloc[-6]) / df["Close"].iloc[-6]

            # 20일 고가/저가 대비
            high_20d = df["High"].rolling(20).max().iloc[-1] if len(df) >= 20 else None
            low_20d = df["Low"].rolling(20).min().iloc[-1] if len(df) >= 20 else None
            pct_from_20d_high = (close - high_20d) / high_20d if high_20d else None
            pct_from_20d_low = (close - low_20d) / low_20d if low_20d else None

            # 변동성 (20일 수익률 표준편차 연율화)
            if len(df) >= 20:
                returns = df["Close"].pct_change().dropna()
                volatility = returns.tail(20).std() * (252 ** 0.5)
            else:
                volatility = None

            # RSI (14일)
            rsi = None
            if len(df) >= 15:
                delta = df["Close"].diff()
                gain = delta.where(delta > 0, 0).rolling(14).mean().iloc[-1]
                loss = (-delta.where(delta < 0, 0)).rolling(14).mean().iloc[-1]
                if loss != 0:
                    rs = gain / loss
                    rsi = 100 - (100 / (1 + rs))

            # 종합 점수 (0~100) - 단순 가중합
            score = 0
            score_parts = 0
            if daily_chg is not None:
                score += max(min(daily_chg * 500 + 50, 100), 0)  # 일간변동
                score_parts += 1
            if chg_5d is not None:
                score += max(min(chg_5d * 200 + 50, 100), 0)     # 5일변동
                score_parts += 1
            if vol_ratio > 0:
                score += max(min(vol_ratio * 25, 100), 0)         # 거래량
                score_parts += 1
            if ma5 is not None and ma20 is not None and ma20 != 0:
                ma_score = ((ma5 / ma20) - 1) * 1000 + 50
                score += max(min(ma_score, 100), 0)               # 이평선 배열
                score_parts += 1
            if rsi is not None:
                # RSI 50 근처가 가장 좋고, 과매수/과매도 감점
                rsi_score = 100 - abs(rsi - 55) * 1.5
                score += max(min(rsi_score, 100), 0)
                score_parts += 1

            composite_score = round(score / score_parts, 1) if score_parts else 0

            rows.append({
                "Date": today,
                "Ticker": ticker,
                "Close": round(float(close), 2),
                "Prev_Close": round(float(prev_close), 2),
                "Daily_Chg_Pct": round(float(daily_chg) * 100, 2),
                "MA5": round(float(ma5), 2) if ma5 is not None else None,
                "MA20": round(float(ma20), 2) if ma20 is not None else None,
                "MA5_vs_MA20": round(float(ma5 - ma20), 2) if (ma5 is not None and ma20 is not None) else None,
                "Cross_Signal": cross_signal,
                "Volume": int(vol),
                "Vol_MA5": round(float(vol_ma5)) if vol_ma5 else None,
                "Vol_Ratio": round(float(vol_ratio), 2),
                "Chg_5D_Pct": round(float(chg_5d) * 100, 2) if chg_5d is not None else None,
                "High_20D": round(float(high_20d), 2) if high_20d is not None else None,
                "Low_20D": round(float(low_20d), 2) if low_20d is not None else None,
                "Pct_From_20D_High": round(float(pct_from_20d_high) * 100, 2) if pct_from_20d_high is not None else None,
                "Pct_From_20D_Low": round(float(pct_from_20d_low) * 100, 2) if pct_from_20d_low is not None else None,
                "Volatility_Ann": round(float(volatility) * 100, 1) if volatility is not None else None,
                "RSI_14": round(float(rsi), 1) if rsi is not None else None,
                "Composite_Score": composite_score,
            })

        except Exception as e:
            log.warning(f"  [{ticker}] skipped: {e}")

    return rows


def extract_chart_data(raw_data, tickers: list[str]) -> dict:
    """대시보드 차트용 20일 데이터 추출."""
    chart = {}
    for ticker in tickers:
        try:
            if len(tickers) == 1:
                df = raw_data.copy()
            else:
                if ticker not in raw_data.columns.get_level_values(0):
                    continue
                df = raw_data[ticker].copy()

            df = df.dropna(subset=["Close"])
            if len(df) < 2:
                continue

            ma5 = df["Close"].rolling(5).mean()
            ma20 = df["Close"].rolling(20).mean()
            last20 = df.tail(20)
            entries = []
            for date_idx, row in last20.iterrows():
                entries.append({
                    "d": date_idx.strftime("%m/%d"),
                    "c": round(float(row["Close"]), 2),
                    "h": round(float(row["High"]), 2),
                    "l": round(float(row["Low"]), 2),
                    "v": int(row["Volume"]),
                    "a": round(float(ma5.loc[date_idx]), 2) if pd.notna(ma5.loc[date_idx]) else round(float(row["Close"]), 2),
                    "b": round(float(ma20.loc[date_idx]), 2) if pd.notna(ma20.loc[date_idx]) else round(float(row["Close"]), 2),
                })
            if entries:
                chart[ticker] = entries
        except Exception:
            pass
    return chart


def fetch_analyst_targets():
    """yfinance .info를 통해 포트폴리오 종목의 애널리스트 목표가 데이터 수집."""
    import time
    targets = {}
    for tk in MYPORT_TICKERS:
        try:
            info = yf.Ticker(tk).info
            cp = info.get("currentPrice") or info.get("regularMarketPrice")
            tgt_lo = info.get("targetLowPrice")
            tgt_mean = info.get("targetMeanPrice")
            tgt_hi = info.get("targetHighPrice")
            tgt_med = info.get("targetMedianPrice")
            n_analysts = info.get("numberOfAnalystOpinions")
            rec_key = info.get("recommendationKey", "")
            rec_mean = info.get("recommendationMean")
            targets[tk] = {
                "currentPrice": round(float(cp), 2) if cp else None,
                "targetLowPrice": round(float(tgt_lo), 2) if tgt_lo else None,
                "targetMeanPrice": round(float(tgt_mean), 2) if tgt_mean else None,
                "targetHighPrice": round(float(tgt_hi), 2) if tgt_hi else None,
                "targetMedianPrice": round(float(tgt_med), 2) if tgt_med else None,
                "numberOfAnalystOpinions": int(n_analysts) if n_analysts else 0,
                "recommendationKey": rec_key,
                "recommendationMean": round(float(rec_mean), 2) if rec_mean else None,
            }
            log.info(f"  [{tk}] Price=${cp}, TargetMean=${tgt_mean}, "
                     f"Low=${tgt_lo}, High=${tgt_hi}, Analysts={n_analysts}, Rec={rec_key}")
            time.sleep(0.3)  # API 속도 제한 방지
        except Exception as e:
            log.warning(f"  [{tk}] analyst data failed: {e}")
            targets[tk] = {
                "currentPrice": None, "targetLowPrice": None,
                "targetMeanPrice": None, "targetHighPrice": None,
                "targetMedianPrice": None, "numberOfAnalystOpinions": 0,
                "recommendationKey": "", "recommendationMean": None,
            }
    return targets


def main():
    log.info("=" * 60)
    log.info("S&P 500 Daily Fetcher START")
    log.info("=" * 60)

    # 1) S&P 500 종목 리스트
    try:
        sp500_info = get_sp500_tickers()
        tickers = sp500_info["Ticker"].tolist()
        log.info(f"S&P 500 ticker count: {len(tickers)}")
    except Exception as e:
        log.error(f"Failed to fetch S&P 500 list: {e}")
        sys.exit(1)

    # 2) 추가 티커 병합 (포트폴리오 + 섹터 ETF)
    extra = [t for t in EXTRA_TICKERS if t not in tickers]
    all_tickers = tickers + extra
    log.info(f"Extra tickers added: {len(extra)} → total {len(all_tickers)}")

    # 3) 시장 데이터 다운로드
    try:
        raw_data = fetch_market_data(all_tickers)
    except Exception as e:
        log.error(f"Failed to download market data: {e}")
        sys.exit(1)

    # 4) 지표 계산 (S&P 500 종목만)
    rows = compute_indicators(raw_data, tickers)
    log.info(f"Computed indicators for {len(rows)} tickers")

    if not rows:
        log.error("No data computed. Exiting.")
        sys.exit(1)

    # 4) DataFrame 병합 (종목 정보 + 지표)
    df_indicators = pd.DataFrame(rows)
    df_result = df_indicators.merge(
        sp500_info[["Ticker", "Company", "Sector", "SubIndustry"]],
        on="Ticker",
        how="left",
    )

    # 열 순서 정리
    col_order = [
        "Date", "Ticker", "Company", "Sector", "SubIndustry",
        "Close", "Prev_Close", "Daily_Chg_Pct",
        "MA5", "MA20", "MA5_vs_MA20", "Cross_Signal",
        "Volume", "Vol_MA5", "Vol_Ratio",
        "Chg_5D_Pct",
        "High_20D", "Low_20D", "Pct_From_20D_High", "Pct_From_20D_Low",
        "Volatility_Ann", "RSI_14", "Composite_Score",
    ]
    df_result = df_result[[c for c in col_order if c in df_result.columns]]
    df_result.sort_values("Composite_Score", ascending=False, inplace=True)
    df_result.reset_index(drop=True, inplace=True)

    # 5) 엑셀 저장 (시트 2개: Summary + Raw)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Sheet 1: 요약 (점수순 정렬)
        df_result.to_excel(writer, sheet_name="Summary", index=False)

        # Sheet 2: 섹터별 집계
        sector_agg = df_result.groupby("Sector").agg(
            Avg_Daily_Chg=("Daily_Chg_Pct", "mean"),
            Avg_Chg_5D=("Chg_5D_Pct", "mean"),
            Avg_Vol_Ratio=("Vol_Ratio", "mean"),
            Avg_RSI=("RSI_14", "mean"),
            Avg_Score=("Composite_Score", "mean"),
            Count=("Ticker", "count"),
            Golden_Cross=("Cross_Signal", lambda x: (x == "GOLDEN_CROSS").sum()),
            Dead_Cross=("Cross_Signal", lambda x: (x == "DEAD_CROSS").sum()),
        ).round(2).sort_values("Avg_Score", ascending=False)
        sector_agg.to_excel(writer, sheet_name="Sector_Summary")

        # Sheet 3: 메타 정보
        meta = pd.DataFrame([{
            "Generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Total_Tickers": len(tickers),
            "Successful": len(rows),
            "Failed": len(tickers) - len(rows),
            "Period": PERIOD,
            "Top1_Ticker": df_result.iloc[0]["Ticker"] if len(df_result) > 0 else "",
            "Top1_Score": df_result.iloc[0]["Composite_Score"] if len(df_result) > 0 else "",
        }])
        meta.to_excel(writer, sheet_name="Meta", index=False)

    log.info(f"Saved to {OUTPUT_FILE}")

    # 6) 대시보드용 차트 데이터 저장
    chart_data = extract_chart_data(raw_data, all_tickers)
    with open(CHART_FILE, "w", encoding="utf-8") as f:
        json.dump(chart_data, f, separators=(",", ":"))
    log.info(f"Saved chart data for {len(chart_data)} tickers to {CHART_FILE}")

    log.info(f"Top 10 by Composite Score:")
    for i, row in df_result.head(10).iterrows():
        log.info(f"  {i+1:>2}. {row['Ticker']:<6} Score={row['Composite_Score']:>5}  "
                 f"Close=${row['Close']:<8}  Chg={row['Daily_Chg_Pct']:>+6.2f}%  RSI={row.get('RSI_14','N/A')}")

    # 7) 애널리스트 목표가 데이터 수집
    log.info("Fetching analyst target data for portfolio...")
    analyst = fetch_analyst_targets()
    with open(ANALYST_FILE, "w", encoding="utf-8") as f:
        json.dump(analyst, f, ensure_ascii=False, indent=2)
    valid_count = sum(1 for v in analyst.values() if v.get("targetMeanPrice"))
    log.info(f"Saved analyst targets: {valid_count}/{len(analyst)} tickers with data → {ANALYST_FILE}")

    log.info("DONE")


if __name__ == "__main__":
    main()
