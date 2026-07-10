#!/usr/bin/env python3
"""Fetch TWSE daily institutional money flow and generate public site data."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import time
from pathlib import Path
from typing import Iterable

import pandas as pd
import requests


TWSE_API_URL = "https://www.twse.com.tw/rwd/zh/fund/BFI82U"
INVESTORS = ["外资", "投信", "自营商"]


class TwseClient:
    def __init__(self, pause_seconds: float = 0.35) -> None:
        self.pause_seconds = pause_seconds
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/json,text/plain,*/*",
                "Referer": "https://www.twse.com.tw/zh/trading/foreign/bfi82u.html",
            }
        )

    def get_daily(self, trading_date: dt.date | None = None) -> dict:
        params = {"response": "json", "type": "day"}
        if trading_date is not None:
            ymd = trading_date.strftime("%Y%m%d")
            params.update({"dayDate": ymd, "weekDate": ymd})
        last_error: Exception | None = None
        for attempt in range(3):
            response = self.session.get(TWSE_API_URL, params=params, timeout=30)
            try:
                response.raise_for_status()
                payload = response.json()
                time.sleep(self.pause_seconds)
                return payload
            except Exception as exc:
                last_error = exc
                time.sleep(self.pause_seconds * (attempt + 2))
        raise RuntimeError(f"TWSE 响应无法解析为 JSON: {params}") from last_error


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="TWSE 台股每日三大法人资金流网站数据生成器"
    )
    parser.add_argument("--days", type=int, default=365, help="最多抓取最近多少个自然日，默认 365")
    parser.add_argument("--min-rows", type=int, default=180, help="至少尝试收集多少个交易日，默认 180")
    parser.add_argument("--end", default=None, help="结束日期 YYYYMMDD；默认 TWSE 最新交易日")
    parser.add_argument(
        "--site-dir",
        default="site",
        help="公开网站目录，默认 site；脚本会写入 site/data/taiwan/latest.json",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="JSON 输出路径；默认 site/data/taiwan/latest.json",
    )
    return parser.parse_args()


def parse_number(value: object) -> float:
    text = str(value or "").strip()
    if text in {"", "-"}:
        return 0.0
    text = re.sub(r"[^\d.-]", "", text)
    if text in {"", "-"}:
        return 0.0
    return float(text)


def parse_date(value: str) -> dt.date:
    return dt.datetime.strptime(value, "%Y%m%d").date()


def money_to_100m_twd(value: object) -> float:
    return parse_number(value) / 100_000_000


def parse_daily_payload(payload: dict) -> dict | None:
    if payload.get("stat") != "OK":
        return None
    rows = payload.get("data") or []
    if not rows:
        return None

    raw = {row[0]: row for row in rows if row}
    broker_own = raw.get("自營商(自行買賣)", ["", 0, 0, 0])
    broker_hedge = raw.get("自營商(避險)", ["", 0, 0, 0])
    investment_trust = raw.get("投信", ["", 0, 0, 0])
    foreign = raw.get("外資及陸資(不含外資自營商)", ["", 0, 0, 0])
    total = raw.get("合計", ["", 0, 0, 0])
    dealer = (
        parse_number(broker_own[3] if len(broker_own) > 3 else 0)
        + parse_number(broker_hedge[3] if len(broker_hedge) > 3 else 0)
    )
    trading_date = parse_date(str(payload["date"]))
    return {
        "日期": trading_date,
        "外资": money_to_100m_twd(foreign[3] if len(foreign) > 3 else 0),
        "投信": money_to_100m_twd(investment_trust[3] if len(investment_trust) > 3 else 0),
        "自营商": dealer / 100_000_000,
        "合计": money_to_100m_twd(total[3] if len(total) > 3 else 0),
    }


def current_streak(values: Iterable[float]) -> int:
    cleaned = [value for value in values if value != 0]
    if not cleaned:
        return 0
    sign = 1 if cleaned[-1] > 0 else -1
    streak = 0
    for value in reversed(cleaned):
        if (value > 0 and sign > 0) or (value < 0 and sign < 0):
            streak += 1
        else:
            break
    return streak * sign


def streak_label(streak: int) -> str:
    if streak > 0:
        return f"连续净买入 {streak} 天"
    if streak < 0:
        return f"连续净卖出 {abs(streak)} 天"
    return "无连续方向"


def frame_to_records(frame: pd.DataFrame) -> list[dict]:
    if frame.empty:
        return []
    normalized = frame.copy()
    for column in normalized.columns:
        if pd.api.types.is_datetime64_any_dtype(normalized[column]):
            normalized[column] = normalized[column].dt.strftime("%Y-%m-%d")
        elif all(isinstance(value, dt.date) or pd.isna(value) for value in normalized[column]):
            normalized[column] = normalized[column].map(
                lambda value: value.strftime("%Y-%m-%d") if isinstance(value, dt.date) else value
            )
    return normalized.to_dict(orient="records")


def build_summary(market_flow: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for investor in INVESTORS:
        streak = current_streak(market_flow[investor])
        rows.append(
            {
                "对象": "TWSE市场",
                "指标": f"{investor}净买卖金额",
                "最近值": market_flow[investor].iloc[-1] if not market_flow.empty else 0,
                "单位": "亿新台币",
                "连续状态": streak_label(streak),
            }
        )
    return pd.DataFrame(rows)


def add_cumulative_columns(frame: pd.DataFrame) -> list[dict]:
    if frame.empty:
        return []
    view = frame.copy()
    for investor in INVESTORS:
        view[f"{investor}累计"] = view[investor].cumsum()
    view["合计累计"] = view["合计"].cumsum()
    return frame_to_records(view)


def build_payload(market_flow: pd.DataFrame, summary: pd.DataFrame) -> dict:
    latest = market_flow.iloc[-1] if not market_flow.empty else {}
    latest_date = latest["日期"].strftime("%Y-%m-%d") if len(latest) else "-"
    start = market_flow["日期"].min().strftime("%Y-%m-%d") if not market_flow.empty else "-"
    return {
        "meta": {
            "start": start,
            "end": latest_date,
            "market": "TWSE",
            "frequency": "daily",
            "unit": "亿新台币",
            "source": "TWSE 三大法人買賣金額統計表",
            "sourceUrl": TWSE_API_URL,
            "generatedAt": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "latestDate": latest_date,
        },
        "summary": frame_to_records(summary),
        "marketFlow": add_cumulative_columns(market_flow),
        "stockFlows": {},
    }


def write_json(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def collect_rows(client: TwseClient, end_date: dt.date, days: int, min_rows: int) -> list[dict]:
    rows = []
    seen_dates: set[dt.date] = set()
    attempts = max(days, min_rows * 2)
    for offset in range(attempts):
        trading_date = end_date - dt.timedelta(days=offset)
        try:
            payload = client.get_daily(trading_date)
        except RuntimeError:
            continue
        row = parse_daily_payload(payload)
        if row and row["日期"] not in seen_dates:
            rows.append(row)
            seen_dates.add(row["日期"])
        if offset + 1 >= days and len(rows) >= min_rows:
            break
    return rows


def main() -> None:
    args = parse_args()
    output_path = Path(args.output) if args.output else Path(args.site_dir) / "data" / "taiwan" / "latest.json"
    client = TwseClient()

    try:
        if args.end:
            end_date = parse_date(args.end)
        else:
            latest_payload = client.get_daily()
            latest = parse_daily_payload(latest_payload)
            if not latest:
                raise RuntimeError(f"TWSE 最新交易日资料不可用: {latest_payload.get('stat')}")
            end_date = latest["日期"]

        rows = collect_rows(client, end_date, args.days, args.min_rows)
        if not rows:
            raise RuntimeError("未取得 TWSE 三大法人买卖金额数据")

        market_flow = pd.DataFrame(rows).sort_values("日期").reset_index(drop=True)
        summary = build_summary(market_flow)
        payload = build_payload(market_flow, summary)
        write_json(output_path, payload)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print(f"已生成台股网站数据: {output_path}")
    print(f"最新交易日: {payload['meta']['latestDate']}")
    print(f"交易日数量: {len(payload['marketFlow'])}")


if __name__ == "__main__":
    main()
