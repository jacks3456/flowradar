#!/usr/bin/env python3
"""Fetch Korea stock-market leverage data from KOFIA FreeSIS."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
from pathlib import Path

import requests


META_URL = "https://freesis.kofia.or.kr/meta/getMetaDataList.do"
SOURCE_URL = "https://freesis.kofia.or.kr/"

SERVICES = {
    "credit": "STATSCU0100000070BO",
    "funds": "STATSCU0100000060BO",
    "kospi": "STATSCU0100000020BO",
    "kosdaq": "STATSCU0100000030BO",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="韩国股市杠杆率网站数据生成器")
    parser.add_argument("--start", default="19980701", help="开始日期 YYYYMMDD，默认官方信用融资最早可查日 19980701")
    parser.add_argument("--years", type=int, default=0, help="仅抓取最近多少年；默认 0 表示从最早日期开始")
    parser.add_argument("--end", help="结束日期 YYYYMMDD，默认今天")
    parser.add_argument("--site-dir", default="site", help="网站目录，默认 site")
    parser.add_argument("--output", help="JSON 输出路径")
    return parser.parse_args()


class KofiaClient:
    def __init__(self) -> None:
        self.masked_value_count = 0
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0",
                "Accept": "application/json,text/plain,*/*",
                "Content-Type": "application/json",
                "Referer": SOURCE_URL,
            }
        )

    def fetch_range(self, service: str, start: dt.date, end: dt.date) -> list[dict]:
        payload = {
            "dmSearch": {
                "OBJ_NM": SERVICES[service],
                "tmpV1": "D",
                "tmpV45": start.strftime("%Y%m%d"),
                "tmpV46": end.strftime("%Y%m%d"),
                "tmpV40": "08" if service in {"kospi", "kosdaq"} else "06",
                "tmpV41": "04" if service in {"kospi", "kosdaq"} else "원",
            }
        }
        last_error: Exception | None = None
        for attempt in range(3):
            try:
                response = self.session.post(META_URL, json=payload, timeout=90)
                break
            except requests.RequestException as exc:
                last_error = exc
                if attempt == 2:
                    raise
        else:
            raise RuntimeError(f"KOFIA {service} 请求失败: {last_error}")
        response.raise_for_status()
        text = response.text
        if "#" in text:
            self.masked_value_count += text.count("#")
            text = re.sub(r"(-?\d+(?:\.\d+)?)#+", lambda match: match.group(0).replace("#", "0"), text)
        data = json.loads(text)
        rows = data.get("ds1") or []
        if not rows:
            raise RuntimeError(f"KOFIA {service} 未返回数据")
        return rows

    def fetch(self, service: str, start: dt.date, end: dt.date) -> list[dict]:
        """Fetch in five-year chunks because FreeSIS rejects very long requests."""
        by_date: dict[str, dict] = {}
        chunk_start = start
        while chunk_start <= end:
            chunk_end = min(chunk_start + dt.timedelta(days=5 * 366 - 1), end)
            for row in self.fetch_range(service, chunk_start, chunk_end):
                if row.get("TMPV1"):
                    by_date[str(row["TMPV1"])] = row
            chunk_start = chunk_end + dt.timedelta(days=1)
        if not by_date:
            raise RuntimeError(f"KOFIA {service} 分段查询未返回数据")
        return list(by_date.values())


def as_float(value: object) -> float:
    try:
        result = float(value)
        return result if math.isfinite(result) else 0.0
    except (TypeError, ValueError):
        return 0.0


def pct(numerator: float, denominator: float) -> float:
    return numerator / denominator * 100 if denominator else 0.0


def change_pct(current: float, previous: float) -> float:
    return (current / previous - 1) * 100 if previous else 0.0


def percentile(values: list[float], current: float) -> float:
    clean = sorted(value for value in values if math.isfinite(value))
    if not clean:
        return 0.0
    below = sum(value < current for value in clean)
    equal = sum(value == current for value in clean)
    return (below + 0.5 * equal) / len(clean) * 100


def round_or_none(value: float, digits: int = 4) -> float | None:
    return round(value, digits) if math.isfinite(value) else None


def index_by_date(rows: list[dict]) -> dict[str, dict]:
    return {str(row.get("TMPV1")): row for row in rows if row.get("TMPV1")}


def build_rows(raw: dict[str, list[dict]]) -> list[dict]:
    indexed = {name: index_by_date(rows) for name, rows in raw.items()}
    dates = sorted(set.intersection(*(set(rows) for rows in indexed.values())))
    output: list[dict] = []
    turnover_history: list[float] = []

    for date in dates:
        credit = indexed["credit"][date]
        funds = indexed["funds"][date]
        kospi = indexed["kospi"][date]
        kosdaq = indexed["kosdaq"][date]

        total_credit = as_float(credit.get("TMPV2"))
        kospi_credit = as_float(credit.get("TMPV3"))
        kosdaq_credit = as_float(credit.get("TMPV4"))
        deposits = as_float(funds.get("TMPV2"))
        receivables = as_float(funds.get("TMPV5"))
        forced_sales = as_float(funds.get("TMPV6"))
        forced_sale_rate = as_float(funds.get("TMPV7"))
        kospi_cap = as_float(kospi.get("TMPV5"))
        kosdaq_cap = as_float(kosdaq.get("TMPV5"))
        turnover = as_float(kospi.get("TMPV4")) + as_float(kosdaq.get("TMPV4"))
        turnover_history.append(turnover)
        avg_turnover_20d = sum(turnover_history[-20:]) / min(len(turnover_history), 20)

        row = {
            "date": dt.datetime.strptime(date, "%Y%m%d").strftime("%Y-%m-%d"),
            "credit100m": total_credit / 100_000_000,
            "kospiCredit100m": kospi_credit / 100_000_000,
            "kosdaqCredit100m": kosdaq_credit / 100_000_000,
            "marketCap100m": (kospi_cap + kosdaq_cap) / 100_000_000,
            "kospiMarketCap100m": kospi_cap / 100_000_000,
            "kosdaqMarketCap100m": kosdaq_cap / 100_000_000,
            "turnover100m": turnover / 100_000_000,
            "avgTurnover20d100m": avg_turnover_20d / 100_000_000,
            "deposits100m": deposits / 100_000_000,
            "receivables100m": receivables / 100_000_000,
            "forcedSales100m": forced_sales / 100_000_000,
            "forcedSaleRatePct": forced_sale_rate,
            "leveragePct": pct(total_credit, kospi_cap + kosdaq_cap),
            "kospiLeveragePct": pct(kospi_credit, kospi_cap),
            "kosdaqLeveragePct": pct(kosdaq_credit, kosdaq_cap),
            "turnoverDays": total_credit / avg_turnover_20d if avg_turnover_20d else 0.0,
            "creditDepositPct": pct(total_credit, deposits),
            "kospiIndex": as_float(kospi.get("TMPV2")),
            "kosdaqIndex": as_float(kosdaq.get("TMPV2")),
        }
        if len(output) >= 20:
            row["credit20dChangePct"] = change_pct(total_credit, output[-20]["credit100m"] * 100_000_000)
        else:
            row["credit20dChangePct"] = 0.0
        output.append(row)

    metric_names = [
        "leveragePct",
        "kospiLeveragePct",
        "kosdaqLeveragePct",
        "turnoverDays",
        "creditDepositPct",
        "forcedSales100m",
        "credit20dChangePct",
    ]
    histories = {name: [as_float(row[name]) for row in output] for name in metric_names}
    weights = {
        "leveragePct": 0.35,
        "turnoverDays": 0.20,
        "creditDepositPct": 0.15,
        "forcedSales100m": 0.15,
        "credit20dChangePct": 0.15,
    }
    for row in output:
        component_percentiles = {
            name: percentile(histories[name], as_float(row[name])) for name in metric_names
        }
        row["leveragePercentile"] = component_percentiles["leveragePct"]
        row["kospiLeveragePercentile"] = component_percentiles["kospiLeveragePct"]
        row["kosdaqLeveragePercentile"] = component_percentiles["kosdaqLeveragePct"]
        row["stressScore"] = sum(component_percentiles[name] * weight for name, weight in weights.items())
        for key, value in list(row.items()):
            if isinstance(value, float):
                row[key] = round_or_none(value)
    return output


def stress_label(score: float) -> str:
    if score >= 95:
        return "极端"
    if score >= 85:
        return "拥挤"
    if score >= 70:
        return "偏高"
    if score >= 40:
        return "正常"
    return "偏低"


def build_payload(rows: list[dict], start: dt.date, end: dt.date, masked_value_count: int = 0) -> dict:
    if not rows:
        raise RuntimeError("合并后没有共同交易日")
    latest = rows[-1]
    return {
        "meta": {
            "market": "Korea",
            "frequency": "daily",
            "unit": "亿韩元",
            "start": rows[0]["date"],
            "end": rows[-1]["date"],
            "requestedStart": start.isoformat(),
            "requestedEnd": end.isoformat(),
            "latestDate": rows[-1]["date"],
            "generatedAt": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
            "source": "KOFIA FreeSIS（市值与成交额的资料源为 KRX）",
            "sourceUrl": SOURCE_URL,
            "historyDays": len(rows),
            "maskedValueCount": masked_value_count,
            "maskedValueHandling": "KOFIA returns some large values with trailing #; hidden trailing digits are treated as 0.",
        },
        "methodology": {
            "leveragePct": "信用交易融资余额 / KOSPI与KOSDAQ总市值",
            "turnoverDays": "信用交易融资余额 / 全市场20日平均成交额",
            "creditDepositPct": "信用交易融资余额 / 投资者保证金",
            "stressScore": "历史分位加权：市值杠杆35%、成交杠杆20%、融资/保证金15%、强平15%、融资20日变化15%",
        },
        "latest": {**latest, "stressLabel": stress_label(as_float(latest["stressScore"]))},
        "rows": rows,
    }


def main() -> int:
    args = parse_args()
    end = dt.datetime.strptime(args.end, "%Y%m%d").date() if args.end else dt.date.today()
    start = dt.datetime.strptime(args.start, "%Y%m%d").date()
    if args.years > 0:
        start = end - dt.timedelta(days=args.years * 366)
    client = KofiaClient()
    raw = {name: client.fetch(name, start, end) for name in SERVICES}
    rows = build_rows(raw)
    payload = build_payload(rows, start, end, masked_value_count=client.masked_value_count)
    output = Path(args.output) if args.output else Path(args.site_dir) / "data" / "korea-leverage" / "latest.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已生成 {output}，最新日期 {payload['meta']['latestDate']}，共 {len(rows)} 个交易日")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
