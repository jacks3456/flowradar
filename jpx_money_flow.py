#!/usr/bin/env python3
"""Fetch JPX weekly investor money flow and generate public site data."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin

import pandas as pd
import requests


JPX_BASE_URL = "https://www.jpx.co.jp"
INVESTOR_TYPE_PATH = "/english/markets/statistics-equities/investor-type/index.html"
ARCHIVE_PATH_TEMPLATE = (
    "/english/markets/statistics-equities/investor-type/00-00-archives-{index:02d}.html"
)
VALUE_XLS_PATTERN = re.compile(r'href="([^"]*stock_val_1_[^"]*\.xls)"')
DEFAULT_SHEET = "Tokyo & Nagoya"
INVESTOR_ROWS = {
    "外资": ["Foreigners"],
    "散户": ["Individuals"],
    "机构": ["Institutions", "Proprietary", "Securities Cos."],
}


@dataclass(frozen=True)
class WeeklyFile:
    url: str
    code: str


class JpxClient:
    def __init__(self, pause_seconds: float = 0.2) -> None:
        self.pause_seconds = pause_seconds
        self.session = requests.Session()
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            }
        )

    def get_text(self, path_or_url: str) -> str:
        response = self.session.get(urljoin(JPX_BASE_URL, path_or_url), timeout=30)
        response.raise_for_status()
        time.sleep(self.pause_seconds)
        return response.text

    def get_bytes(self, path_or_url: str) -> bytes:
        response = self.session.get(urljoin(JPX_BASE_URL, path_or_url), timeout=30)
        response.raise_for_status()
        time.sleep(self.pause_seconds)
        return response.content

    def weekly_value_files(self, weeks: int, archive_pages: int) -> list[WeeklyFile]:
        pages = [INVESTOR_TYPE_PATH]
        pages.extend(
            ARCHIVE_PATH_TEMPLATE.format(index=index)
            for index in range(max(archive_pages, 0))
        )
        seen: set[str] = set()
        files: list[WeeklyFile] = []
        for page in pages:
            html = self.get_text(page)
            for match in VALUE_XLS_PATTERN.finditer(html):
                path = match.group(1)
                url = urljoin(JPX_BASE_URL, path)
                if url in seen:
                    continue
                seen.add(url)
                files.append(WeeklyFile(url=url, code=file_code(path)))
                if len(files) >= weeks:
                    return files
        return files


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="JPX 日股周频投资者资金流网站数据生成器"
    )
    parser.add_argument("--weeks", type=int, default=52, help="抓取最近多少周，默认 52")
    parser.add_argument(
        "--archive-pages",
        type=int,
        default=2,
        help="额外读取多少个年度归档页，默认 2（当前年度与上一年度）",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"JPX Excel sheet 名称，默认 {DEFAULT_SHEET}",
    )
    parser.add_argument(
        "--site-dir",
        default="site",
        help="公开网站目录，默认 site；脚本会写入 site/data/japan/latest.json",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="JSON 输出路径；默认 site/data/japan/latest.json",
    )
    parser.add_argument(
        "--raw-dir",
        default=None,
        help="可选：保存下载的 JPX 原始 xls 文件目录",
    )
    return parser.parse_args()


def file_code(path: str) -> str:
    match = re.search(r"stock_val_1_(\d+)\.xls", path)
    return match.group(1) if match else Path(path).name


def parse_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip()
    if text in {"", "-"}:
        return 0.0
    text = re.sub(r"[^\d.-]", "", text)
    if text in {"", "-"}:
        return 0.0
    return float(text)


def parse_week_info(raw: pd.DataFrame) -> tuple[dt.date, str]:
    text = " ".join(str(value) for value in raw.iloc[:6, :4].to_numpy().ravel() if not pd.isna(value))
    year_match = re.search(r"(20\d{2})年", text) or re.search(r"(20\d{2})/\d+\s+week", text)
    range_match = re.search(r"\(\s*(\d{1,2})/(\d{1,2})\s*-\s*(\d{1,2})/(\d{1,2})\s*\)", text)
    if not year_match or not range_match:
        raise ValueError(f"JPX Excel 周期信息无法识别: {text}")
    year = int(year_match.group(1))
    start_month = int(range_match.group(1))
    end_month = int(range_match.group(3))
    end_day = int(range_match.group(4))
    if start_month == 12 and end_month == 1:
        year += 1
    end_date = dt.date(year, end_month, end_day)
    label = f"{range_match.group(1)}/{range_match.group(2)} - {range_match.group(3)}/{range_match.group(4)}"
    return end_date, label


def balance_column(raw: pd.DataFrame) -> int:
    candidates: list[int] = []
    for column in range(raw.shape[1]):
        values = [str(raw.iat[row, column]) for row in range(min(14, raw.shape[0])) if not pd.isna(raw.iat[row, column])]
        if any("Balance" in value or "差引き" in value for value in values):
            candidates.append(column)
    if not candidates:
        raise ValueError("JPX Excel 中未找到 Balance 列")
    return max(candidates)


def find_label_row(raw: pd.DataFrame, label: str) -> int:
    for row in range(raw.shape[0]):
        for column in range(min(3, raw.shape[1])):
            value = raw.iat[row, column]
            if isinstance(value, str) and value.strip() == label:
                return row
    raise ValueError(f"JPX Excel 中未找到分类: {label}")


def net_balance(raw: pd.DataFrame, label: str, column: int) -> float:
    row = find_label_row(raw, label)
    values = [
        raw.iat[candidate, column]
        for candidate in (row - 1, row, row + 1)
        if 0 <= candidate < raw.shape[0]
    ]
    for value in values:
        if not pd.isna(value):
            return parse_number(value)
    return 0.0


def parse_weekly_value_excel(content: bytes, source_url: str, sheet_name: str) -> dict:
    raw = pd.read_excel(BytesIO(content), sheet_name=sheet_name, header=None, engine="xlrd")
    end_date, period_label = parse_week_info(raw)
    column = balance_column(raw)
    values: dict[str, float] = {}
    for display_name, labels in INVESTOR_ROWS.items():
        values[display_name] = sum(net_balance(raw, label, column) for label in labels) / 100_000
    values["合计"] = values["外资"] + values["机构"] + values["散户"]
    return {
        "日期": end_date,
        "区间": period_label,
        "机构": values["机构"],
        "散户": values["散户"],
        "外资": values["外资"],
        "合计": values["合计"],
        "来源": source_url,
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
        return f"连续净买入 {streak} 周"
    if streak < 0:
        return f"连续净卖出 {abs(streak)} 周"
    return "无连续方向"


def build_summary(market_flow: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for investor in ["外资", "机构", "散户"]:
        streak = current_streak(market_flow[investor])
        rows.append(
            {
                "对象": "JPX市场",
                "指标": f"{investor}净买卖金额",
                "最近值": market_flow[investor].iloc[-1] if not market_flow.empty else 0,
                "单位": "亿日元",
                "连续状态": streak_label(streak),
            }
        )
    return pd.DataFrame(rows)


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


def add_cumulative_columns(frame: pd.DataFrame) -> list[dict]:
    if frame.empty:
        return []
    view = frame.copy()
    for investor in ["外资", "机构", "散户"]:
        view[f"{investor}累计"] = view[investor].cumsum()
    return frame_to_records(view)


def build_payload(market_flow: pd.DataFrame, summary: pd.DataFrame, source_urls: list[str]) -> dict:
    latest = market_flow.iloc[-1] if not market_flow.empty else {}
    latest_date = latest["日期"].strftime("%Y-%m-%d") if len(latest) else "-"
    latest_period = latest.get("区间", "-") if len(latest) else "-"
    start = market_flow["日期"].min().strftime("%Y-%m-%d") if not market_flow.empty else "-"
    return {
        "meta": {
            "start": start,
            "end": latest_date,
            "market": "JPX",
            "frequency": "weekly",
            "unit": "亿日元",
            "source": "JPX Trading by Type of Investors",
            "sourceUrl": urljoin(JPX_BASE_URL, INVESTOR_TYPE_PATH),
            "sourceFiles": source_urls,
            "generatedAt": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "latestDate": latest_date,
            "latestPeriodLabel": latest_period,
        },
        "summary": frame_to_records(summary),
        "marketFlow": add_cumulative_columns(market_flow),
        "stockFlows": {},
    }


def write_json(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def save_raw(raw_dir: Path | None, weekly_file: WeeklyFile, content: bytes) -> None:
    if raw_dir is None:
        return
    raw_dir.mkdir(parents=True, exist_ok=True)
    (raw_dir / f"stock_val_1_{weekly_file.code}.xls").write_bytes(content)


def main() -> None:
    args = parse_args()
    output_path = Path(args.output) if args.output else Path(args.site_dir) / "data" / "japan" / "latest.json"
    raw_dir = Path(args.raw_dir) if args.raw_dir else None
    client = JpxClient()

    try:
        weekly_files = client.weekly_value_files(args.weeks, args.archive_pages)
        if not weekly_files:
            raise RuntimeError("未找到 JPX stock_val Excel 下载链接")

        rows = []
        source_urls = []
        for weekly_file in weekly_files:
            content = client.get_bytes(weekly_file.url)
            save_raw(raw_dir, weekly_file, content)
            rows.append(parse_weekly_value_excel(content, weekly_file.url, args.sheet))
            source_urls.append(weekly_file.url)

        market_flow = pd.DataFrame(rows).sort_values("日期").reset_index(drop=True)
        summary = build_summary(market_flow)
        payload = build_payload(market_flow, summary, source_urls)
        write_json(output_path, payload)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print(f"已生成日股网站数据: {output_path}")
    print(f"最新周期: {payload['meta']['latestDate']} ({payload['meta']['latestPeriodLabel']})")
    print(f"数据周数: {len(payload['marketFlow'])}")


if __name__ == "__main__":
    main()
