#!/usr/bin/env python3
"""扫描阿拉丁与爱择公开列表，生成铝信息速递候选审计文件。

只读取公开标题、日期和链接，不抓取或复述付费正文；不自动写入 news.json，
避免把未经核验的标题直接升级为研究结论。
"""

from __future__ import annotations

import html
import json
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urljoin

import requests


ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "aluminum_source_scan.json"
SHANGHAI = timezone(timedelta(hours=8))
LOOKBACK_DAYS = 45

ALADDINY_URL = "https://www.aladdiny.com/WebService/WebData.ashx?act=kindcol"
ALADDINY_KINDS = {
    "0103": "铝土矿",
    "0101": "氧化铝",
    "0102": "电解铝",
}
AZCHINA_CLASSES = {
    "37": "电解铝专题调研",
    "287": "铝土矿氧化铝专题调研",
    "54": "电解铝独家调研",
    "290": "铝土矿氧化铝独家调研",
    "50": "电解铝产能产量",
    "52": "电解铝库存",
}
HIGH_SIGNAL = re.compile(
    r"减产|停产|复产|投产|通电|关停|检修|事故|发运|配额|出口|产能|产量|矿山|断供|恢复|袭击",
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/127 Safari/537.36",
}


def clean_text(value: str) -> str:
    value = html.unescape(re.sub(r"<[^>]+>", " ", value or ""))
    return re.sub(r"\s+", " ", value).strip()


def scan_aladdiny(session: requests.Session) -> list[dict]:
    headers = {
        **HEADERS,
        "Referer": "https://www.aladdiny.com/",
        "X-Requested-With": "XMLHttpRequest",
    }
    rows: list[dict] = []
    for code, product in ALADDINY_KINDS.items():
        response = session.post(
            ALADDINY_URL,
            headers=headers,
            data={"kindcode": code, "colcode": "", "keyword": "", "pageSize": "50", "pageIndex": "1"},
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        for item in payload.get("list") or []:
            info_code = str(item.get("infoCode") or "")
            raw_day = info_code[:8]
            day = f"{raw_day[:4]}-{raw_day[4:6]}-{raw_day[6:8]}" if re.fullmatch(r"20\d{6}", raw_day) else None
            title = clean_text(str(item.get("title") or ""))
            rows.append({
                "source": "阿拉丁(ALD)",
                "channel": product,
                "date": day,
                "title": title,
                "kind": item.get("kindName"),
                "url": f"https://www.aladdiny.com/news/article/{info_code}.html",
                "high_signal": bool(HIGH_SIGNAL.search(title)),
            })
    return rows


def article_date(session: requests.Session, url: str) -> str | None:
    try:
        response = session.get(url, headers={**HEADERS, "Referer": "https://www.azchina-cn.com/"}, timeout=25)
        response.raise_for_status()
        matches = re.findall(r"20\d{2}[-./年]\d{1,2}[-./月]\d{1,2}", response.text)
        if not matches:
            return None
        parts = re.findall(r"\d+", matches[0])
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}-{int(parts[2]):02d}"
    except requests.RequestException:
        return None


def scan_azchina(session: requests.Session) -> list[dict]:
    base = "https://www.azchina-cn.com/"
    found: dict[str, dict] = {}
    pattern = re.compile(
        r"href=[\']([^\']*ShowInfo\.php\?[^\']+)[\'][^>]*>(.*?)</a>",
        re.I | re.S,
    )
    for class_id, channel in AZCHINA_CLASSES.items():
        list_url = urljoin(base, f"e/action/ListInfo/?classid={class_id}")
        response = session.get(list_url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        for link, raw_title in pattern.findall(response.text):
            title = clean_text(raw_title)
            if len(title) < 5:
                continue
            url = urljoin(base, link)
            row = found.setdefault(url, {
                "source": "爱择咨询",
                "channel": channel,
                "date": None,
                "title": title,
                "url": url,
                "high_signal": bool(HIGH_SIGNAL.search(title)),
            })
            if channel not in row["channel"]:
                row["channel"] += "/" + channel
    rows = list(found.values())
    for row in rows[:30]:
        row["date"] = article_date(session, row["url"])
    return rows


def main() -> None:
    cutoff = (date.today() - timedelta(days=LOOKBACK_DAYS)).isoformat()
    errors: dict[str, str] = {}
    rows: list[dict] = []
    with requests.Session() as session:
        for name, scanner in (("aladdiny", scan_aladdiny), ("azchina", scan_azchina)):
            try:
                rows.extend(scanner(session))
            except Exception as error:  # noqa: BLE001
                errors[name] = f"{type(error).__name__}: {error}"[:300]

    dedup: dict[str, dict] = {}
    for row in rows:
        if row.get("date") and row["date"] < cutoff:
            continue
        dedup[row["url"]] = row
    ordered = sorted(
        dedup.values(),
        key=lambda row: (row.get("date") or "", bool(row.get("high_signal")), row.get("title") or ""),
        reverse=True,
    )
    payload = {
        "fetched_at": datetime.now(SHANGHAI).isoformat(timespec="seconds"),
        "lookback_days": LOOKBACK_DAYS,
        "method": "公开列表标题扫描；付费正文不抓取，入 news.json 前必须交叉核验",
        "errors": errors,
        "items": ordered,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "output": str(OUT),
        "items": len(ordered),
        "high_signal": sum(bool(row.get("high_signal")) for row in ordered),
        "latest": [row for row in ordered if row.get("date")][:5],
        "errors": errors,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
