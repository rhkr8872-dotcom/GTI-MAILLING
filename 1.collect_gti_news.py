# -*- coding: utf-8 -*-
"""
GTI STEP1-3 Collector

Runs in GitHub Actions or locally. It collects trade-policy news from:
  - Naver News API
  - Google News RSS
  - User-managed RSS / government site URLs in sites.xlsx
  - NewsAPI and SerpAPI when keys are available

Output files:
  - 3.news_master_raw.xlsx
  - 3.news_ai_summary.xlsx
  - 3.news_ai_cumulative.xlsx
"""

from __future__ import annotations

import os
import re
import time
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import quote_plus, urlparse

import feedparser
import pandas as pd
import requests
from bs4 import BeautifulSoup


BASE_DIR = Path(os.getenv("GTI_BASE_DIR", "data")).resolve()
BASE_DIR.mkdir(parents=True, exist_ok=True)

RUN_DATE = os.getenv("GTI_RUN_DATE", datetime.now().strftime("%Y-%m-%d"))
LOOKBACK_HOURS = int(os.getenv("GTI_LOOKBACK_HOURS", "24"))
MAX_PER_SOURCE = int(os.getenv("GTI_MAX_PER_SOURCE", "50"))

NAVER_CLIENT_ID = os.getenv("NAVER_CLIENT_ID", "").strip()
NAVER_CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET", "").strip()
NEWS_KEY = os.getenv("NEWS_KEY", "").strip()
NEWS_API = os.getenv("NEWS_API", "https://newsapi.org/v2/everything").strip()
SERPAPI_KEY = os.getenv("SERPAPI_KEY", "").strip()

KEYWORD_FILE = BASE_DIR / "keyword.xlsx"
SITES_FILE = BASE_DIR / "sites.xlsx"
RAW_OUT = BASE_DIR / "3.news_master_raw.xlsx"
SUMMARY_OUT = BASE_DIR / "3.news_ai_summary.xlsx"
CUMULATIVE_OUT = BASE_DIR / "3.news_ai_cumulative.xlsx"

POLICY_HINTS = [
    "tariff", "customs", "trade", "fta", "import", "export", "origin", "anti-dumping",
    "countervailing", "safeguard", "section 232", "section 301", "cbam", "export control",
    "관세", "통관", "무역", "통상", "수입", "수출", "원산지", "덤핑", "상계관세", "세이프가드",
]


def clean(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = BeautifulSoup(str(value), "html.parser").get_text(" ")
    return re.sub(r"\s+", " ", text).strip()


def parse_date(value) -> datetime | None:
    if not value:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    try:
        dt = parsedate_to_datetime(str(value))
        if dt.tzinfo:
            dt = dt.astimezone(timezone.utc).replace(tzinfo=None)
        return dt
    except Exception:
        pass
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.to_pydatetime().replace(tzinfo=None)
    except Exception:
        return None


def load_keywords() -> list[str]:
    defaults = [
        "tariff electronics",
        "customs clearance electronics",
        "FTA origin rules",
        "export control semiconductor",
        "관세 통상 전자",
        "통관 원산지 반도체",
    ]
    if not KEYWORD_FILE.exists():
        return defaults
    df = pd.read_excel(KEYWORD_FILE)
    values = []
    for item in df.astype(str).fillna("").values.ravel().tolist():
        item = clean(item)
        if item and item.lower() not in {"nan", "none"}:
            values.append(item)
    return list(dict.fromkeys(values))[:80] or defaults


def load_sites() -> list[dict]:
    if not SITES_FILE.exists():
        return []
    df = pd.read_excel(SITES_FILE)
    sites = []
    for _, row in df.iterrows():
        values = [clean(v) for v in row.tolist()]
        urls = [v for v in values if v.startswith("http")]
        if not urls:
            continue
        name = next((v for v in values if v and not v.startswith("http")), urlparse(urls[0]).netloc)
        sites.append({"name": name, "url": urls[0]})
    return sites


def score_row(title: str, source: str) -> int:
    text = f"{title} {source}".lower()
    score = sum(2 for hint in POLICY_HINTS if hint.lower() in text)
    if any(x in text for x in ["samsung", "semiconductor", "electronics", "smartphone", "appliance", "반도체", "전자", "스마트폰", "가전"]):
        score += 5
    if any(x in text for x in ["ustr", "cbp", "customs.go.kr", "taxud", "mofcom", "gacc", "federal register"]):
        score += 4
    return score


def make_row(date, title, url, source, agency="", source_file="", keyword="", publisher="", category="") -> dict:
    title = clean(title)
    url = clean(url)
    source = clean(source)
    dt = parse_date(date) or datetime.now()
    score = score_row(title, source)
    importance = "HIGH" if score >= 8 else "MID" if score >= 4 else ""
    return {
        "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "title": title,
        "url": url,
        "source": source,
        "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "agency": agency,
        "source_file": source_file,
        "keyword": keyword,
        "language": "",
        "publisher": publisher,
        "category": category,
        "importance": importance,
        "title_norm": normalize_title(title),
        "score": score,
    }


def normalize_title(title: str) -> str:
    text = clean(title).lower()
    text = re.sub(r"[-–|].*$", "", text)
    text = re.sub(r"[^0-9a-z가-힣]+", " ", text)
    return " ".join(text.split())


def collect_google_news(keywords: list[str]) -> list[dict]:
    rows = []
    for keyword in keywords:
        url = f"https://news.google.com/rss/search?q={quote_plus(keyword)}%20when:1d&hl=en-US&gl=US&ceid=US:en"
        feed = feedparser.parse(url)
        for entry in feed.entries[:MAX_PER_SOURCE]:
            rows.append(make_row(entry.get("published"), entry.get("title"), entry.get("link"), url, source_file="google_news_rss", keyword=keyword))
        time.sleep(0.2)
    return rows


def collect_naver_news(keywords: list[str]) -> list[dict]:
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        return []
    headers = {"X-Naver-Client-Id": NAVER_CLIENT_ID, "X-Naver-Client-Secret": NAVER_CLIENT_SECRET}
    rows = []
    for keyword in keywords:
        params = {"query": keyword, "display": min(MAX_PER_SOURCE, 100), "sort": "date"}
        try:
            res = requests.get("https://openapi.naver.com/v1/search/news.json", headers=headers, params=params, timeout=20)
            res.raise_for_status()
            for item in res.json().get("items", []):
                rows.append(make_row(item.get("pubDate"), item.get("title"), item.get("originallink") or item.get("link"), item.get("link"), source_file="naver_news_api", keyword=keyword))
        except Exception as exc:
            print(f"[WARN] Naver skipped keyword={keyword}: {exc}")
        time.sleep(0.2)
    return rows


def collect_newsapi(keywords: list[str]) -> list[dict]:
    if not NEWS_KEY:
        return []
    rows = []
    from_dt = (datetime.now() - timedelta(hours=LOOKBACK_HOURS)).strftime("%Y-%m-%dT%H:%M:%S")
    for keyword in keywords[:20]:
        params = {"q": keyword, "from": from_dt, "sortBy": "publishedAt", "language": "en", "pageSize": min(MAX_PER_SOURCE, 100), "apiKey": NEWS_KEY}
        try:
            res = requests.get(NEWS_API, params=params, timeout=20)
            res.raise_for_status()
            for article in res.json().get("articles", []):
                source = article.get("source") or {}
                rows.append(make_row(article.get("publishedAt"), article.get("title"), article.get("url"), source.get("name", ""), source_file="newsapi", keyword=keyword, publisher=source.get("name", "")))
        except Exception as exc:
            print(f"[WARN] NewsAPI skipped keyword={keyword}: {exc}")
        time.sleep(0.2)
    return rows


def collect_serpapi(keywords: list[str]) -> list[dict]:
    if not SERPAPI_KEY:
        return []
    rows = []
    for keyword in keywords[:20]:
        params = {"engine": "google_news", "q": keyword, "api_key": SERPAPI_KEY}
        try:
            res = requests.get("https://serpapi.com/search.json", params=params, timeout=20)
            res.raise_for_status()
            for item in res.json().get("news_results", [])[:MAX_PER_SOURCE]:
                rows.append(make_row(item.get("date"), item.get("title"), item.get("link"), item.get("source", ""), source_file="serpapi_google_news", keyword=keyword))
        except Exception as exc:
            print(f"[WARN] SerpAPI skipped keyword={keyword}: {exc}")
        time.sleep(0.2)
    return rows


def collect_site_feeds(sites: list[dict]) -> list[dict]:
    rows = []
    for site in sites:
        url = site["url"]
        feed = feedparser.parse(url)
        if not feed.entries:
            continue
        for entry in feed.entries[:MAX_PER_SOURCE]:
            rows.append(make_row(entry.get("published") or entry.get("updated"), entry.get("title"), entry.get("link"), url, agency=site["name"], source_file="site_or_rss", keyword=site["name"]))
        time.sleep(0.2)
    return rows


def filter_recent_and_relevant(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df["date_dt"] = pd.to_datetime(df["date"], errors="coerce")
    cutoff = datetime.now() - timedelta(hours=LOOKBACK_HOURS)
    df = df[df["date_dt"].isna() | (df["date_dt"] >= cutoff)].copy()
    df = df[df["title"].map(lambda x: any(h.lower() in str(x).lower() for h in POLICY_HINTS)) | (df["score"] >= 4)].copy()
    df = df.drop_duplicates(subset=["url"], keep="first")
    df = df.drop_duplicates(subset=["title_norm"], keep="first")
    return df.drop(columns=["date_dt"], errors="ignore").sort_values(["score", "date"], ascending=[False, False])


def main() -> None:
    keywords = load_keywords()
    sites = load_sites()
    print(f"[COLLECT] keywords={len(keywords)} sites={len(sites)}")
    rows = []
    rows.extend(collect_site_feeds(sites))
    rows.extend(collect_google_news(keywords))
    rows.extend(collect_naver_news(keywords))
    rows.extend(collect_newsapi(keywords))
    rows.extend(collect_serpapi(keywords))

    df = pd.DataFrame(rows)
    df = filter_recent_and_relevant(df)
    df.to_excel(RAW_OUT, index=False)
    df.head(250).to_excel(SUMMARY_OUT, index=False)

    cumulative = df.copy()
    if CUMULATIVE_OUT.exists():
        old = pd.read_excel(CUMULATIVE_OUT)
        cumulative = pd.concat([old, cumulative], ignore_index=True)
    cumulative = cumulative.drop_duplicates(subset=["url"], keep="last")
    cumulative.to_excel(CUMULATIVE_OUT, index=False)

    print(f"[SAVE] raw={RAW_OUT} rows={len(df)}")
    print(f"[SAVE] summary={SUMMARY_OUT} rows={min(len(df), 250)}")
    print(f"[SAVE] cumulative={CUMULATIVE_OUT} rows={len(cumulative)}")


if __name__ == "__main__":
    main()
