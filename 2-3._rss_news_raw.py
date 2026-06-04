# =========================================================
# GTI STEP2-3 - RSS NEWS RAW FINAL v3.1
# Google Alert Redirect URL 해제 + RSS Feed 오류 수정
# feed_name + keyword/category/importance 추가
# URL 기준 Dedup 안정화
# =========================================================

import os
import re
import pandas as pd
import feedparser
from datetime import datetime, timedelta
from dateutil import parser
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs, unquote

print("🚀 GTI STEP2-3 RSS START v3.1")

# ===================== CONFIG =====================
OUTPUT_FILE = "C:/Temp/2-3.rss_news_raw.xlsx"
os.makedirs("C:/Temp", exist_ok=True)

LOOKBACK_HOURS = 24
CUT_OFF = datetime.now() - timedelta(hours=LOOKBACK_HOURS)

RSS_FEEDS = [
    {
        "feed_name": "Google Alert - Customs",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/17985243432244334502",
        "keyword": "customs",
        "category": "CUSTOMS",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "Google Alert - Tariff",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/4516205150948857615",
        "keyword": "tariff",
        "category": "TARIFF",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "Google Alert - FTA Origin",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/3080841471748731166",
        "keyword": "FTA origin",
        "category": "FTA_ORIGIN",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "Google Alert - Export Control",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/10501608952048361545",
        "keyword": "export control",
        "category": "EXPORT_CONTROL",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "Google Alert - HS Code",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/16961830737247159792",
        "keyword": "HS code",
        "category": "HS_CODE",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "Google Alert - Trade Policy 01",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/948091112940518023",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 02",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/6142932470653053280",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 03",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/6142932470653051343",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 04",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/13174877188573923497",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 05",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/2076463365346997951",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 06",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/4335278970500797203",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 07",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/17562799721792586971",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 08",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/6604744804871358523",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 09",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/3186112205794992580",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 10",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/4949128239165847220",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 11",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/3855478583015316147",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 12",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/10245374560221016947",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 13",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/4949128239165849404",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 14",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/1786883755329664177",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 15",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/6604744804871357744",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 16",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/2076463365346996799",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 17",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/17038245884356439696",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Google Alert - Trade Policy 18",
        "url": "https://www.google.co.kr/alerts/feeds/10136715025889488318/10635279785075251135",
        "keyword": "trade policy",
        "category": "TRADE_POLICY",
        "importance": "MEDIUM",
        "importance_score": 70,
    },
    {
        "feed_name": "Korea Customs UNIPASS",
        "url": "https://unipass.customs.go.kr/rss.do",
        "keyword": "Korea customs",
        "category": "CUSTOMS_REGULATION",
        "importance": "HIGH",
        "importance_score": 100,
    },
    {
        "feed_name": "EU TAXUD",
        "url": "https://taxation-customs.ec.europa.eu/node/2/rss_en",
        "keyword": "EU customs taxation",
        "category": "CUSTOMS_REGULATION",
        "importance": "HIGH",
        "importance_score": 100,
    },
]

FINAL_COLS = [
    "date", "title", "url", "source", "feed_name",
    "summary", "collected_at",
    "keyword", "category", "importance", "importance_score"
]

# ===================== UTILS =====================

def parse_date(d):
    try:
        return parser.parse(str(d)).replace(tzinfo=None)
    except Exception:
        return None


def clean_html(text):
    if text is None:
        return ""
    soup = BeautifulSoup(str(text), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def normalize_url(url):
    if not url:
        return ""

    url = str(url).strip()
    url = re.sub(r"#.*$", "", url)
    return url


def decode_google_redirect(url):
    """
    Google Alert / Google News redirect URL을 실제 원문 URL로 변환.
    지원 형태:
    - https://www.google.com/url?...&url=https://original...
    - https://www.google.com/url?...&q=https://original...
    - 기타 URL은 원문 그대로 반환
    """
    if not url:
        return ""

    url = str(url).strip()

    try:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)

        for key in ["url", "q"]:
            if key in qs and qs[key]:
                candidate = unquote(qs[key][0]).strip()
                if candidate.startswith("http"):
                    return normalize_url(candidate)

    except Exception:
        pass

    return normalize_url(url)


def extract_title(entry):
    return clean_html(entry.get("title", ""))


def extract_summary(entry):
    return clean_html(entry.get("summary", entry.get("description", "")))


def extract_url(entry):
    # 1. Google Alert summary 내부 링크 우선 사용
    try:
        soup = BeautifulSoup(entry.get("summary", ""), "html.parser")
        for a in soup.find_all("a"):
            href = a.get("href", "")
            if href.startswith("http"):
                return decode_google_redirect(href)
    except Exception:
        pass

    # 2. fallback: RSS entry link 사용
    return decode_google_redirect(entry.get("link", ""))

# ===================== COLLECT =====================

def collect():
    rows = []

    for feed_info in RSS_FEEDS:
        feed_name = feed_info["feed_name"]
        feed_url = feed_info["url"]

        print("FETCH:", feed_name, "|", feed_url)

        feed = feedparser.parse(feed_url)

        for e in feed.entries:
            title = extract_title(e)
            if not title:
                continue

            dt = parse_date(e.get("published", e.get("updated", "")))
            if not dt or dt < CUT_OFF:
                continue

            url = extract_url(e)
            summary = extract_summary(e)

            rows.append({
                "date": dt,
                "title": title,
                "url": url,
                "source": feed_url,
                "feed_name": feed_name,
                "summary": summary,
                "collected_at": datetime.now().replace(microsecond=0),
                "keyword": feed_info.get("keyword", ""),
                "category": feed_info.get("category", ""),
                "importance": feed_info.get("importance", ""),
                "importance_score": feed_info.get("importance_score", 50),
            })

    return pd.DataFrame(rows)

# ===================== DEDUP =====================

def dedup(df):
    before = len(df)

    df["url"] = df["url"].apply(normalize_url)
    df["url_key"] = df["url"].astype(str).str.strip().str.lower()
    df["title_key"] = df["title"].astype(str).str.strip().str.lower()

    df = df.sort_values(["importance_score", "date"], ascending=[False, False])

    # URL 있는 건은 URL 기준 중복 제거
    df_url = df[df["url_key"] != ""].drop_duplicates(subset=["url_key"], keep="first")

    # URL 없는 건만 title 기준 중복 제거
    df_no_url = df[df["url_key"] == ""].drop_duplicates(subset=["title_key"], keep="first")

    out = pd.concat([df_url, df_no_url], ignore_index=True)
    out = out.drop(columns=["url_key", "title_key"], errors="ignore")

    print(f"📊 DEDUP URL/TITLE: {before} -> {len(out)}")
    return out

# ===================== MAIN =====================

def main():
    df = collect()
    print("📊 Collected:", len(df))

    if df.empty:
        print("❌ NO DATA")
        pd.DataFrame(columns=FINAL_COLS).to_excel(OUTPUT_FILE, index=False)
        print("💾 saved empty file:", OUTPUT_FILE)
        return

    df = dedup(df)
    df = df.sort_values(["importance_score", "date"], ascending=[False, False]).head(300)
    df = df[FINAL_COLS]

    df.to_excel(OUTPUT_FILE, index=False)

    print("📁 SAVED:", OUTPUT_FILE)
    print("✅ DONE:", len(df))


if __name__ == "__main__":
    main()
