# =========================================================
# GTI RSS Collector FINAL (5-COLUMN VERSION)
# =========================================================
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "output"

OUT_DIR.mkdir(exist_ok=True)

SITE_FILE = DATA_DIR / "site.xlsx"
KEYWORD_FILE = DATA_DIR / "keyword.xlsx"
MAIL_FILE = DATA_DIR / "00.xlsx"

RAW1 = OUT_DIR / "1.site_news_raw.xlsx"
RAW2 = OUT_DIR / "2-1.naver_news_raw.xlsx"
RAW3 = OUT_DIR / "2-2.google_news_raw.xlsx"
RAW4 = OUT_DIR / "2-3.rss_news_raw.xlsx"

MERGED = OUT_DIR / "news_ai_summary.xlsx"
FINAL = OUT_DIR / "news_raw.xlsx"


import os
import re
import pandas as pd
import feedparser
from datetime import datetime, timedelta
from dateutil import parser
from bs4 import BeautifulSoup

# ===================== CONFIG =====================
OUTPUT_FILE = "C:/Temp/2-3.rss_news_raw.xlsx"
os.makedirs("C:/Temp", exist_ok=True)

CUT_OFF = datetime.now() - timedelta(hours=24)

RSS_FEEDS = [
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/17985243432244334502",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/4516205150948857615",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/3080841471748731166",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/10501608952048361545",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/16961830737247159792",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/948091112940518023",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/6142932470653053280",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/6142932470653051343",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/13174877188573923497",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/2076463365346997951",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/4335278970500797203",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/17562799721792586971",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/6604744804871358523",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/3186112205794992580",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/4949128239165847220",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/3855478583015316147",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/10245374560221016947",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/4949128239165849404",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/1786883755329664177",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/6604744804871357744",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/2076463365346996799",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/17038245884356439696",
    "https://www.google.co.kr/alerts/feeds/10136715025889488318/10635279785075251135",
    "https://unipass.customs.go.kr/rss.do"
]

# ===================== UTILS =====================

def parse_date(d):
    try:
        return parser.parse(d).replace(tzinfo=None)
    except:
        return None

def clean_html(text):
    return re.sub("<.*?>", "", str(text)).strip()

def extract_title(entry):
    return clean_html(entry.get("title", ""))

def extract_url(entry):
    # 1️⃣ Alert RSS 원문 URL 시도
    try:
        soup = BeautifulSoup(entry.get("summary", ""), "html.parser")
        for a in soup.find_all("a"):
            href = a.get("href", "")
            if href.startswith("http") and "google" not in href:
                return href
    except:
        pass

    # 2️⃣ fallback
    return entry.get("link", "")

# ===================== COLLECT =====================

def collect():
    rows = []

    for f in RSS_FEEDS:
        print("FETCH:", f)
        feed = feedparser.parse(f)

        for e in feed.entries:

            title = extract_title(e)
            if not title:
                continue

            dt = parse_date(e.get("published", e.get("updated", "")))
            if not dt or dt < CUT_OFF:
                continue

            url = extract_url(e)

            rows.append({
                "date": dt,
                "title": title,
                "url": url,
                "source": f,
                "collected_at": datetime.now().replace(microsecond=0)
            })

    return pd.DataFrame(rows)

# ===================== DEDUP =====================

def dedup(df):
    return df.drop_duplicates(subset=["title"]).copy()

# ===================== MAIN =====================

def main():
    print("🚀 RSS START")

    df = collect()
    print("Collected:", len(df))

    if df.empty:
        print("❌ NO DATA")
        return

    df = dedup(df)
    print("After dedup:", len(df))

    df = df.sort_values("date", ascending=False).head(200)

    # 🔥 컬럼 순서 고정
    df = df[["date", "title", "url", "source", "collected_at"]]

    df.to_excel(OUTPUT_FILE, index=False)

    print("📁 SAVED:", OUTPUT_FILE)
    print("✅ DONE:", len(df))


if __name__ == "__main__":
    main()
