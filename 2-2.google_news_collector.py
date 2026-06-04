# =========================================================
# GTI STEP2-2 - GOOGLE NEWS RAW FAST FINAL v3.2
# 목적: 빠른 수집 + metadata 확보
# 원문 URL 변환은 STEP3 Cluster Merge 이후 Top 기사에 대해서만 수행
# =========================================================

import re
import time
import pandas as pd
import feedparser
from datetime import datetime, timedelta
from urllib.parse import quote
from bs4 import BeautifulSoup

print("🚀 GTI STEP2-2 GOOGLE NEWS FAST START v3.2")

# =============================
# PATH / CONFIG
# =============================
BASE_PATH = "C:\\temp\\"
KEYWORD_FILE = BASE_PATH + "keyword.xlsx"
RAW_FILE = BASE_PATH + "2-2.google_news_raw.xlsx"

LOOKBACK_HOURS = 24
SLEEP_SEC = 0.05

FINAL_COLS = [
    "date",
    "title",
    "url",
    "google_url",
    "source",
    "summary",
    "collected_at",
    "keyword",
    "language",
    "publisher",
    "category",
    "importance",
    "importance_score",
    "url_decode_status",
]

# =============================
# UTILS
# =============================

def clean_html(text):
    if pd.isna(text):
        return ""
    soup = BeautifulSoup(str(text), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def parse_datetime(entry):
    for key in ["published_parsed", "updated_parsed"]:
        try:
            v = getattr(entry, key, None)
            if v:
                return datetime(*v[:6])
        except Exception:
            pass
    return datetime.now()


def is_recent(dt):
    return dt >= datetime.now() - timedelta(hours=LOOKBACK_HOURS)


def normalize_title(title):
    title = clean_html(title).lower()

    # Google News title usually: "Article title - Publisher"
    # dedup 목적상 뒤 publisher 제거
    if " - " in title:
        title = title.rsplit(" - ", 1)[0]

    title = re.sub(r"[^0-9a-z가-힣一-龥ぁ-ゔァ-ヴー\s]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def get_google_locale(language):
    lang = str(language).upper().strip()
    locale_map = {
        "EN": {"hl": "en", "gl": "US", "ceid": "US:en"},
        "KR": {"hl": "ko", "gl": "KR", "ceid": "KR:ko"},
        "CN": {"hl": "zh-CN", "gl": "CN", "ceid": "CN:zh-Hans"},
        "ES": {"hl": "es", "gl": "ES", "ceid": "ES:es"},
        "PT": {"hl": "pt", "gl": "BR", "ceid": "BR:pt-419"},
        "TR": {"hl": "tr", "gl": "TR", "ceid": "TR:tr"},
        "VI": {"hl": "vi", "gl": "VN", "ceid": "VN:vi"},
        "HI": {"hl": "hi", "gl": "IN", "ceid": "IN:hi"},
    }
    return locale_map.get(lang, locale_map["EN"])


def importance_to_score(v):
    s = str(v).strip().upper()
    if s in ["100", "HIGH", "H", "상", "중요", "A"]:
        return 100
    if s in ["70", "80", "MEDIUM", "M", "중", "B"]:
        return 70
    if s in ["50", "LOW", "L", "하", "C"]:
        return 50
    try:
        return int(float(s))
    except Exception:
        return 50


def extract_publisher(entry, title):
    try:
        src = entry.get("source", {})
        if isinstance(src, dict):
            src_title = src.get("title", "")
            if src_title:
                return clean_html(src_title)
    except Exception:
        pass

    try:
        if " - " in str(title):
            return str(title).rsplit(" - ", 1)[-1].strip()
    except Exception:
        pass

    return ""

# =============================
# KEYWORD LOAD
# =============================

def load_keywords():
    keywords = pd.read_excel(KEYWORD_FILE)
    keywords.columns = [str(c).strip().lower() for c in keywords.columns]

    required_cols = ["keyword", "language", "category", "importance", "active"]
    for col in required_cols:
        if col not in keywords.columns:
            raise Exception(f"❌ KEYWORD 파일 필수 컬럼 없음: {col}")

    keywords = keywords[keywords["active"].astype(str).str.upper().str.strip() == "Y"]
    keywords = keywords.dropna(subset=["keyword"])
    keywords["keyword"] = keywords["keyword"].astype(str).str.strip()
    keywords = keywords[keywords["keyword"] != ""].copy()

    return keywords

# =============================
# COLLECT
# =============================

def collect_google_rss(keywords):
    rows = []

    for _, row in keywords.iterrows():
        kw = str(row.get("keyword", "")).strip()
        lang = str(row.get("language", "EN")).strip().upper()
        category = str(row.get("category", "")).strip()
        importance = row.get("importance", "")
        importance_score = importance_to_score(importance)

        locale = get_google_locale(lang)
        query = quote(kw)

        rss_url = (
            "https://news.google.com/rss/search?"
            f"q={query}"
            f"&hl={locale['hl']}"
            f"&gl={locale['gl']}"
            f"&ceid={locale['ceid']}"
        )

        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            dt = parse_datetime(entry)
            if not is_recent(dt):
                continue

            title = clean_html(entry.get("title", ""))
            if not title:
                continue

            google_url = entry.get("link", "")
            summary = clean_html(entry.get("summary", ""))
            publisher = extract_publisher(entry, title)

            # FAST 원칙: STEP2에서는 Google URL 유지
            # 실제 원문 URL은 STEP3 Cluster Merge 후 Top 기사만 decode
            rows.append({
                "date": dt,
                "title": title,
                "url": google_url,
                "google_url": google_url,
                "source": "Google News RSS",
                "summary": summary,
                "collected_at": datetime.now().replace(microsecond=0),
                "keyword": kw,
                "language": lang,
                "publisher": publisher,
                "category": category,
                "importance": importance,
                "importance_score": importance_score,
                "url_decode_status": "SKIPPED_STEP2_FAST",
            })

        time.sleep(SLEEP_SEC)

    return pd.DataFrame(rows)

# =============================
# DEDUP
# =============================

def dedup_fast(df):
    before = len(df)

    df["google_url_key"] = df["google_url"].astype(str).str.strip().str.lower()
    df["title_key"] = df["title"].apply(normalize_title)

    df = df.sort_values(["importance_score", "date"], ascending=[False, False])

    # 1차: Google URL 기준 정확 중복 제거
    df = df.drop_duplicates(subset=["google_url_key"], keep="first")

    # 2차: title_key 기준 중복 제거
    # 동일 기사가 keyword만 다르게 여러 번 잡히는 경우 제거
    before_title = len(df)
    df = df.drop_duplicates(subset=["title_key"], keep="first")

    print(f"📊 DEDUP GOOGLE_URL/TITLE: {before} -> {len(df)}")
    print(f"   - title dedup effect: {before_title} -> {len(df)}")

    df = df.drop(columns=["google_url_key", "title_key"], errors="ignore")
    return df

# =============================
# MAIN
# =============================

def main():
    keywords = load_keywords()
    print(f"🔎 active keywords: {len(keywords)}")

    df = collect_google_rss(keywords)

    if df.empty:
        print("❌ No data collected")
        pd.DataFrame(columns=FINAL_COLS).to_excel(RAW_FILE, index=False)
        print("💾 saved empty file:", RAW_FILE)
        return

    print(f"📊 TOTAL RAW: {len(df)}")

    before = len(df)
    df = df[df["date"].apply(is_recent)].copy()
    print(f"📊 24h FILTER: {before} -> {len(df)}")

    df = dedup_fast(df)
    df = df.sort_values(["importance_score", "date"], ascending=[False, False])
    df = df[FINAL_COLS]

    df.to_excel(RAW_FILE, index=False)

    print("💾 saved:", RAW_FILE)
    print(f"✅ FINAL SAVE ROWS: {len(df)}")
    print("✅ STEP2-2 GOOGLE NEWS FAST DONE v3.2")


if __name__ == "__main__":
    main()
