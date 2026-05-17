# =========================================================
# GTI STEP2 - GOOGLE NEWS RAW (RSS ONLY FINAL v2.0)
# 기존 Form 유지 + Keyword Metadata 추가
# =========================================================

import pandas as pd
import re
from datetime import datetime, timedelta
import feedparser
from urllib.parse import quote

print("🚀 GTI STEP2 GOOGLE RSS START v2.0")

# =============================
# PATH
# =============================
BASE_PATH = "C:\\temp\\"
KEYWORD_FILE = BASE_PATH + "keyword.xlsx"
RAW_FILE = BASE_PATH + "2-2.google_news_raw.xlsx"

# =============================
# CLEAN
# =============================
def clean_html(text):
    return re.sub("<.*?>", "", str(text)).strip()

def is_recent(dt):
    return dt >= datetime.now() - timedelta(days=1)

def extract_publisher(title):
    """
    Google News RSS title 예:
    'Police probe major customs data breach - The Times of India'
    """
    try:
        if " - " in title:
            return title.split(" - ")[-1].strip()
        return ""
    except:
        return ""

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

# =============================
# KEYWORD LOAD
# =============================
keywords = pd.read_excel(KEYWORD_FILE)

# 컬럼명 정리
keywords.columns = [str(c).strip().lower() for c in keywords.columns]

required_cols = ["keyword", "language", "category", "importance", "active"]
for col in required_cols:
    if col not in keywords.columns:
        raise Exception(f"❌ KEYWORD 파일 필수 컬럼 없음: {col}")

# active = Y 만 사용
keywords = keywords[keywords["active"].astype(str).str.upper().str.strip() == "Y"]

# keyword 공백 제거
keywords = keywords.dropna(subset=["keyword"])
keywords["keyword"] = keywords["keyword"].astype(str).str.strip()
keywords = keywords[keywords["keyword"] != ""]

print(f"🔎 active keywords: {len(keywords)}")

# =============================
# GOOGLE RSS COLLECT
# =============================
def collect_google_rss():
    results = []

    for _, row in keywords.iterrows():
        kw = str(row.get("keyword", "")).strip()
        lang = str(row.get("language", "EN")).strip().upper()
        category = str(row.get("category", "")).strip()
        importance = str(row.get("importance", "")).strip()

        locale = get_google_locale(lang)

        query = quote(kw)
        rss_url = (
            f"https://news.google.com/rss/search?"
            f"q={query}"
            f"&hl={locale['hl']}"
            f"&gl={locale['gl']}"
            f"&ceid={locale['ceid']}"
        )

        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            try:
                dt = datetime(*entry.published_parsed[:6])
            except:
                dt = datetime.now()

            title = clean_html(entry.title)
            publisher = extract_publisher(title)

            results.append({
                "date": dt,
                "title": title,
                "url": entry.link,
                "source": entry.link,
                "collected_at": datetime.now(),

                # 추가 컬럼
                "keyword": kw,
                "language": lang,
                "publisher": publisher,
                "category": category,
                "importance": importance
            })

    print(f"🟢 GOOGLE RSS collected: {len(results)}")
    return results

# =============================
# MAIN
# =============================
data = collect_google_rss()
df = pd.DataFrame(data)

if df.empty:
    print("❌ No data collected")
    df = pd.DataFrame(columns=[
        "date", "title", "url", "source", "collected_at",
        "keyword", "language", "publisher", "category", "importance"
    ])
    df.to_excel(RAW_FILE, index=False)
    print("💾 saved empty file:", RAW_FILE)
    raise SystemExit

print(f"📊 TOTAL RAW: {len(df)}")

# 24시간 필터
before_24h = len(df)
df = df[df["date"].apply(is_recent)]
print(f"📊 24h FILTER: {before_24h} → {len(df)}")

# URL 중복 제거
before_dedup = len(df)
df = df.drop_duplicates(subset=["url"])
print(f"📊 URL DEDUP: {before_dedup} → {len(df)}")

# 최종 컬럼 순서
final_cols = [
    "date",
    "title",
    "url",
    "source",
    "collected_at",
    "keyword",
    "language",
    "publisher",
    "category",
    "importance"
]

df = df[final_cols]

# 저장
df.to_excel(RAW_FILE, index=False)

print("💾 saved:", RAW_FILE)
print(f"✅ FINAL SAVE ROWS: {len(df)}")
print("✅ GOOGLE RSS DONE v2.0")