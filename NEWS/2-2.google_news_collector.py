# =========================================================
# GTI STEP2 - GOOGLE NEWS RAW (RSS ONLY FINAL)
# =========================================================

import pandas as pd
import re
from datetime import datetime, timedelta
import feedparser
from urllib.parse import quote

print("🚀 GTI STEP2 GOOGLE RSS START")

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
    return re.sub('<.*?>', '', str(text))

# =============================
# 24H FILTER
# =============================
def is_recent(dt):
    return dt >= datetime.now() - timedelta(days=1)

# =============================
# KEYWORD LOAD
# =============================
keywords = pd.read_excel(KEYWORD_FILE)
keyword_list = keywords.iloc[:,0].dropna().tolist()

print(f"🔎 keywords: {len(keyword_list)}")

# =============================
# GOOGLE RSS COLLECT
# =============================
def collect_google_rss():
    results = []

    for kw in keyword_list:
        query = quote(kw)
        rss_url = f"https://news.google.com/rss/search?q={query}"

        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            try:
                dt = datetime(*entry.published_parsed[:6])
            except:
                dt = datetime.now()

            results.append({
                "date": dt,
                "title": clean_html(entry.title),
                "url": entry.link,          # Google 뉴스 링크 (허용)
                "source": entry.link,       # 동일 처리
                "collected_at": datetime.now()
            })

    print(f"🟢 GOOGLE RSS collected: {len(results)}")
    return results

# =============================
# MAIN
# =============================
data = collect_google_rss()
df = pd.DataFrame(data)

print(f"📊 TOTAL RAW: {len(df)}")

# 24시간 필터
df = df[df['date'].apply(is_recent)]
print(f"📊 24h FILTER: {len(df)}")

# URL 중복 제거
df = df.drop_duplicates(subset=['url'])

# 컬럼 순서 정리
df = df[['date','title','url','source','collected_at']]

# 저장
df.to_excel(RAW_FILE, index=False)

print("💾 saved:", RAW_FILE)
print("✅ GOOGLE RSS DONE")