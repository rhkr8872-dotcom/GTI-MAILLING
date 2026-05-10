# =========================================================
# GTI STEP2 - NAVER ONLY (FINAL MASTER VERSION)
# =========================================================

import requests
import pandas as pd
import re
import html
from datetime import datetime, timedelta

print("🚀 GTI STEP2 NAVER FINAL START")

# =============================
# PATH
# =============================
BASE_PATH = "C:\\temp\\"
KEYWORD_FILE = BASE_PATH + "keyword.xlsx"
RAW_FILE = BASE_PATH + "2-1.naver_news_raw.xlsx"

# =============================
# NAVER API
# =============================
NAVER_CLIENT_ID = "9UlPlXyaXYVEzAdNmvra"
NAVER_CLIENT_SECRET = "MedSQxOnsH"

NAVER_URL = "https://openapi.naver.com/v1/search/news.json"

HEADERS = {
    "X-Naver-Client-Id": NAVER_CLIENT_ID,
    "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
}

# =============================
# TITLE 정규화
# =============================
def normalize_title(title):

    if not title:
        return ""

    t = html.unescape(title)
    t = re.sub(r'<.*?>', '', t)
    t = re.sub(r'\s*\[.*?\]\s*$', '', t)
    t = re.sub(r'\(.*?\)', '', t)
    t = t.replace('"', '').replace("'", "")
    t = re.sub(r'[…·]', ' ', t)
    t = re.sub(r'\s?기자$', '', t)
    t = re.sub(r'\s+', ' ', t)

    return t.strip()

# =============================
# KEYWORD LOAD
# =============================
def load_keywords():
    keywords = pd.read_excel(KEYWORD_FILE)
    return keywords.iloc[:, 0].dropna().tolist()

# =============================
# KEYWORD 포함 여부
# =============================
def has_keyword(title, keyword_list):
    t = title.lower()
    return any(k.lower() in t for k in keyword_list)

# =============================
# 주식 뉴스 제거
# =============================
STOCK_KEYWORDS = [
    "코스피","코스닥","시총","주가","급등","급락","상승","하락",
    "증시","투자","매수","매도","외인","기관","개인","공매도"
]

def is_stock(title):
    return any(k in title for k in STOCK_KEYWORDS)

# =============================
# 정책 키워드 (확장 완료)
# =============================
POLICY_KEYWORDS = [
    "관세","통관","무역","수출","수입","fta",
    "tariff","customs","trade","import","export",

    "협정","경제동반자","epa","fta 협상",
    "301조","ustr","무역법","통상법",
    "공급망","리쇼어링","디커플링",
    "전략물자","수출통제","제재",
    "경제안보","핵심광물","희토류"
]

def is_policy(title):
    t = title.lower()
    return any(k.lower() in t for k in POLICY_KEYWORDS)

# =============================
# PROTECT 키워드 (핵심)
# =============================
PROTECT_KEYWORDS = [
    "수출","수입","원산지","통관",
    "FTA","협정","단속",
    "공급망","전략물자",
    "경제안보","수출통제"
]

def is_protected(title):
    return any(k in title for k in PROTECT_KEYWORDS)

# =============================
# 보완 allow
# =============================
EXTRA_ALLOW = [
    "트럼프","미국","중국","eu","유럽",
    "이란","호르무즈"
]

def is_extra_allowed(title):
    return any(k in title for k in EXTRA_ALLOW)

# =============================
# 제외 필터
# =============================
EXCLUDE_KEYWORDS = [
    "연예","배우","아이돌","결혼","열애",
    "날씨","운세","로또","단상",
    "선거","후보","대통령","국민의힘","민주당"
]

def is_excluded(title):
    return any(k in title for k in EXCLUDE_KEYWORDS)

# =============================
# FINAL FILTER (완성형)
# =============================
def final_filter(title, keyword_list):

    # 1️⃣ 주식 제거 (최우선)
    if is_stock(title):
        return False

    # 2️⃣ 검색 KEYWORD
    if has_keyword(title, keyword_list):
        return True

    # 3️⃣ 보호 키워드
    if is_protected(title):
        return True

    # 4️⃣ 정책 키워드
    if is_policy(title):
        return True

    # 5️⃣ 보완 allow
    if is_extra_allowed(title):
        return True

    # 6️⃣ 잡뉴스 제거
    if is_excluded(title):
        return False

    return False

# =============================
# NAVER 수집
# =============================
def collect_naver(keyword_list):

    results = []

    for kw in keyword_list:
        for start in [1, 51]:

            params = {
                "query": kw,
                "display": 50,
                "start": start,
                "sort": "date"
            }

            res = requests.get(NAVER_URL, headers=HEADERS, params=params)

            if res.status_code != 200:
                continue

            items = res.json().get("items", [])

            for item in items:
                try:
                    title = normalize_title(item.get("title", ""))

                    if len(title) < 10:
                        continue

                    url = item.get("originallink") or item.get("link")
                    source = url

                    pub = item.get("pubDate")

                    try:
                        date = datetime.strptime(pub, "%a, %d %b %Y %H:%M:%S +0900")
                    except:
                        continue

                    results.append({
                        "date": date,
                        "title": title,
                        "url": url,
                        "source": source,
                        "collected_at": datetime.now()
                    })

                except:
                    continue

    print(f"🟢 NAVER collected: {len(results)}")
    return results

# =============================
# MAIN
# =============================
def main():

    keyword_list = load_keywords()
    print(f"🔎 keywords loaded: {len(keyword_list)}")

    data = collect_naver(keyword_list)

    df = pd.DataFrame(data)

    print(f"📊 TOTAL RAW: {len(df)}")

    # 24시간 필터
    df = df[df["date"] >= datetime.now() - timedelta(days=1)]
    print(f"📊 24h FILTER: {len(df)}")

    # FINAL FILTER
    before = len(df)
    df = df[df["title"].apply(lambda x: final_filter(x, keyword_list))]
    print(f"📊 FINAL FILTER 제거: {before - len(df)}")

    # 중복 제거
    df["title_clean"] = df["title"].str.lower().str.strip()

    df = df.drop_duplicates(subset=["url"])
    df = df.drop_duplicates(subset=["title_clean"])

    print(f"📊 DEDUP: {len(df)}")

    # 정렬
    df = df.sort_values(by="date", ascending=False)

    # 저장
    df = df[["date","title","url","source","collected_at"]]

    df.to_excel(RAW_FILE, index=False)

    print("💾 saved:", RAW_FILE)
    print("✅ STEP2 DONE")

# =============================
# 실행
# =============================
if __name__ == "__main__":
    main()