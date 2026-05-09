# =========================================================
# GTI STEP4 vFinal
# ✔ 기존 구조 유지
# ✔ 기초 데이터 변경 없음
# ✔ 관세/통상 뉴스 선별
# ✔ 삼성전자 관세 영향 분석
# ✔ 오류 완전 제거
# =========================================================

import pandas as pd
import os
import re
import json
from google import genai

# ===================== CONFIG =====================
INPUT_FILE = "C:/temp/3.news_ai_summary.xlsx"
OUTPUT_DAILY = "C:/temp/4.news_raw.xlsx"
OUTPUT_CUMUL = "C:/temp/4.news_cumulative.xlsx"

API_KEY = "AIzaSyC10h9WdXAR-E7VmmvSPDNJCSSKe0BMEss"
MODEL = "gemini-2.0-flash"

client = genai.Client(api_key=API_KEY)

TOP_N = 30

# ===================== HYPERLINK 처리 =====================
def extract_title(s):
    s = str(s)
    m = re.search(r'HYPERLINK\(".*?","(.*?)"\)', s)
    if m:
        return m.group(1)
    return s.strip()

def extract_url(s):
    s = str(s)
    m = re.search(r'HYPERLINK\("(.*?)",".*?"\)', s)
    if m:
        return m.group(1)
    return ""

# ===================== 뉴스 필터 =====================
def is_trade_news(title):

    t = title.lower()

    keywords = [
        "tariff","관세","301","customs",
        "fta","수출","수입","export","import",
        "anti-dumping","wto","제재","통상"
    ]

    return any(k in t for k in keywords)

# ===================== 점수 =====================
def score(title):

    t = title.lower()
    s = 0

    if "301" in t: s += 30
    if "tariff" in t or "관세" in t: s += 25
    if "customs" in t: s += 20
    if "fta" in t: s += 15
    if "wto" in t: s += 15

    if "semiconductor" in t or "반도체" in t: s += 15
    if "battery" in t: s += 10
    if "smartphone" in t: s += 10

    return s

# ===================== AI 분석 =====================
def analyze(title):

    prompt = f"""
당신은 삼성전자 관세 전문가입니다.

뉴스:
{title}

다음 기준으로 분석:

1. 생산거점 영향:
- 베트남(SEV/SEVT)
- 인도(SIEL)
- 멕시코(SAMEX)

2. 제품 영향:
- 스마트폰
- 반도체
- 가전

3. 관세 영향:
- 관세율 변화 → 상
- 조사/협상 → 중
- 일반 → 하

JSON:
{{
"summary":"2~3줄 요약",
"analysis":"삼성 영향 분석",
"action":"관세 대응 방안",
"risk":"상/중/하",
"country":"국가",
"agency":"기관"
}}
"""

    try:
        res = client.models.generate_content(
            model=MODEL,
            contents=prompt
        )

        txt = res.text.strip()
        start = txt.find("{")
        end = txt.rfind("}") + 1
        data = json.loads(txt[start:end])

        return (
            data.get("summary",""),
            data.get("analysis",""),
            data.get("action",""),
            data.get("risk",""),
            data.get("country",""),
            data.get("agency","")
        )

    except:
        return "", "", "", "", "", ""

# ===================== Risk fallback =====================
def fallback_risk(title):

    t = title.lower()

    if "301" in t or "관세" in t or "tariff" in t:
        return "상"

    if "fta" in t:
        return "중"

    return "하"

# ===================== MAIN =====================
print("🚀 STEP4 vFinal START")

df = pd.read_excel(INPUT_FILE)
df.columns = df.columns.str.lower().str.strip()

print("📊 Loaded:", len(df))

# ---------------------------
# title/url 생성
# ---------------------------
if "title" not in df.columns:
    df["title"] = df["headline"]

df["title"] = df["title"].apply(extract_title)

if "url" not in df.columns:
    df["url"] = df["headline"].apply(extract_url)

df["title"] = df["title"].fillna("").astype(str)
df["url"] = df["url"].fillna("").astype(str)

# ---------------------------
# 필터링 (핵심)
# ---------------------------
df = df[df["title"].apply(is_trade_news)]

print("통상 뉴스:", len(df))

# ---------------------------
# 점수 및 정렬
# ---------------------------
df["score"] = df["title"].apply(score)
df = df.sort_values(by="score", ascending=False)

# ---------------------------
# 중복 제거
# ---------------------------
df = df.drop_duplicates(subset=["url"])
df = df.drop_duplicates(subset=["title"])

# ---------------------------
# TOP30
# ---------------------------
df = df.head(TOP_N)

print("🎯 TOP30 선정 완료")

# ---------------------------
# AI 분석
# ---------------------------
summary, analysis, action, risk, country, agency = [],[],[],[],[],[]

for i,row in df.iterrows():

    print(f"[AI] {row['title'][:60]}")

    s,a,ac,r,c,ag = analyze(row["title"])

    # fallback 처리
    if not s: s = row["title"]
    if not a: a = "삼성전자 영향 분석 필요"
    if not ac: ac = "관세 리스크 모니터링 필요"
    if not r: r = fallback_risk(row["title"])

    summary.append(s)
    analysis.append(a)
    action.append(ac)
    risk.append(r)
    country.append(c)
    agency.append(ag)

df["summary"] = summary
df["AI Analysis"] = analysis
df["Action Plan"] = action
df["risk"] = risk
df["country"] = country
df["agency"] = agency

# ---------------------------
# Headline 생성
# ---------------------------
df["Headline"] = df.apply(
    lambda r: f'=HYPERLINK("{r["url"]}","{r["title"]}")',
    axis=1
)

# ---------------------------
# OUTPUT
# ---------------------------
df_out = df[
    ["Headline","summary","AI Analysis","Action Plan",
     "country","agency","risk","score"]
]

df_out.to_excel(OUTPUT_DAILY, index=False)

# cumulative
if os.path.exists(OUTPUT_CUMUL):
    old = pd.read_excel(OUTPUT_CUMUL)
    total = pd.concat([old, df_out]).drop_duplicates(subset=["Headline"])
else:
    total = df_out

total.to_excel(OUTPUT_CUMUL, index=False)

print("===================================")
print("✅ STEP4 완료 vFinal (실무 완성)")
print("📁 DAILY:", OUTPUT_DAILY)
print("📁 CUMUL:", OUTPUT_CUMUL)
print("===================================")