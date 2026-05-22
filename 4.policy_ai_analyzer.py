# =========================================================
# GTI STEP4 FINAL v10.0 + LEGAL PRIORITY PATCH
# file name : 4.policy_ai_analyzer.py
#
# INPUT  : C:/temp/3.news_ai_summary.xlsx
# DAILY  : C:/temp/news_raw.xlsx
# CUMUL  : C:/temp/news_cumulative.xlsx
#
# 기준:
# - v10.0 기존 구조 유지
# - 법령/규칙/고시/공고/입법예고/행정규칙 신규 게시물 최우선 선정 로직 추가
# - GTI Executive Prompt 유지
# - HS/원산지/FTA/수출통제/SEV·SEVT·SIEL·SAMEX 강제 반영
# =========================================================

import os
import re
import json
import time
import random
import warnings
from datetime import datetime
from difflib import SequenceMatcher
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, MarkupResemblesLocatorWarning
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=MarkupResemblesLocatorWarning)

# =========================
# CONFIG
# =========================
INPUT_FILE = r"C:/temp/3.news_ai_summary.xlsx"
OUTPUT_DAILY = r"C:/temp/news_raw.xlsx"
OUTPUT_CUMUL = r"C:/temp/news_cumulative.xlsx"

TOP_N = 30
MAX_PER_CLUSTER = 1
MAX_KOREA_NEWS = 8
MAX_SOURCE_NEWS = 8

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = "gemini-2.0-flash"
USE_AI = bool(GEMINI_API_KEY)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/123.0 Safari/537.36"
    )
}

client = None
if USE_AI:
    try:
        from google import genai
        client = genai.Client(api_key=GEMINI_API_KEY)
    except Exception:
        client = None
        USE_AI = False


# =========================
# BASIC UTILS
# =========================
def safe_str(x):
    if x is None:
        return ""

    if isinstance(x, (list, tuple, set)):
        return ", ".join([str(v).strip() for v in x if str(v).strip()])

    if isinstance(x, dict):
        return ", ".join([f"{k}:{v}" for k, v in x.items()])

    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass

    return str(x).strip()


def normalize_text(x):
    x = safe_str(x).lower()
    x = re.sub(r"https?://\S+", " ", x)
    x = re.sub(r"&quot;|&amp;|&lt;|&gt;", " ", x)
    x = re.sub(r"[^a-z0-9가-힣\s]", " ", x)
    x = re.sub(r"\s+", " ", x).strip()
    return x


def clean_ai_text(x):
    x = safe_str(x)
    x = x.replace("**", "").replace("##", "")
    x = re.sub(r"\s+", " ", x).strip()
    return x


def trim(x, n=650):
    return clean_ai_text(x)[:n].strip()


def split_sentences(text):
    text = clean_ai_text(text)
    if not text:
        return []
    parts = re.split(r"(?<=[.!?。！？])\s+|(?<=[다요음함임됨니다])\.\s*", text)
    parts = [p.strip(" .") for p in parts if p and p.strip(" .")]
    return parts


def limit_sentences(text, max_sentences=2, max_len=650):
    text = clean_ai_text(text)
    if not text:
        return ""
    sentences = split_sentences(text)
    if len(sentences) >= 2:
        text = ". ".join(sentences[:max_sentences]).strip()
        if not text.endswith("."):
            text += "."
    return trim(text, max_len)


def parse_excel_hyperlink_formula(x):
    s = safe_str(x)
    m = re.search(r'=HYPERLINK\("([^"]+)","([^"]+)"\)', s, re.I)
    if m:
        return m.group(2), m.group(1)
    return "", ""


# =========================
# COLUMN NORMALIZATION
# =========================
def normalize_columns(df):
    df.columns = [str(c).strip().lower() for c in df.columns]

    if "headline" not in df.columns and "title" in df.columns:
        df["headline"] = df["title"]

    if "title" not in df.columns and "headline" in df.columns:
        df["title"] = df["headline"]

    if "url" not in df.columns:
        if "link" in df.columns:
            df["url"] = df["link"]
        else:
            df["url"] = ""

    if "source" not in df.columns:
        df["source"] = ""

    if "date" not in df.columns:
        df["date"] = ""

    return df


def recover_title_url(row):
    title = safe_str(row.get("title", ""))
    url = safe_str(row.get("url", ""))

    if title.lower() in ["0", "nan", "none"]:
        title = ""

    h_title, h_url = parse_excel_hyperlink_formula(row.get("headline", ""))
    if not title and h_title:
        title = h_title
    if not url and h_url:
        url = h_url

    if not title:
        headline = safe_str(row.get("headline", ""))
        if headline and headline.lower() not in ["0", "nan", "none"]:
            title = headline

    if not title and url.startswith("http"):
        path = urlparse(url).path.strip("/").split("/")[-1]
        title = path[:80] if path else url[:80]

    return pd.Series({"title": title.strip(), "url": url.strip()})


# =========================
# KEYWORDS
# =========================
NOISE_KEYWORDS = [
    "연예", "배우", "드라마", "예능", "영화", "스포츠", "축구", "야구",
    "결혼", "이혼", "맛집", "날씨", "복권", "주가", "증시",
    "교육", "세미나", "설명회", "컨퍼런스", "워크숍", "모집", "채용",
    "webinar", "seminar", "training", "workshop", "celebrity", "movie",
    "부동산", "아파트", "여행", "맛집",
]

TRADE_KEYWORDS = [
    "관세", "tariff", "duty", "customs", "세관", "통관",
    "fta", "cepa", "epa", "원산지", "origin",
    "수출", "수입", "export", "import",
    "301", "section 301", "ustr", "wto",
    "anti-dumping", "antidumping", "반덤핑",
    "countervailing", "상계관세",
    "제재", "sanction",
    "export control", "수출통제",
    "cbam", "supply chain", "공급망",
    "trade", "통상", "무역", "수입규제",
]

# =========================
# LEGAL PRIORITY PATCH
# =========================
LEGAL_PRIORITY_KEYWORDS = [
    "법령", "법률", "시행령", "시행규칙", "행정규칙",
    "고시", "공고", "훈령", "예규", "지침", "개정",
    "입법예고", "행정예고", "관보", "시행", "공포",
    "규칙", "규정", "notice", "regulation", "rule", "law",
    "decree", "ordinance", "amendment", "federal register",
    "final rule", "proposed rule",
]

LEGAL_SOURCE_KEYWORDS = [
    "관세청", "관보", "법제처", "국가법령정보센터",
    "국민참여입법센터", "입법예고", "행정예고",
    "ustr", "cbp", "federal register", "eu commission",
    "wto", ".gov", "government",
]

SAMSUNG_KEYWORDS = [
    "samsung", "삼성",
    "semiconductor", "chip", "반도체",
    "smartphone", "mobile", "휴대폰", "스마트폰", "galaxy",
    "display", "디스플레이",
    "battery", "배터리",
    "electronics", "전자", "consumer electronics", "가전",
    "network", "네트워크",
    "medical", "의료기기",
    "server", "서버",
    "memory", "메모리",
    "hbm", "dram", "nand",
]

PRODUCTION_COUNTRIES = [
    "vietnam", "베트남",
    "india", "인도",
    "mexico", "멕시코",
    "china", "중국",
    "korea", "한국",
    "poland", "폴란드",
    "slovakia", "슬로바키아",
    "turkey", "튀르키예", "터키",
    "brazil", "브라질",
    "indonesia", "인도네시아",
]

GOV_KEYWORDS = [
    "관세청", "ustr", "cbp", "customs and border protection",
    "wto", "eu commission", "european commission",
    "ministry", "commerce department", "상무부",
    "산업부", "기재부", "government", ".gov",
    "official", "commission", "국세청", "세관",
]

LOW_RELEVANCE_KEYWORDS = [
    "beef", "소고기", "쇠고기", "농산물", "farm", "farmer", "agriculture", "농업",
    "seafood", "수산물", "fish", "gold", "silver", "귀금속",
    "fentanyl", "마약", "drug", "narcotic", "cocaine",
    "firearm", "gun", "weapon", "suppressor", "총기",
    "관세청장", "차장 발탁", "인사", "임명",
    "방산", "호위함", "잠수함", "무기",
]

EXCLUDE_IF_LOW_ONLY = [
    "fentanyl", "마약", "drug", "narcotic", "cocaine",
    "firearm", "gun", "weapon", "suppressor", "총기",
    "관세청장", "차장 발탁", "인사", "임명",
    "방산", "호위함", "잠수함", "무기",
]

FORBIDDEN_PHRASES = [
    "모니터링 필요",
    "영향 분석 필요",
    "대응 필요",
    "검토 필요",
    "관련 이슈",
    "주의 필요",
    "리스크 존재",
    "본사 관세담당자는",
    "직접 영향은 제한적이나 글로벌 통상 환경",
    "관세 정책 변화 영향 분석 필요",
    "관세 영향 점검 및 대응 필요",
]


def is_noise(title):
    t = normalize_text(title)
    return any(k.lower() in t for k in NOISE_KEYWORDS)


def is_trade_news(title):
    t = normalize_text(title)
    return any(k.lower() in t for k in TRADE_KEYWORDS)


def is_legal_priority(title, source=""):
    t = normalize_text(title + " " + source)
    has_legal = any(k.lower() in t for k in LEGAL_PRIORITY_KEYWORDS)
    has_trade = any(k.lower() in t for k in TRADE_KEYWORDS)
    has_legal_source = any(k.lower() in t for k in LEGAL_SOURCE_KEYWORDS)
    return has_legal and (has_trade or has_legal_source)


def legal_priority_score(title, source=""):
    t = normalize_text(title + " " + source)
    score = 0

    if any(k.lower() in t for k in LEGAL_PRIORITY_KEYWORDS):
        score += 160

    if any(k.lower() in t for k in LEGAL_SOURCE_KEYWORDS):
        score += 100

    if any(k in t for k in ["관세", "customs", "tariff", "origin", "원산지", "fta", "export control", "수출통제", "sanction", "제재"]):
        score += 60

    return score


def is_low_relevance_only(title):
    t = normalize_text(title)

    has_low = any(k.lower() in t for k in EXCLUDE_IF_LOW_ONLY)
    has_strong = any(k.lower() in t for k in [
        "samsung", "삼성", "semiconductor", "반도체", "smartphone", "mobile",
        "electronics", "전자", "display", "battery", "배터리",
        "301", "ustr", "수출통제", "export control", "제재", "sanction",
        "고시", "공고", "법령", "행정규칙", "입법예고", "관보",
    ])

    return has_low and not has_strong


# =========================
# COUNTRY / AGENCY
# =========================
def extract_country(title):
    t = normalize_text(title)

    mapping = [
        ("united states", "미국"), ("u s", "미국"), ("usa", "미국"), ("미국", "미국"),
        ("china", "중국"), ("중국", "중국"),
        ("eu", "EU"), ("europe", "EU"), ("유럽", "EU"),
        ("india", "인도"), ("인도", "인도"),
        ("vietnam", "베트남"), ("베트남", "베트남"),
        ("mexico", "멕시코"), ("멕시코", "멕시코"),
        ("japan", "일본"), ("일본", "일본"),
        ("korea", "한국"), ("한국", "한국"),
        ("brazil", "브라질"), ("브라질", "브라질"),
        ("australia", "호주"), ("호주", "호주"),
        ("canada", "캐나다"), ("캐나다", "캐나다"),
        ("uk", "영국"), ("britain", "영국"), ("영국", "영국"),
        ("indonesia", "인도네시아"), ("인도네시아", "인도네시아"),
    ]

    found = []
    for k, v in mapping:
        if k in t and v not in found:
            found.append(v)

    return ", ".join(found[:2]) if found else "글로벌"


def extract_agency(title, source=""):
    t = normalize_text(title + " " + source)

    mapping = [
        ("ustr", "USTR"),
        ("cbp", "CBP"),
        ("customs and border protection", "CBP"),
        ("관세청", "관세청"),
        ("wto", "WTO"),
        ("european commission", "EU Commission"),
        ("eu commission", "EU Commission"),
        ("commerce department", "U.S. Department of Commerce"),
        ("상무부", "U.S. Department of Commerce"),
        ("customs", "세관"),
        ("세관", "세관"),
        ("ministry", "정부기관"),
        ("commission", "정부/위원회"),
        ("법제처", "법제처"),
        ("국가법령정보센터", "국가법령정보센터"),
        ("관보", "관보"),
        ("federal register", "Federal Register"),
    ]

    for k, v in mapping:
        if k in t:
            return v

    return "N/A"


# =========================
# ISSUE CLUSTERING
# =========================
CLUSTER_RULES = [
    ("LEGAL_CUSTOMS_RULE", ["고시"]),
    ("LEGAL_CUSTOMS_RULE", ["공고"]),
    ("LEGAL_CUSTOMS_RULE", ["행정규칙"]),
    ("LEGAL_CUSTOMS_RULE", ["입법예고"]),
    ("LEGAL_CUSTOMS_RULE", ["관보"]),
    ("LEGAL_CUSTOMS_RULE", ["federal", "register"]),
    ("US_CHINA_TARIFF_DEAL", ["중국", "미국", "관세", "인하"]),
    ("US_CHINA_TARIFF_DEAL", ["중", "미", "관세", "인하"]),
    ("US_CHINA_TARIFF_DEAL", ["미중", "관세"]),
    ("US_CHINA_TARIFF_DEAL", ["china", "us", "tariff"]),
    ("US_CHINA_TARIFF_DEAL", ["china", "u s", "tariff"]),
    ("US_CHINA_TARIFF_DEAL", ["trump", "china", "tariff"]),
    ("SECTION_301", ["301"]),
    ("CBP_CUSTOMS_POLICY", ["cbp"]),
    ("CBP_CUSTOMS_POLICY", ["customs", "border", "protection"]),
    ("AD_CVD", ["반덤핑"]),
    ("AD_CVD", ["anti", "dumping"]),
    ("FTA_GENERAL", ["fta"]),
    ("ORIGIN_RULE", ["원산지"]),
    ("ORIGIN_RULE", ["origin"]),
    ("EXPORT_CONTROL", ["수출통제"]),
    ("EXPORT_CONTROL", ["export", "control"]),
    ("SANCTION", ["sanction"]),
    ("SANCTION", ["제재"]),
    ("CBAM", ["cbam"]),
]


def issue_cluster_key(title):
    t = normalize_text(title)

    for key, terms in CLUSTER_RULES:
        if all(term in t for term in terms):
            return key

    tokens = [x for x in t.split() if len(x) >= 2]
    important = [
        x for x in tokens
        if x in [
            "미국", "중국", "인도", "베트남", "멕시코", "eu",
            "ustr", "cbp", "관세청", "tariff", "customs",
            "origin", "fta", "wto", "관세", "원산지", "반덤핑", "수출통제",
            "고시", "공고", "법령", "행정규칙", "입법예고",
        ]
    ]

    if important:
        return "_".join(important[:5])

    return t[:50]


# =========================
# SCORING
# =========================
def policy_score(title, source=""):
    t = normalize_text(title + " " + source)
    s = 0

    weights = {
        "section 301": 50,
        "301": 50,
        "ustr": 45,
        "관세": 38,
        "tariff": 38,
        "duty": 25,
        "customs": 24,
        "세관": 24,
        "cbp": 40,
        "anti dumping": 42,
        "antidumping": 42,
        "반덤핑": 42,
        "countervailing": 35,
        "상계관세": 35,
        "export control": 45,
        "수출통제": 45,
        "sanction": 38,
        "제재": 38,
        "cbam": 38,
        "wto": 30,
        "fta": 26,
        "cepa": 22,
        "원산지": 35,
        "origin": 30,
    }

    for k, v in weights.items():
        if k in t:
            s += v

    if any(k.lower() in t for k in GOV_KEYWORDS):
        s += 30

    return s


def samsung_score(title):
    t = normalize_text(title)
    s = 0

    for k in SAMSUNG_KEYWORDS:
        if k.lower() in t:
            s += 35

    for k in PRODUCTION_COUNTRIES:
        if k.lower() in t:
            s += 18

    if any(k.lower() in t for k in ["vietnam", "베트남", "india", "인도", "mexico", "멕시코"]):
        s += 20

    if any(k.lower() in t for k in LOW_RELEVANCE_KEYWORDS):
        s -= 45

    return s


def region_balance_score(title):
    t = normalize_text(title)
    s = 0

    if any(k in t for k in ["미국", "united states", "usa", "u s", "ustr", "cbp"]):
        s += 25
    if any(k in t for k in ["중국", "china"]):
        s += 18
    if any(k in t for k in ["eu", "europe", "유럽"]):
        s += 18
    if any(k in t for k in ["인도", "india", "베트남", "vietnam", "멕시코", "mexico"]):
        s += 18

    if "한국" in t and not any(k in t for k in ["관세청", "ustr", "301", "관세", "반덤핑", "수출통제", "삼성", "반도체", "고시", "공고", "법령"]):
        s -= 15

    return s


def final_score(row):
    return (
        legal_priority_score(row["title"], row.get("source", ""))
        + policy_score(row["title"], row.get("source", ""))
        + samsung_score(row["title"])
        + region_balance_score(row["title"])
    )


def importance_by_score(score, title, source=""):
    if is_legal_priority(title, source):
        return "상"

    t = normalize_text(title)

    if any(k in t for k in LOW_RELEVANCE_KEYWORDS) and score < 120:
        return "하"

    if score >= 120:
        return "상"
    if score >= 60:
        return "중"
    return "하"


def risk_by_title(title, score, source=""):
    if is_legal_priority(title, source):
        return "상"

    t = normalize_text(title)

    if any(k in t for k in LOW_RELEVANCE_KEYWORDS) and not any(k in t for k in ["삼성", "반도체", "301", "수출통제", "ustr"]):
        return "하"

    if any(k in t for k in ["301", "추가관세", "반덤핑", "anti dumping", "antidumping", "수출통제", "export control", "제재", "sanction"]):
        return "상"

    if any(k in t for k in ["관세", "tariff", "customs", "fta", "원산지", "origin"]):
        return "상" if score >= 120 else "중"

    return "하"


# =========================
# NEWS TYPE
# =========================
def classify_news_type(title):
    t = normalize_text(title)

    if any(k.lower() in t for k in LEGAL_PRIORITY_KEYWORDS):
        return "LEGAL_RULE"
    if any(k in t for k in ["301", "section 301", "추가관세", "tariff", "관세"]):
        return "TARIFF"
    if any(k in t for k in ["fta", "cepa", "원산지", "origin"]):
        return "FTA_ORIGIN"
    if any(k in t for k in ["anti dumping", "antidumping", "반덤핑", "상계관세", "countervailing"]):
        return "AD_CVD"
    if any(k in t for k in ["export control", "수출통제", "sanction", "제재"]):
        return "EXPORT_CONTROL"
    if any(k in t for k in ["customs", "세관", "통관", "cbp"]):
        return "CUSTOMS_AUDIT"
    if any(k in t for k in ["cbam", "carbon", "탄소"]):
        return "CBAM"
    return "GENERAL_TRADE"


# =========================
# DEDUP + CLUSTER LIMIT
# =========================
def is_similar(a, b, threshold=0.80):
    a = normalize_text(a)
    b = normalize_text(b)

    if not a or not b:
        return False

    if a[:35] == b[:35]:
        return True

    return SequenceMatcher(None, a, b).ratio() >= threshold


def dedup_news(df):
    df = df.copy()

    df["url_norm"] = df["url"].fillna("").astype(str).str.strip()
    df = df.drop_duplicates(subset=["url_norm"], keep="first")

    selected_titles = []
    rows = []

    for _, row in df.sort_values("score", ascending=False).iterrows():
        title = row["title"]
        if any(is_similar(title, old) for old in selected_titles):
            continue
        selected_titles.append(title)
        rows.append(row)

    if not rows:
        return df.head(0)

    return pd.DataFrame(rows).drop(columns=["url_norm"], errors="ignore")


def select_balanced_top(df, top_n=30):
    country_count = {}
    source_count = {}
    cluster_count = {}
    selected = []

    legal_df = df[df.apply(lambda r: is_legal_priority(r["title"], r.get("source", "")), axis=1)]
    normal_df = df[~df.index.isin(legal_df.index)]

    ordered = pd.concat([
        legal_df.sort_values("score", ascending=False),
        normal_df.sort_values("score", ascending=False)
    ])

    for _, row in ordered.iterrows():
        country = safe_str(row.get("country_rule", "글로벌"))
        source = safe_str(row.get("source", ""))
        cluster = safe_str(row.get("cluster_key", ""))

        c_key = country.split(",")[0].strip() if country else "글로벌"
        s_key = source[:40]

        if cluster_count.get(cluster, 0) >= MAX_PER_CLUSTER:
            continue

        if c_key == "한국" and country_count.get(c_key, 0) >= MAX_KOREA_NEWS:
            continue

        if s_key and source_count.get(s_key, 0) >= MAX_SOURCE_NEWS:
            continue

        selected.append(row)
        country_count[c_key] = country_count.get(c_key, 0) + 1
        source_count[s_key] = source_count.get(s_key, 0) + 1
        cluster_count[cluster] = cluster_count.get(cluster, 0) + 1

        if len(selected) >= top_n:
            break

    if len(selected) < top_n:
        selected_titles = {r["title"] for r in selected}

        for _, row in ordered.iterrows():
            if row["title"] in selected_titles:
                continue

            cluster = safe_str(row.get("cluster_key", ""))
            if cluster_count.get(cluster, 0) >= MAX_PER_CLUSTER:
                continue

            selected.append(row)
            cluster_count[cluster] = cluster_count.get(cluster, 0) + 1

            if len(selected) >= top_n:
                break

    return pd.DataFrame(selected).reset_index(drop=True)


# =========================
# BODY FETCH
# =========================
def safe_request(url):
    if not safe_str(url).startswith("http"):
        return ""

    for _ in range(2):
        try:
            r = requests.get(url, headers=HEADERS, timeout=8, allow_redirects=True)
            if r.status_code == 200 and r.text:
                return r.text
        except Exception:
            pass
        time.sleep(random.uniform(0.5, 1.2))

    return ""


def fetch_body(url):
    html = safe_request(url)
    if not html:
        return ""

    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "noscript", "header", "footer", "nav"]):
        tag.decompose()

    text = soup.get_text(" ")
    text = re.sub(r"\s+", " ", text).strip()
    return text[:3000]


# =========================
# AI ANALYSIS
# =========================
def parse_json_from_text(txt):
    txt = safe_str(txt)
    start = txt.find("{")
    end = txt.rfind("}") + 1

    if start < 0 or end <= start:
        return {}

    try:
        return json.loads(txt[start:end])
    except Exception:
        return {}


def bad_ai_text(text, min_len=80):
    t = clean_ai_text(text)
    if not t:
        return True
    if len(t) < min_len:
        return True

    if any(p in t for p in FORBIDDEN_PHRASES):
        return True

    if t.count("삼성전자") > 3:
        return True

    return False


def analyze_with_ai(title, body, news_type):
    if not USE_AI or client is None:
        return {}

    prompt = f"""
너는 삼성전자 본사 관세전략 담당 임원이다.
아래 뉴스 제목과 본문을 기준으로 삼성전자 관세·통상 업무 관점의 GTI 보고서용 분석을 작성한다.

[뉴스 유형]
{news_type}

[제목]
{title}

[본문 일부]
{body[:2200]}

[반드시 아래 기준으로 작성]
- 2문장 이내
- 실행지시형
- HS/원산지/FTA/수출통제 포함
- 삼성 생산거점(SEV/SEVT/SIEL/SAMEX) 언급
- 관세전문가 어조
- 일반론 금지
- "모니터링 필요" 금지
- 실제 실무지시 작성
- 법령/규칙/고시/공고/입법예고/행정규칙이면 시행일, 적용대상, 내부통제 변경 필요성을 우선 판단

[삼성전자 분석 기준]
- 생산거점: SEV/SEVT(베트남), SIEL(인도), SAMEX(멕시코), 중국, 한국, 폴란드, 브라질
- 제품군: Mobile, Consumer Electronics, Network Equipment, Semiconductor/Component, Display
- 관세 포인트: HS Code, 관세율, 301조, 반덤핑/상계관세, 원산지, FTA, 수출통제, 제재, 통관심사, 과세가격
- 직접 영향: 관세율/수입규제/수출통제/원산지 규정이 삼성 제품·생산국·판매국과 연결될 때
- 간접 영향: 통상협상, 시장접근, 공급망, 물류, 세관 집행 강화 등

[Action Plan 작성 규칙]
- 반드시 실행 동사 사용: 확인, 매핑, 산출, 비교, 업데이트, 공유, 준비, 재산출
- 아래 항목 중 최소 3개를 포함: HS Mapping, 생산거점별 원산지 영향, FTA 적용 가능 여부, 수출통제/EAR 대상 여부, 북미/EU 판매법인 영향
- "검토 필요", "대응 필요", "모니터링 필요" 금지

[금지 표현]
- 모니터링 필요
- 영향 분석 필요
- 대응 필요
- 검토 필요
- 관련 이슈
- 주의 필요
- 리스크 존재

[좋은 예시]
AI Analysis:
미국 추가관세 적용 시 SEV/SEVT 생산 Mobile 제품의 CIF 기준 과세가격 상승이 북미 판매법인의 가격 경쟁력에 직접 반영될 수 있으며, SIEL 생산품은 FTA 적용 여부에 따라 관세율 차이가 발생한다.

Action Plan:
① 미국향 Mobile 모델 HS Mapping을 재점검하고 ② SEV/SEVT·SIEL 원산지 판정 구조와 FTA 적용세율을 재산출하며 ③ EAR/수출통제 대상 여부와 북미 판매법인 가격전가 영향을 사업부에 공유한다.

[출력 JSON만 작성]
{{
  "summary": "뉴스 핵심 2문장 이내",
  "ai_analysis": "삼성전자 관세·통상 영향 2문장 이내",
  "action_plan": "관세 담당자 실행지시 2문장 이내",
  "country": "대표 2개국 이하",
  "agency": "대표 기관 2개 이하",
  "risk": "상/중/하"
}}
"""

    try:
        res = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        return parse_json_from_text(res.text)
    except Exception:
        return {}


# =========================
# EXPERT FALLBACK ANALYSIS
# =========================
def fallback_summary(title, body=""):
    if body and len(body) > 120:
        return limit_sentences(trim(body, 300), 2, 300)
    return f"{title} 관련 관세·통상 조치입니다."


def detect_product_area(title):
    t = normalize_text(title)
    areas = []

    if any(k in t for k in ["반도체", "semiconductor", "chip", "hbm", "dram", "nand"]):
        areas.append("Semiconductor/Component")
    if any(k in t for k in ["smartphone", "mobile", "스마트폰", "휴대폰", "galaxy"]):
        areas.append("Mobile")
    if any(k in t for k in ["전자", "electronics", "가전", "consumer electronics"]):
        areas.append("Consumer Electronics")
    if any(k in t for k in ["network", "네트워크"]):
        areas.append("Network Equipment")
    if any(k in t for k in ["display", "디스플레이"]):
        areas.append("Display")

    return ", ".join(areas[:2]) if areas else "Mobile/CE/Component 공통"


def detect_geo_focus(title):
    t = normalize_text(title)
    geos = []

    for label, keys in [
        ("SEV/SEVT", ["vietnam", "베트남"]),
        ("SIEL", ["india", "인도"]),
        ("SAMEX", ["mexico", "멕시코"]),
        ("중국법인", ["china", "중국"]),
        ("북미 판매법인", ["united states", "usa", "u s", "미국"]),
        ("EU 판매법인", ["eu", "europe", "유럽"]),
        ("한국 본사/생산", ["korea", "한국"]),
    ]:
        if any(k in t for k in keys):
            geos.append(label)

    if not geos:
        geos = ["SEV/SEVT", "SIEL", "SAMEX"]

    return ", ".join(geos[:3])


def fallback_analysis(title, news_type="GENERAL_TRADE"):
    product = detect_product_area(title)
    geo = detect_geo_focus(title)

    if news_type == "LEGAL_RULE":
        return (
            f"해당 법령·규칙 신규 게시물은 {geo} 생산품의 HS·원산지·FTA·수출통제 내부통제 기준 변경 여부를 즉시 판단해야 하는 사안입니다. "
            f"{product} 적용 품목과 시행일을 기준으로 통관 신고 기준, 증빙 보관, 법인·관세사 업무지침 변경 여부를 확정해야 합니다."
        )

    if news_type == "TARIFF":
        return (
            f"{geo} 관련 관세율·추가관세 변경은 {product} 제품의 HS별 수입원가와 북미/EU 판매법인 가격전가에 직접 반영됩니다. "
            "원산지 판정과 FTA 적용 가능성에 따라 동일 모델도 생산거점별 관세 부담이 달라집니다."
        )

    if news_type == "FTA_ORIGIN":
        return (
            f"{geo} 관련 FTA·원산지 기준 변화는 {product} 제품의 협정세율 적용과 사후검증 리스크를 직접 좌우합니다. "
            "BOM상 역외산 핵심부품 비중이 높으면 SEV/SEVT·SIEL·SAMEX 생산품의 원산지 충족률이 흔들릴 수 있습니다."
        )

    if news_type == "AD_CVD":
        return (
            f"반덤핑·상계관세 조치는 {product} 관련 원재료·부품의 조달비용과 우회수출 판정 리스크를 높입니다. "
            "SEV/SEVT·SIEL·SAMEX 생산품에 투입되는 대상 품목의 HS·원산지·거래가격 연결성을 즉시 확인해야 합니다."
        )

    if news_type == "EXPORT_CONTROL":
        return (
            "수출통제·제재 변화는 Semiconductor/Component와 Network Equipment 거래의 최종사용자·최종용도 심사를 강화합니다. "
            "SEV/SEVT·SIEL·SAMEX 출하품은 EAR·제재리스트·재수출 규정 적용 여부를 거래 전 단계에서 차단해야 합니다."
        )

    if news_type == "CUSTOMS_AUDIT":
        return (
            f"세관 집행 강화는 {product} 제품의 HS 분류, 과세가격, 원산지 증빙, FTA 적용 신고의 정합성을 직접 겨냥합니다. "
            "SEV/SEVT·SIEL·SAMEX 신고자료와 ERP Invoice·계약가격 간 불일치가 추징 포인트가 될 수 있습니다."
        )

    if news_type == "CBAM":
        return (
            "CBAM·탄소통상 조치는 EU향 제품의 원재료 탄소정보와 원산지·공급망 증빙을 결합해 요구할 가능성이 높습니다. "
            "SEV/SEVT·SIEL·SAMEX 생산품의 HS별 원재료 출처와 탄소자료 확보 수준을 EU 판매법인 기준으로 정렬해야 합니다."
        )

    return (
        f"{geo} 관련 통상정책 변화가 {product} 제품의 HS·원산지·FTA·수출통제 적용 여부와 연결되는지 1차 판정해야 합니다. "
        "직접 관련 품목이 확인되면 생산거점별 관세율·과세가격·증빙자료를 즉시 재산출합니다."
    )


def fallback_action(title, news_type="GENERAL_TRADE"):
    product = detect_product_area(title)
    geo = detect_geo_focus(title)

    if news_type == "LEGAL_RULE":
        return (
            f"① 신규 법령·규칙의 시행일·적용대상 HS를 확정하고 ② {geo} 생산거점별 원산지·FTA·수출통제 영향표를 업데이트하며 ③ 법인·관세사 신고지침과 증빙 보관 체크리스트를 개정합니다."
        )

    if news_type == "TARIFF":
        return (
            f"① {product} 미국/EU향 모델의 HS Mapping을 재점검하고 ② {geo} 생산거점별 원산지 판정과 FTA 적용세율을 재산출하며 ③ 관세율 시나리오별 원가·가격전가 영향을 사업부에 공유합니다."
        )

    if news_type == "FTA_ORIGIN":
        return (
            f"① {geo} 생산품의 BOM 기준 원산지 충족률을 재산출하고 ② FTA 적용 가능 HS를 분리하며 ③ 공급업체 원산지확인서·제조공정 증빙을 사후검증 패키지로 정리합니다."
        )

    if news_type == "AD_CVD":
        return (
            "① 대상 품목 HS와 공급처 원산지를 매핑하고 ② SEV/SEVT·SIEL·SAMEX 투입 부품과 연결성을 비교하며 ③ 반덤핑 관세율 반영 시 대체 공급처·가격조건을 구매부서와 재산출합니다."
        )

    if news_type == "EXPORT_CONTROL":
        return (
            "① 거래상대방·최종사용자·목적지 국가를 수출통제 리스트와 대조하고 ② EAR·재수출 규정 적용 여부를 판정하며 ③ 반도체·네트워크 출하 승인 기준을 법무·영업과 업데이트합니다."
        )

    if news_type == "CUSTOMS_AUDIT":
        return (
            f"① {geo} 최근 신고 건의 HS·과세가격·원산지·FTA 적용 내역을 샘플링하고 ② ERP Invoice와 신고금액·수량 불일치를 정리하며 ③ 관세사별 소명자료 패키지를 준비합니다."
        )

    if news_type == "CBAM":
        return (
            "① EU향 대상 HS를 식별하고 ② 원재료 원산지·탄소자료 확보 수준을 점검하며 ③ SEV/SEVT·SIEL·SAMEX 공급망별 CBAM 증빙 공백을 EU 판매법인에 공유합니다."
        )

    return (
        f"① {geo} 관련 품목의 HS·원산지·FTA 적용 여부를 1차 분류하고 ② 수출통제 대상 여부를 확인하며 ③ 직접 영향 품목만 GTI 후속 과제로 등록합니다."
    )


def enforce_gti_quality(text, kind="analysis", title="", news_type="GENERAL_TRADE"):
    text = limit_sentences(text, 2, 650)

    if any(p in text for p in FORBIDDEN_PHRASES):
        return ""

    if kind == "action":
        required_any = ["HS", "원산지", "FTA", "수출통제", "EAR", "SEV", "SEVT", "SIEL", "SAMEX"]
        if not any(k in text for k in required_any):
            return ""
        if len(text) < 80:
            return ""

    elif kind == "analysis":
        required_any = ["HS", "원산지", "FTA", "수출통제", "SEV", "SEVT", "SIEL", "SAMEX", "관세", "과세가격"]
        if not any(k in text for k in required_any):
            return ""
        if len(text) < 80:
            return ""

    return text


# =========================
# EXCEL OUTPUT
# =========================
def write_excel(df_out, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    df_out.to_excel(path, index=False)

    wb = load_workbook(path)
    ws = wb.active

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    link_font = Font(color="0563C1", underline="single")
    wrap = Alignment(wrap_text=True, vertical="top")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    col_map = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}

    headline_col = col_map.get("Headline")
    url_col = col_map.get("URL")

    if headline_col and url_col:
        for row in range(2, ws.max_row + 1):
            h = ws.cell(row=row, column=headline_col)
            u = ws.cell(row=row, column=url_col).value
            if u and str(u).startswith("http"):
                h.hyperlink = str(u)
                h.font = link_font

    widths = {
        "Date": 18,
        "Headline": 50,
        "importance": 10,
        "URL": 30,
        "source": 22,
        "last_checked": 18,
        "Summary": 58,
        "AI Analysis": 78,
        "Action Plan": 78,
        "Country": 18,
        "agency": 24,
        "risk": 8,
        "score": 8,
        "news_type": 18,
        "cluster_key": 22,
    }

    for col_name, width in widths.items():
        if col_name in col_map:
            ws.column_dimensions[get_column_letter(col_map[col_name])].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrap

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


# =========================
# MAIN
# =========================
def main():
    print("🚀 GTI STEP4 FINAL v10.0 + LEGAL PATCH START")
    print(f"[AI] {'ON' if USE_AI else 'OFF - rule fallback only'}")

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"INPUT FILE 없음: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE)
    df = normalize_columns(df)
    print(f"[LOAD] {len(df)} rows")

    recovered = df.apply(recover_title_url, axis=1)
    df["title"] = recovered["title"]
    df["url"] = recovered["url"]

    df["title"] = df["title"].apply(safe_str)
    df["url"] = df["url"].apply(safe_str)
    df["source"] = df["source"].apply(safe_str)
    df["date"] = df["date"].apply(safe_str)

    df = df[(df["title"] != "") & (df["title"] != "0")]
    print(f"[TITLE OK] {len(df)} rows")

    df = df[~df["title"].apply(is_noise)]
    print(f"[NOISE REMOVED] {len(df)} rows")

    df = df[
        df["title"].apply(is_trade_news)
        | df.apply(lambda r: is_legal_priority(r["title"], r.get("source", "")), axis=1)
    ]
    print(f"[TRADE/LEGAL FILTER] {len(df)} rows")

    df = df[~df["title"].apply(is_low_relevance_only)]
    print(f"[LOW-ONLY REMOVED] {len(df)} rows")

    df["country_rule"] = df["title"].apply(extract_country)
    df["agency_rule"] = df.apply(lambda r: extract_agency(r["title"], r.get("source", "")), axis=1)
    df["cluster_key"] = df["title"].apply(issue_cluster_key)
    df["score"] = df.apply(final_score, axis=1)
    df["importance"] = df.apply(lambda r: importance_by_score(r["score"], r["title"], r.get("source", "")), axis=1)
    df["risk_rule"] = df.apply(lambda r: risk_by_title(r["title"], r["score"], r.get("source", "")), axis=1)
    df["news_type"] = df["title"].apply(classify_news_type)

    df = dedup_news(df)
    print(f"[DEDUP] {len(df)} rows")

    top = select_balanced_top(df, TOP_N)
    print(f"[TOP] {len(top)} rows")

    records = []
    last_checked = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for idx, row in top.iterrows():
        title = safe_str(row["title"])
        url = safe_str(row["url"])
        news_type = safe_str(row.get("news_type", "GENERAL_TRADE"))

        print(f"[AI {idx + 1}/{len(top)}] {title[:80]}")

        body = fetch_body(url)
        ai = analyze_with_ai(title, body, news_type)

        summary = clean_ai_text(ai.get("summary", ""))
        analysis = clean_ai_text(ai.get("ai_analysis", ""))
        action = clean_ai_text(ai.get("action_plan", ""))
        country = clean_ai_text(ai.get("country", ""))
        agency = clean_ai_text(ai.get("agency", ""))
        risk = clean_ai_text(ai.get("risk", ""))

        if bad_ai_text(summary, min_len=40):
            summary = fallback_summary(title, body)

        analysis = enforce_gti_quality(analysis, "analysis", title, news_type)
        action = enforce_gti_quality(action, "action", title, news_type)

        if not analysis:
            analysis = fallback_analysis(title, news_type)

        if not action:
            action = fallback_action(title, news_type)

        analysis = limit_sentences(analysis, 2, 650)
        action = limit_sentences(action, 2, 650)

        if not country:
            country = safe_str(row.get("country_rule", "글로벌"))

        if not agency or agency == "N/A":
            agency = safe_str(row.get("agency_rule", "N/A"))

        if risk not in ["상", "중", "하"]:
            risk = safe_str(row.get("risk_rule", "중"))

        records.append({
            "Date": safe_str(row.get("date", "")),
            "Headline": title,
            "importance": safe_str(row.get("importance", "")),
            "URL": url,
            "source": safe_str(row.get("source", "")),
            "last_checked": last_checked,
            "Summary": trim(summary, 520),
            "AI Analysis": trim(analysis, 650),
            "Action Plan": trim(action, 650),
            "Country": country,
            "agency": agency,
            "risk": risk,
            "score": int(row.get("score", 0)),
            "news_type": news_type,
            "cluster_key": safe_str(row.get("cluster_key", "")),
        })

    out = pd.DataFrame(records)

    final_cols = [
        "Date", "Headline", "importance", "URL", "source", "last_checked",
        "Summary", "AI Analysis", "Action Plan",
        "Country", "agency", "risk", "score", "news_type", "cluster_key"
    ]

    out = out[final_cols]

    write_excel(out, OUTPUT_DAILY)

    if os.path.exists(OUTPUT_CUMUL):
        old = pd.read_excel(OUTPUT_CUMUL)
        total = pd.concat([old, out], ignore_index=True)
        if "URL" in total.columns:
            total = total.drop_duplicates(subset=["URL"], keep="last")
        else:
            total = total.drop_duplicates(subset=["Headline"], keep="last")
    else:
        total = out.copy()

    write_excel(total, OUTPUT_CUMUL)

    print("===================================")
    print("✅ GTI STEP4 FINAL v10.0 + LEGAL PATCH COMPLETE")
    print(f"📁 DAILY : {OUTPUT_DAILY}")
    print(f"📁 CUMUL : {OUTPUT_CUMUL}")
    print("===================================")


if __name__ == "__main__":
    main()