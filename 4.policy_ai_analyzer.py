# =========================================================
# GTI STEP4 STRUCTURAL FINAL v12.0
# Samsung Electronics HQ Customs / Trade Policy Daily Sensing
#
# INPUT  : C:/temp/3.news_ai_summary.xlsx
# DAILY  : C:/temp/news_raw.xlsx
#
# 
#  CUMUL  : C:/temp/news_cumulative.xlsx
#
# v12.0 핵심:
# 1) 기존 priority/score 단일 정렬 폐기
# 2) Reject → Category → Issue Cluster → Category Quota → Representative Pick
# 3) 법령/공식문서/고시/입법예고/USTR/CBP/Federal Register 최우선
# 4) 면세점/관광/교육/행사/커피/농산물/비료/방산행사/의약품/수주 강제 제외
# 5) Top30 중복 이슈 최소화
# 6) AI는 문장화 보조, 판단은 Rule Engine
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

CATEGORY_ORDER = [
    "A_LEGAL_OFFICIAL",
    "B_SEMICONDUCTOR_TARIFF",
    "C_ORIGIN_FTA_USMCA",
    "D_EXPORT_CONTROL_SANCTION",
    "E_AD_CVD_TRADE_REMEDY",
    "F_CUSTOMS_AUDIT_VALUATION",
    "G_CBAM_SUPPLY_CHAIN",
    "H_SAMSUNG_GEO_POLICY",
    "I_GENERAL_REFERENCE",
]

CATEGORY_QUOTA = {
    "A_LEGAL_OFFICIAL": 8,
    "B_SEMICONDUCTOR_TARIFF": 5,
    "C_ORIGIN_FTA_USMCA": 5,
    "D_EXPORT_CONTROL_SANCTION": 4,
    "E_AD_CVD_TRADE_REMEDY": 4,
    "F_CUSTOMS_AUDIT_VALUATION": 3,
    "G_CBAM_SUPPLY_CHAIN": 3,
    "H_SAMSUNG_GEO_POLICY": 4,
    "I_GENERAL_REFERENCE": 2,
}

MAX_SOURCE = 5
MAX_KOREA = 6

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
        return ", ".join(str(v).strip() for v in x if str(v).strip())
    if isinstance(x, dict):
        return ", ".join(f"{k}:{v}" for k, v in x.items())
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    return str(x).strip()


def norm(x):
    x = safe_str(x).lower()
    x = re.sub(r"https?://\S+", " ", x)
    x = re.sub(r"&quot;|&amp;|&lt;|&gt;|&#x27;", " ", x)
    x = re.sub(r"[^a-z0-9가-힣\s./_-]", " ", x)
    x = re.sub(r"\s+", " ", x).strip()
    return x


def clean_text(x):
    x = safe_str(x)
    x = x.replace("**", "").replace("##", "")
    x = re.sub(r"\s+", " ", x).strip()
    return x


def trim(x, n=650):
    return clean_text(x)[:n].strip()


def contains(text, keywords):
    t = norm(text)
    return any(k.lower() in t for k in keywords)


def count_hits(text, keywords):
    t = norm(text)
    return sum(1 for k in keywords if k.lower() in t)


def limit_sentences(text, max_sentences=2, max_len=650):
    text = clean_text(text)
    if not text:
        return ""
    parts = re.split(r"(?<=[.!?。！？])\s+|(?<=다\.)\s+|(?<=니다\.)\s+", text)
    parts = [p.strip(" .") for p in parts if p and p.strip(" .")]
    if len(parts) >= max_sentences:
        text = ". ".join(parts[:max_sentences])
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
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    if "headline" not in df.columns and "title" in df.columns:
        df["headline"] = df["title"]
    if "title" not in df.columns and "headline" in df.columns:
        df["title"] = df["headline"]
    if "url" not in df.columns:
        df["url"] = df["link"] if "link" in df.columns else ""
    if "source" not in df.columns:
        df["source"] = ""
    if "date" not in df.columns:
        df["date"] = ""
    if "summary" not in df.columns:
        df["summary"] = ""

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
        title = path[:90] if path else url[:90]

    return pd.Series({"title": title.strip(), "url": url.strip()})


# =========================
# KEYWORDS
# =========================
LEGAL_KW = [
    "법령", "법률", "시행령", "시행규칙", "행정규칙", "고시", "공고", "훈령",
    "예규", "지침", "개정", "입법예고", "행정예고", "관보", "시행", "공포",
    "규칙", "규정", "notice", "regulation", "rule", "law", "decree",
    "ordinance", "amendment", "federal register", "final rule", "proposed rule",
    "cbp notice", "ustr notice", "official journal"
]

OFFICIAL_KW = [
    "관세청", "법제처", "국가법령정보센터", "국민참여입법센터", "관보",
    "ustr", "cbp", "federal register", "eu commission", "european commission",
    "official journal", "taxud", "wto", "bis", "mofcom", ".gov", "customs",
    "border protection", "department of commerce", "commerce department"
]

TRADE_KW = [
    "관세", "tariff", "duty", "customs", "세관", "통관", "fta", "cepa", "epa",
    "원산지", "origin", "수출", "수입", "export", "import", "301", "232",
    "section 301", "section 232", "ustr", "wto", "anti-dumping", "antidumping",
    "반덤핑", "countervailing", "상계관세", "제재", "sanction", "export control",
    "수출통제", "cbam", "supply chain", "공급망", "trade", "통상", "무역",
    "수입규제", "valuation", "과세가격", "customs valuation", "usmca", "ear",
    "forced labor", "uflpa"
]

SEMICON_KW = [
    "semiconductor", "chip", "반도체", "hbm", "dram", "nand", "memory",
    "wafer", "foundry", "fab", "processor", "ai chip"
]

PRODUCT_KW = [
    "samsung", "삼성", "mobile", "smartphone", "휴대폰", "스마트폰", "galaxy",
    "electronics", "전자", "consumer electronics", "가전", "tv", "television",
    "display", "디스플레이", "battery", "배터리", "network", "네트워크",
    "server", "medical", "의료기기"
]

GEO_KW = [
    "vietnam", "베트남", "india", "인도", "mexico", "멕시코", "china", "중국",
    "korea", "한국", "poland", "폴란드", "slovakia", "슬로바키아",
    "brazil", "브라질", "indonesia", "인도네시아", "united states", "usa",
    "u.s.", "미국", "eu", "europe", "유럽"
]

ORIGIN_FTA_KW = [
    "fta", "cepa", "epa", "원산지", "origin", "rules of origin", "usmca",
    "rvc", "regional value content", "certificate of origin", "coo", "co "
]

EXPORT_CONTROL_KW = [
    "export control", "수출통제", "sanction", "제재", "ear", "bis",
    "entity list", "restricted", "dual-use", "재수출", "최종사용자", "end user",
    "end-use", "첨단반도체"
]

AD_CVD_KW = [
    "anti-dumping", "antidumping", "반덤핑", "countervailing", "상계관세",
    "trade remedy", "safeguard", "세이프가드", "mip", "minimum import price"
]

CUSTOMS_AUDIT_KW = [
    "customs", "세관", "통관", "cbp", "valuation", "과세가격", "classification",
    "품목분류", "hs code", "hs classification", "audit", "심사", "enforcement"
]

CBAM_KW = ["cbam", "carbon", "탄소", "emissions", "steel tariff", "철강 관세"]

STRICT_REJECT_KW = [
    "롯데면세점", "면세점", "duty free", "관광", "tourism", "공공캐릭터", "팝업존",
    "교육", "세미나", "설명회", "컨퍼런스", "워크숍", "인재 양성", "전문가 배출",
    "합격", "training", "webinar", "seminar", "workshop", "conference",
    "커피", "coffee", "cocoa", "beef", "소고기", "쇠고기", "농산물", "farmer",
    "agriculture", "농업", "비료", "fertilizer", "seafood", "수산물", "fish",
    "마약", "fentanyl", "drug", "narcotic", "cocaine", "firearm", "gun", "weapon",
    "총기", "방산", "방위산업", "호위함", "잠수함", "무기", "dx korea",
    "의약품", "제약", "약품", "medicine", "pharma", "롯데", "캐릭터",
    "맛집", "연예", "드라마", "축구", "야구", "증시", "주가", "부동산",
    "채용", "모집", "수주", "보안검색 장비"
]

REJECT_OVERRIDE_KW = [
    "federal register", "final rule", "proposed rule", "입법예고", "행정예고",
    "관보", "고시", "공고", "section 301", "section 232", "uflpa",
    "수출통제", "export control", "sanction", "제재", "반덤핑", "상계관세",
    "anti-dumping", "countervailing"
]

FORBIDDEN_PHRASES = [
    "모니터링 필요", "영향 분석 필요", "대응 필요", "검토 필요", "관련 이슈",
    "주의 필요", "리스크 존재", "본사 관세담당자는",
    "직접 영향은 제한적이나 글로벌 통상 환경"
]


# =========================
# CLASSIFICATION ENGINE
# =========================
def strict_reject(title, source=""):
    text = title + " " + source
    if contains(text, REJECT_OVERRIDE_KW):
        return False
    return contains(text, STRICT_REJECT_KW)


def is_trade_related(title, source=""):
    return contains(title + " " + source, TRADE_KW)


def is_official(title, source=""):
    return contains(title + " " + source, OFFICIAL_KW)


def is_legal(title, source=""):
    text = title + " " + source
    return contains(text, LEGAL_KW) and (contains(text, TRADE_KW) or contains(text, OFFICIAL_KW))


def classify_category(title, source=""):
    text = title + " " + source

    if strict_reject(title, source):
        return "REJECT"

    if is_legal(title, source):
        return "A_LEGAL_OFFICIAL"

    if is_official(title, source) and is_trade_related(title, source):
        return "A_LEGAL_OFFICIAL"

    if contains(text, SEMICON_KW) and contains(text, ["관세", "tariff", "232", "301", "ustr", "수출통제", "export control", "제재", "sanction"]):
        return "B_SEMICONDUCTOR_TARIFF"

    if contains(text, ORIGIN_FTA_KW):
        return "C_ORIGIN_FTA_USMCA"

    if contains(text, EXPORT_CONTROL_KW):
        return "D_EXPORT_CONTROL_SANCTION"

    if contains(text, AD_CVD_KW):
        return "E_AD_CVD_TRADE_REMEDY"

    if contains(text, CUSTOMS_AUDIT_KW):
        return "F_CUSTOMS_AUDIT_VALUATION"

    if contains(text, CBAM_KW) or contains(text, ["supply chain", "공급망", "rare earth", "희토류"]):
        return "G_CBAM_SUPPLY_CHAIN"

    if contains(text, PRODUCT_KW) and is_trade_related(title, source):
        return "H_SAMSUNG_GEO_POLICY"

    if contains(text, GEO_KW) and is_trade_related(title, source):
        return "H_SAMSUNG_GEO_POLICY"

    if is_trade_related(title, source):
        return "I_GENERAL_REFERENCE"

    return "REJECT"


def classify_news_type(title, category):
    if category == "A_LEGAL_OFFICIAL":
        return "LEGAL_OFFICIAL"
    if contains(title, SEMICON_KW):
        return "SEMICONDUCTOR"
    if contains(title, ORIGIN_FTA_KW):
        return "FTA_ORIGIN"
    if contains(title, EXPORT_CONTROL_KW):
        return "EXPORT_CONTROL"
    if contains(title, AD_CVD_KW):
        return "AD_CVD"
    if contains(title, CUSTOMS_AUDIT_KW):
        return "CUSTOMS_AUDIT"
    if contains(title, CBAM_KW):
        return "CBAM"
    t = norm(title)
    if "301" in t:
        return "SECTION_301"
    if "232" in t:
        return "SECTION_232"
    return "GENERAL_TRADE"


def category_priority(category):
    return {cat: i + 1 for i, cat in enumerate(CATEGORY_ORDER)}.get(category, 99)


def score_article(title, source="", summary=""):
    text = title + " " + source + " " + summary
    score = 0

    score += count_hits(text, OFFICIAL_KW) * 500
    score += count_hits(text, LEGAL_KW) * 450
    score += count_hits(text, SEMICON_KW) * 350
    score += count_hits(text, HIGH_VALUE_TERMS()) * 300
    score += count_hits(text, ORIGIN_FTA_KW) * 250
    score += count_hits(text, EXPORT_CONTROL_KW) * 280
    score += count_hits(text, AD_CVD_KW) * 260
    score += count_hits(text, CUSTOMS_AUDIT_KW) * 180
    score += count_hits(text, PRODUCT_KW) * 150
    score += count_hits(text, GEO_KW) * 100

    if "news.google.com" in norm(source) or "news.google.com" in norm(title):
        score -= 50

    if strict_reject(title, source):
        score -= 5000

    return score


def HIGH_VALUE_TERMS():
    return [
        "section 301", "301", "section 232", "232", "additional tariff", "추가관세",
        "tariff", "관세", "uflpa", "forced labor", "customs valuation", "과세가격",
        "rules of origin", "원산지", "usmca", "federal register", "final rule",
        "proposed rule", "입법예고", "행정규칙", "고시", "공고"
    ]


def decision_engine(row):
    title = safe_str(row.get("title", ""))
    source = safe_str(row.get("source", ""))
    summary = safe_str(row.get("summary", ""))

    if not title:
        return {"include": False, "category": "REJECT", "news_type": "REJECT", "priority": 99, "score": 0, "decision": "REJECT_EMPTY"}

    category = classify_category(title, source)
    if category == "REJECT":
        return {"include": False, "category": "REJECT", "news_type": "REJECT", "priority": 99, "score": -9999, "decision": "REJECT_RULE"}

    news_type = classify_news_type(title, category)
    priority = category_priority(category)
    score = 10000 - priority * 500 + score_article(title, source, summary)

    return {
        "include": True,
        "category": category,
        "news_type": news_type,
        "priority": priority,
        "score": score,
        "decision": f"INCLUDE_{category}",
    }


# =========================
# ISSUE CLUSTERING
# =========================
def issue_cluster_key(title, category="", news_type=""):
    t = norm(title)

    rules = [
        ("USTR_CHIP_TARIFF", ["ustr", "chip"]),
        ("USTR_CHIP_TARIFF", ["반도체", "관세"]),
        ("SECTION_232_SEMICON", ["232", "semiconductor"]),
        ("USMCA_ORIGIN", ["usmca"]),
        ("USMCA_ORIGIN", ["원산지", "멕시코"]),
        ("US_CHINA_TARIFF", ["미국", "중국", "관세"]),
        ("US_CHINA_TARIFF", ["china", "us", "tariff"]),
        ("EXPORT_CONTROL_CHINA", ["수출통제", "중국"]),
        ("EXPORT_CONTROL_CHINA", ["export", "control", "china"]),
        ("AD_CVD_STEEL", ["반덤핑", "철강"]),
        ("AD_CVD_STEEL", ["anti", "dumping", "steel"]),
        ("EU_STEEL_CBAM", ["eu", "steel"]),
        ("EU_STEEL_CBAM", ["eu", "철강"]),
        ("ASEAN_CUSTOMS_TRANSIT", ["asean", "customs", "transit"]),
        ("ASEAN_CUSTOMS_TRANSIT", ["아세안", "관세", "환승"]),
        ("MEXICO_EU_FTA", ["mexico", "eu", "tariff"]),
        ("MEXICO_EU_FTA", ["멕시코", "eu", "관세"]),
        ("CHINA_SEMICON_EXPORT", ["중국", "반도체", "수출"]),
        ("CHINA_SEMICON_EXPORT", ["china", "semiconductor", "export"]),
    ]

    for key, terms in rules:
        if all(term in t for term in terms):
            return key

    if category == "A_LEGAL_OFFICIAL":
        return "LEGAL_" + "_".join(t.split()[:8])

    tokens = [x for x in t.split() if len(x) >= 2]
    important = [
        x for x in tokens
        if x in [
            "ustr", "cbp", "wto", "관세청", "federal", "register",
            "tariff", "관세", "origin", "원산지", "fta", "usmca", "반덤핑",
            "수출통제", "semiconductor", "반도체", "mexico", "멕시코", "eu",
            "china", "중국", "vietnam", "베트남", "india", "인도"
        ]
    ]

    if important:
        return "_".join(important[:7])

    return f"{category}_{news_type}_" + t[:50]


def is_similar(a, b, threshold=0.84):
    a = norm(a)
    b = norm(b)
    if not a or not b:
        return False
    if a[:40] == b[:40]:
        return True
    return SequenceMatcher(None, a, b).ratio() >= threshold


def deduplicate(df):
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df["url_norm"] = df["url"].fillna("").astype(str).str.strip()
    df = df.drop_duplicates(subset=["url_norm"], keep="first")

    selected = []
    titles = []
    for _, row in df.sort_values(["priority", "score"], ascending=[True, False]).iterrows():
        title = safe_str(row["title"])
        if any(is_similar(title, old) for old in titles):
            continue
        selected.append(row)
        titles.append(title)

    return pd.DataFrame(selected).drop(columns=["url_norm"], errors="ignore") if selected else df.head(0)


# =========================
# SELECTION ENGINE
# =========================
def select_representatives(df, top_n=30):
    df = df.loc[:, ~df.columns.duplicated()].copy()
    df = df.sort_values(["priority", "score"], ascending=[True, False])

    selected_rows = []
    selected_issues = set()
    category_counts = {k: 0 for k in CATEGORY_QUOTA}
    source_counts = {}
    korea_count = 0

    for _, row in df.iterrows():
        category = row["category"]
        issue = row["issue_key"]
        source = safe_str(row.get("source", ""))[:50]
        country = safe_str(row.get("country_rule", ""))

        if category not in CATEGORY_QUOTA:
            continue
        if category_counts[category] >= CATEGORY_QUOTA[category]:
            continue
        if issue in selected_issues:
            continue
        if source and source_counts.get(source, 0) >= MAX_SOURCE:
            continue
        if "한국" in country and korea_count >= MAX_KOREA:
            continue

        selected_rows.append(row)
        selected_issues.add(issue)
        category_counts[category] += 1
        source_counts[source] = source_counts.get(source, 0) + 1
        if "한국" in country:
            korea_count += 1

        if len(selected_rows) >= top_n:
            break

    if len(selected_rows) < top_n:
        for _, row in df.iterrows():
            issue = row["issue_key"]
            source = safe_str(row.get("source", ""))[:50]

            if issue in selected_issues:
                continue
            if source and source_counts.get(source, 0) >= MAX_SOURCE:
                continue

            selected_rows.append(row)
            selected_issues.add(issue)
            source_counts[source] = source_counts.get(source, 0) + 1

            if len(selected_rows) >= top_n:
                break

    return pd.DataFrame(selected_rows).reset_index(drop=True)


# =========================
# COUNTRY / AGENCY
# =========================
def extract_country(title):
    t = norm(title)
    mapping = [
        ("united states", "미국"), ("u.s", "미국"), ("usa", "미국"), ("us", "미국"), ("미국", "미국"),
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
        ("asean", "아세안"), ("아세안", "아세안"),
        ("uae", "UAE"), ("dubai", "UAE"),
    ]

    found = []
    for k, v in mapping:
        if k in t and v not in found:
            found.append(v)
    return ", ".join(found[:2]) if found else "글로벌"


def extract_agency(title, source=""):
    t = norm(title + " " + source)
    mapping = [
        ("ustr", "USTR"),
        ("cbp", "CBP"),
        ("customs and border protection", "CBP"),
        ("federal register", "Federal Register"),
        ("관세청", "관세청"),
        ("wto", "WTO"),
        ("bis", "BIS"),
        ("mofcom", "MOFCOM"),
        ("european commission", "EU Commission"),
        ("eu commission", "EU Commission"),
        ("official journal", "EU Official Journal"),
        ("taxud", "EU TAXUD"),
        ("commerce department", "U.S. Department of Commerce"),
        ("상무부", "U.S. Department of Commerce"),
        ("customs", "세관"),
        ("세관", "세관"),
        ("ministry", "정부기관"),
        ("법제처", "법제처"),
        ("관보", "관보"),
        ("reuters", "Reuters"),
    ]
    for k, v in mapping:
        if k in t:
            return v
    return "N/A"


def importance_by_category(category):
    if category in ["A_LEGAL_OFFICIAL", "B_SEMICONDUCTOR_TARIFF", "C_ORIGIN_FTA_USMCA", "D_EXPORT_CONTROL_SANCTION", "E_AD_CVD_TRADE_REMEDY"]:
        return "상"
    if category in ["F_CUSTOMS_AUDIT_VALUATION", "G_CBAM_SUPPLY_CHAIN", "H_SAMSUNG_GEO_POLICY"]:
        return "중"
    return "하"


def risk_by_category(category):
    if category in ["A_LEGAL_OFFICIAL", "B_SEMICONDUCTOR_TARIFF", "D_EXPORT_CONTROL_SANCTION", "E_AD_CVD_TRADE_REMEDY"]:
        return "상"
    if category in ["C_ORIGIN_FTA_USMCA", "F_CUSTOMS_AUDIT_VALUATION", "G_CBAM_SUPPLY_CHAIN"]:
        return "중"
    return "하"


# =========================
# WEB BODY
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
        time.sleep(random.uniform(0.4, 1.0))
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
# ANALYSIS ENGINE
# =========================
def detect_product_area(title):
    areas = []
    if contains(title, SEMICON_KW):
        areas.append("Semiconductor/Component")
    if contains(title, ["mobile", "smartphone", "스마트폰", "휴대폰", "galaxy"]):
        areas.append("Mobile")
    if contains(title, ["전자", "electronics", "가전", "consumer electronics", "tv"]):
        areas.append("Consumer Electronics")
    if contains(title, ["network", "네트워크"]):
        areas.append("Network Equipment")
    if contains(title, ["display", "디스플레이"]):
        areas.append("Display")
    return ", ".join(areas[:2]) if areas else "Mobile/CE/Component 공통"


def detect_geo_focus(title):
    geos = []
    checks = [
        ("SEV/SEVT", ["vietnam", "베트남"]),
        ("SIEL", ["india", "인도"]),
        ("SAMEX", ["mexico", "멕시코"]),
        ("중국법인", ["china", "중국"]),
        ("북미 판매법인", ["united states", "usa", "u.s", "미국", "us"]),
        ("EU 판매법인", ["eu", "europe", "유럽"]),
        ("한국 본사/생산", ["korea", "한국"]),
    ]
    for label, keys in checks:
        if contains(title, keys):
            geos.append(label)
    return ", ".join(geos[:3]) if geos else "SEV/SEVT, SIEL, SAMEX"


def fallback_summary(title, body="", original_summary=""):
    if original_summary and len(original_summary) > 20:
        return limit_sentences(original_summary, 2, 420)
    if body and len(body) > 120:
        return limit_sentences(trim(body, 350), 2, 420)
    return f"{title} 관련 관세·통상 정책 이슈입니다."


def fallback_analysis(title, news_type, category):
    product = detect_product_area(title)
    geo = detect_geo_focus(title)

    if category == "A_LEGAL_OFFICIAL":
        return (
            f"해당 공식 정책·법령 이슈는 {geo} 생산품의 HS·원산지·FTA·수출통제 내부통제 기준 변경 여부를 즉시 판정해야 하는 사안입니다. "
            f"{product} 적용 품목과 시행일을 기준으로 통관 신고 기준, 증빙 보관, 법인·관세사 업무지침 변경 여부를 확정해야 합니다."
        )

    if category == "B_SEMICONDUCTOR_TARIFF":
        return (
            f"반도체 관세·규제 변화는 {geo}의 Semiconductor/Component 품목 HS별 관세원가와 북미/EU 판매법인 가격전가에 직접 반영됩니다. "
            "Section 232/301 및 수출통제 병행 적용 가능성을 전제로 원산지·FTA·EAR 판정을 동시에 재점검해야 합니다."
        )

    if category == "C_ORIGIN_FTA_USMCA":
        return (
            f"FTA·원산지 기준 변화는 {geo} 생산품의 협정세율 적용과 사후검증 리스크를 직접 좌우합니다. "
            f"{product} BOM상 역외산 핵심부품 비중이 높으면 원산지 충족률과 FTA 적용 가능성이 즉시 변동됩니다."
        )

    if category == "D_EXPORT_CONTROL_SANCTION":
        return (
            "수출통제·제재 변화는 Semiconductor/Component와 Network Equipment 거래의 최종사용자·최종용도 심사를 강화합니다. "
            "SEV/SEVT·SIEL·SAMEX 출하품은 EAR·제재리스트·재수출 규정 적용 여부를 거래 전 단계에서 차단해야 합니다."
        )

    if category == "E_AD_CVD_TRADE_REMEDY":
        return (
            f"반덤핑·상계관세 조치는 {product} 관련 원재료·부품의 조달비용과 우회수출 판정 리스크를 높입니다. "
            "SEV/SEVT·SIEL·SAMEX 생산품에 투입되는 대상 품목의 HS·원산지·거래가격 연결성을 즉시 확인해야 합니다."
        )

    if category == "F_CUSTOMS_AUDIT_VALUATION":
        return (
            f"세관 집행 강화는 {product} 제품의 HS 분류, 과세가격, 원산지 증빙, FTA 적용 신고의 정합성을 직접 겨냥합니다. "
            "SEV/SEVT·SIEL·SAMEX 신고자료와 ERP Invoice·계약가격 간 불일치가 추징 포인트가 될 수 있습니다."
        )

    if category == "G_CBAM_SUPPLY_CHAIN":
        return (
            "CBAM·공급망 정책은 EU향 제품의 원재료 출처, 탄소자료, 원산지 증빙을 결합해 요구할 가능성이 높습니다. "
            "SEV/SEVT·SIEL·SAMEX 생산품의 HS별 원재료 출처와 증빙 확보 수준을 EU 판매법인 기준으로 정렬해야 합니다."
        )

    return (
        f"{geo} 관련 통상정책 변화가 {product} 제품의 HS·원산지·FTA·수출통제 적용 여부와 연결되는지 1차 판정해야 합니다. "
        "직접 관련 품목이 확인되면 생산거점별 관세율·과세가격·증빙자료를 즉시 재산출합니다."
    )


def fallback_action(title, news_type, category):
    product = detect_product_area(title)
    geo = detect_geo_focus(title)

    if category == "A_LEGAL_OFFICIAL":
        return (
            f"① 시행일·적용대상 HS를 확정하고 ② {geo} 생산거점별 원산지·FTA·수출통제 영향표를 업데이트하며 ③ 법인·관세사 신고지침과 증빙 보관 체크리스트를 개정합니다."
        )

    if category == "B_SEMICONDUCTOR_TARIFF":
        return (
            f"① 미국/EU향 Semiconductor HS Mapping을 재점검하고 ② {geo} 생산거점별 관세율·원산지·FTA 적용세율을 재산출하며 ③ EAR/수출통제 대상 여부와 가격전가 영향을 사업부에 공유합니다."
        )

    if category == "C_ORIGIN_FTA_USMCA":
        return (
            f"① {geo} 생산품의 BOM 기준 원산지 충족률을 재산출하고 ② FTA 적용 가능 HS를 분리하며 ③ 공급업체 원산지확인서·제조공정 증빙을 사후검증 패키지로 정리합니다."
        )

    if category == "D_EXPORT_CONTROL_SANCTION":
        return (
            "① 거래상대방·최종사용자·목적지 국가를 수출통제 리스트와 대조하고 ② EAR·재수출 규정 적용 여부를 판정하며 ③ 반도체·네트워크 출하 승인 기준을 법무·영업과 업데이트합니다."
        )

    if category == "E_AD_CVD_TRADE_REMEDY":
        return (
            "① 대상 품목 HS와 공급처 원산지를 매핑하고 ② SEV/SEVT·SIEL·SAMEX 투입 부품과 연결성을 비교하며 ③ 반덤핑 관세율 반영 시 대체 공급처·가격조건을 구매부서와 재산출합니다."
        )

    if category == "F_CUSTOMS_AUDIT_VALUATION":
        return (
            f"① {geo} 최근 신고 건의 HS·과세가격·원산지·FTA 적용 내역을 샘플링하고 ② ERP Invoice와 신고금액·수량 불일치를 정리하며 ③ 관세사별 소명자료 패키지를 준비합니다."
        )

    if category == "G_CBAM_SUPPLY_CHAIN":
        return (
            "① EU향 대상 HS를 식별하고 ② 원재료 원산지·탄소자료 확보 수준을 점검하며 ③ SEV/SEVT·SIEL·SAMEX 공급망별 증빙 공백을 EU 판매법인에 공유합니다."
        )

    return (
        f"① {geo} 관련 품목의 HS·원산지·FTA 적용 여부를 1차 분류하고 ② 수출통제 대상 여부를 확인하며 ③ 직접 영향 품목만 GTI 후속 과제로 등록합니다."
    )


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
    t = clean_text(text)
    if not t or len(t) < min_len:
        return True
    if any(p in t for p in FORBIDDEN_PHRASES):
        return True
    if t.count("삼성전자") > 3:
        return True
    return False


def analyze_with_ai(title, body, news_type, category, decision):
    if not USE_AI or client is None:
        return {}

    prompt = f"""
너는 삼성전자 본사 관세전략 담당 임원이다.
아래 뉴스 제목과 본문을 기준으로 삼성전자 관세·통상 업무 관점의 GTI 보고서용 분석을 작성한다.

[선정 카테고리]
{category}

[뉴스 유형]
{news_type}

[선정 사유]
{decision}

[제목]
{title}

[본문 일부]
{body[:2200]}

[작성 기준]
- 2문장 이내
- 실행지시형
- HS/원산지/FTA/수출통제 포함
- 삼성 생산거점(SEV/SEVT/SIEL/SAMEX) 언급
- 관세전문가 어조
- 일반론 금지
- "모니터링 필요", "검토 필요", "대응 필요" 금지
- 법령/공식문서는 시행일, 적용대상, 내부통제 변경 여부 우선
- 반도체/301/232는 HS·관세율 시나리오·EAR·북미 판매법인 영향 우선
- FTA/원산지는 BOM·원산지 충족률·사후검증 증빙 우선
- 수출통제는 최종사용자·최종용도·재수출 규정 우선

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


def enforce_quality(text):
    text = limit_sentences(text, 2, 650)
    if any(p in text for p in FORBIDDEN_PHRASES):
        return ""
    required = ["HS", "원산지", "FTA", "수출통제", "EAR", "SEV", "SEVT", "SIEL", "SAMEX", "관세", "과세가격"]
    if not any(k in text for k in required):
        return ""
    if len(text) < 70:
        return ""
    return text


# =========================
# EXCEL
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
        "Date": 18, "Headline": 55, "importance": 10, "URL": 30,
        "source": 22, "last_checked": 18, "Summary": 58,
        "AI Analysis": 78, "Action Plan": 78, "Country": 18,
        "agency": 24, "risk": 8, "score": 10, "priority": 8,
        "category": 26, "decision": 28, "news_type": 18, "issue_key": 28,
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
    print("🚀 GTI STEP4 STRUCTURAL FINAL v12.0 START")
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
    df["summary"] = df["summary"].apply(safe_str)

    df = df[(df["title"] != "") & (df["title"] != "0")].copy()
    print(f"[TITLE OK] {len(df)} rows")

    for col in ["include", "category", "news_type", "priority", "score", "decision", "issue_key"]:
        if col in df.columns:
            df = df.drop(columns=[col])

    decisions = df.apply(decision_engine, axis=1, result_type="expand")
    df = pd.concat([df.reset_index(drop=True), decisions.reset_index(drop=True)], axis=1)
    df = df.loc[:, ~df.columns.duplicated()].copy()

    print(f"[INCLUDE] {int(df['include'].sum())} rows")
    print(f"[REJECT] {int((~df['include']).sum())} rows")

    df = df[df["include"] == True].copy()

    df["country_rule"] = df["title"].apply(extract_country)
    df["agency_rule"] = df.apply(lambda r: extract_agency(r["title"], r.get("source", "")), axis=1)
    df["issue_key"] = df.apply(lambda r: issue_cluster_key(r["title"], r["category"], r["news_type"]), axis=1)
    df["importance"] = df["category"].apply(importance_by_category)
    df["risk_rule"] = df["category"].apply(risk_by_category)

    df = deduplicate(df)
    print(f"[DEDUP] {len(df)} rows")

    top = select_representatives(df, TOP_N)
    print(f"[TOP] {len(top)} rows")

    if len(top) > 0:
        print("[CATEGORY MIX]")
        for k, v in top["category"].value_counts().items():
            print(f" - {k}: {v}")

    records = []
    last_checked = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for idx, row in top.iterrows():
        title = safe_str(row["title"])
        url = safe_str(row["url"])
        news_type = safe_str(row.get("news_type", "GENERAL_TRADE"))
        category = safe_str(row.get("category", ""))
        decision = safe_str(row.get("decision", ""))

        print(f"[AI {idx + 1}/{len(top)}] {category} | {title[:70]}")

        body = fetch_body(url)
        ai = analyze_with_ai(title, body, news_type, category, decision)

        summary = clean_text(ai.get("summary", ""))
        analysis = clean_text(ai.get("ai_analysis", ""))
        action = clean_text(ai.get("action_plan", ""))
        country = clean_text(ai.get("country", ""))
        agency = clean_text(ai.get("agency", ""))
        risk = clean_text(ai.get("risk", ""))

        if bad_ai_text(summary, min_len=35):
            summary = fallback_summary(title, body, row.get("summary", ""))

        analysis = enforce_quality(analysis)
        action = enforce_quality(action)

        if not analysis:
            analysis = fallback_analysis(title, news_type, category)

        if not action:
            action = fallback_action(title, news_type, category)

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
            "priority": int(row.get("priority", 99)),
            "category": category,
            "decision": decision,
            "news_type": news_type,
            "issue_key": safe_str(row.get("issue_key", "")),
        })

    out = pd.DataFrame(records)

    final_cols = [
        "Date", "Headline", "importance", "URL", "source", "last_checked",
        "Summary", "AI Analysis", "Action Plan",
        "Country", "agency", "risk",
        "score", "priority", "category", "decision", "news_type", "issue_key"
    ]

    out = out[final_cols]
    write_excel(out, OUTPUT_DAILY)

    if os.path.exists(OUTPUT_CUMUL):
        old = pd.read_excel(OUTPUT_CUMUL)
        old = old.loc[:, ~old.columns.duplicated()].copy()
        total = pd.concat([old, out], ignore_index=True)
        if "URL" in total.columns:
            total = total.drop_duplicates(subset=["URL"], keep="last")
        else:
            total = total.drop_duplicates(subset=["Headline"], keep="last")
    else:
        total = out.copy()

    write_excel(total, OUTPUT_CUMUL)

    print("===================================")
    print("✅ GTI STEP4 STRUCTURAL FINAL v12.0 COMPLETE")
    print(f"📁 DAILY : {OUTPUT_DAILY}")
    print(f"📁 CUMUL : {OUTPUT_CUMUL}")
    print("===================================")


if __name__ == "__main__":
    main()
