# -*- coding: utf-8 -*-
"""
GTI STEP4-2 NEWS AI ANALYSIS - GUARDRAIL v4.1

Purpose
- Read ONLY C:\\Temp\\3-2.news_summary.xlsx.
- Prevent stale/irrelevant items from entering executive mail.
- Do NOT force Top30: select only rows passing quality gates.
- Preserve reliable links via BestLinkURL / GoogleURL.
- Also write legacy C:\\Temp\\4.news_ai_analysis.xlsx so Step5 cannot reuse an old stale file.

Hard gates
- reject old news older than GTI_STEP4_NEWS_MAX_AGE_HOURS (default 72h)
- reject invalid URLs including fonts.googleapis / analytics / ad URLs
- reject webinar/seminar/tender/training/event notices
- reject weak Samsung relevance unless a very strong official trade-control issue
"""
from __future__ import annotations

import os
import re
import json
import ssl
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import pandas as pd

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_ARTICLE_FILE = BASE_DIR / "3-2.news_article_summary.xlsx"
INPUT_SUMMARY_FILE = BASE_DIR / "3-2.news_summary.xlsx"
INPUT_FILE = Path(os.getenv("GTI_STEP4_NEWS_INPUT", str(INPUT_ARTICLE_FILE)))
OUT_SUMMARY = BASE_DIR / "4-2.news_ai_summary.xlsx"
OUT_CUMULATIVE = BASE_DIR / "4-2.news_ai_cumulative.xlsx"
OUT_AUDIT = BASE_DIR / "4-2.news_ai_audit_candidates.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-2.news_ai_excluded.xlsx"
OUT_LEGACY = BASE_DIR / "4.news_ai_analysis.xlsx"
GOOGLE_RESOLVE_CACHE_FILE = BASE_DIR / "google_news_url_cache.csv"

MAX_AGE_HOURS = int(os.getenv("GTI_STEP4_NEWS_MAX_AGE_HOURS", "72"))
TOP_N_MAX = int(os.getenv("GTI_STEP4_TOP_N_MAX", "50"))
NEWS_TARGET_MIN = int(os.getenv("GTI_STEP4_NEWS_TARGET_MIN", "30"))
NEWS_TARGET_MAX = int(os.getenv("GTI_STEP4_NEWS_TARGET_MAX", str(TOP_N_MAX)))
MIN_SELECT_SCORE = int(os.getenv("GTI_STEP4_MIN_SELECT_SCORE", "75"))
POLICY_WATCH_MIN_SCORE = int(os.getenv("GTI_STEP4_POLICY_WATCH_MIN_SCORE", "75"))
NEWS_EXPAND_MIN_SCORE = int(os.getenv("GTI_STEP4_NEWS_EXPAND_MIN_SCORE", "55"))
GOOGLE_RESOLVE_ENABLED = os.getenv("GTI_GOOGLE_NEWS_RESOLVE", "1").strip().upper() not in {"0", "N", "NO", "FALSE"}
GOOGLE_RESOLVE_TIMEOUT = int(os.getenv("GTI_GOOGLE_NEWS_RESOLVE_TIMEOUT", "25"))
GOOGLE_RESOLVE_INTERVAL = float(os.getenv("GTI_GOOGLE_NEWS_RESOLVE_INTERVAL", "0.3"))

BAD_URL_PATTERNS = [
    "google-analytics.com", "googletagmanager.com", "doubleclick.net", "analytics.js", "gtag/js",
    "googlesyndication.com", "googleadservices.com", "google.com/pagead", "fonts.googleapis.com",
    "fonts.gstatic.com", "googleusercontent.com", "static.xx.fbcdn", "pixel", "beacon",
]

GOOGLE_RESOLVE_CACHE: dict[str, str] = {}
GOOGLE_RESOLVE_CACHE_LOADED = False

EVENT_NOISE_TERMS = [
    "webinar", "seminar", "conference", "summit", "workshop", "training", "education", "lecture",
    "forum", "symposium", "registration", "tender", "call for tender", "rfp", "expo", 
    "opening ceremony", "ceremony", "award", "recruit", "invitation", "apply now", "join the upcoming",
    "웨비나", "세미나", "컨퍼런스", "서밋", "워크숍", "교육", "강의", "설명회", "간담회", "포럼",
    "입찰", "공모", "행사", "박람회", "전시회", "수상", "시상", "모집", "참가신청", 
    "cms summit", "aeo vs seo", "seo",
]

LOW_VALUE_TERMS = [
    "수입차", "중국차", "자동차 시장", "건강기능식품", "고등어", "오징어", "냉동", "맛집", "여행",
    "관광", "스포츠", "야구", "축구", "주가", "부동산", "아파트", "범죄", "마약", "밀수범",
    "politics", "election", "war", "ceasefire", "opinion", "editorial", "celebrity",
    "정치", "선거", "전쟁", "휴전", "외교", "사설", "칼럼", "기자회견",
]

GENERAL_ECONOMY_TERMS = [
    "gdp", "growth", "economy", "economic", "investment", "market", "business", "trade volume",
    "경제", "성장률", "투자", "시장", "산업동향", "무역동향", "수출 증가", "수출 감소", "수출 호조", "환율", "원화", "원달러", "원ㆍ달러", "환율", "외환", "투기적", "금융시장", "금리", "f4",
]

FINANCIAL_INDUSTRY_NOISE_TERMS = [
    "stock", "stocks", "share price", "shares", "futures", "perpetual futures", "crypto", "coin",
    "bitcoin", "listing", "listed", "earnings", "profit", "sales outlook", "market outlook",
    "beneficiary", "rally", "주가", "증시", "선물", "무기한 선물", "코인", "상장", "실적",
    "영업이익", "매출", "수혜", "호황", "장비시장", "시장동향", "전망", "랠리",
]

REAL_EVENT_NOISE_TERMS = [
    "webinar", "seminar", "workshop", "training", "conference registration",
    "call for tender", "tender", "rfp", "recruit", "recruitment",
    "채용", "공무직", "합격자", "면접전형", "입찰", "설명회", "교육", "세미나", "웨비나",
]

SAMSUNG_GENERAL_NOISE_TERMS = [
    "brand value", "brand ranking", "share price", "stock", "stocks", "earnings",
    "strategy meeting", "market cap", "analyst", "profit outlook", "sales outlook",
    "brand", "investment", "investor", "money", "logistics hub",
    "브랜드", "브랜드 가치", "브랜드 순위", "주가", "증시", "실적", "전략회의", "글로벌 전략회의",
    "시가총액", "주식", "큰 돈", "돈 벌", "투자", "수혜주", "전망", "칼럼",
    "고환율", "물류 거점", "기술 수출", "초음파 기술 수출",
]

CONCRETE_TRADE_POLICY_TERMS = [
    "tariff", "customs duty", "import duty", "quota", "section 301", "section 232",
    "anti-dumping", "anti dumping", "antidumping", "countervailing", "ad/cvd",
    "safeguard", "forced labor", "uflpa", "export control", "entity list",
    "cbam", "carbon border", "rules of origin", "hs code",
    "관세", "쿼터", "반덤핑", "상계관세", "무역구제", "강제노동", "수출통제",
    "원산지", "품목분류", "통관",
]

TRADE_POLICY_DIRECT_TERMS = [
    "tariff", "tariffs", "customs duty", "import duty", "quota", "duty-free quota",
    "section 301", "section 232", "anti-dumping", "anti dumping", "antidumping",
    "countervailing", "ad/cvd", "safeguard", "forced labor", "uflpa",
    "export control", "entity list", "cbam", "carbon border", "fta", "rules of origin",
    "hs code", "classification", "clearance", "declaration",
    "관세", "관세율", "쿼터", "무관세", "반덤핑", "상계관세", "무역구제",
    "세이프가드", "강제노동", "수출통제", "수출 통제", "전략물자", "탄소국경", "원산지",
    "품목분류", "통관", "신고",
]

SAMSUNG_EXACT_TERMS = [
    "samsung", "samsung electronics", "samsung sdi", "samsung display", "삼성", "삼성전자", "삼성sdi", "삼성디스플레이", "삼전",
]
SEMICONDUCTOR_TERMS = ["semiconductor", "chip", "chips", "hbm", "memory", "반도체", "칩", "메모리", "ai chip", "ai chips"]
MOBILE_TERMS = ["smartphone", "mobile phone", "handset", "galaxy", "스마트폰", "휴대폰", "갤럭시"]
BATTERY_TERMS = ["battery", "batteries", "ev battery", "배터리", "이차전지"]
DISPLAY_TERMS = ["display", "oled", "디스플레이"]
PRODUCT_TERMS = SEMICONDUCTOR_TERMS + MOBILE_TERMS + BATTERY_TERMS + DISPLAY_TERMS

TOPIC_RULES = [
    ("EXPORT_CONTROL", ["export control", "export controls", "entity list", "bis", "denied persons", "수출통제", "전략물자", "제재", "ai chip", "ai chips"]),
    ("AD_CVD", ["anti-dumping", "anti dumping", "antidumping", "countervailing", "ad/cvd", "cvd", "반덤핑", "상계관세", "무역구제"]),
    ("CBAM_CARBON", ["cbam", "carbon border", "carbon border adjustment", "탄소국경"]),
    ("ORIGIN_FTA", ["fta", "cepa", "usmca", "rules of origin", "origin", "원산지", "자유무역협정"]),
    ("HS_CLASSIFICATION", ["hs code", "classification", "tariff classification", "품목분류", "hs코드"]),
    ("TARIFF", ["section 301", "301조", "section 232", "232조", "reciprocal tariff", "tariff", "tariffs", "customs duty", "import duty", "관세", "관세율", "추가관세", "상호관세"]),
    ("CUSTOMS", ["customs", "clearance", "declaration", "통관", "세관", "관세청"]),
]

MUST_KEEP_POLICY_TERMS = [
    "section 301", "301조", "section 232", "232조", "reciprocal tariff",
    "tariff cap", "tariff ceiling", "tariff-rate quota", "tariff rate quota",
    "tariff quota", "duty-free quota", "duty free quota", "무관세 쿼터",
    "관세상한", "관세 쿼터", "anti-dumping", "anti dumping", "antidumping",
    "countervailing", "countervailing duty", "countervailing duties", "ad/cvd",
    "safeguard", "steel safeguard", "steel overcapacity", "steel quota",
    "steel tariff", "aluminum tariff", "forced labor", "uflpa",
    "export control", "entity list", "cbam", "carbon border",
    "반덤핑", "상계관세", "무역구제", "세이프가드", "강제노동", "수출통제",
]

MUST_KEEP_POLICY_COMBOS = [
    (["steel", "철강"], ["quota", "쿼터", "safeguard", "tariff", "관세", "무관세"]),
    (["aluminum", "알루미늄"], ["quota", "쿼터", "safeguard", "tariff", "관세", "무관세"]),
    (["battery", "배터리"], ["tariff", "관세", "301", "section 301"]),
    (["semiconductor", "chip", "반도체", "칩"], ["export control", "entity list", "tariff", "관세", "수출통제"]),
]

TOPIC_KR = {
    "EXPORT_CONTROL": "수출통제",
    "AD_CVD": "반덤핑/상계관세",
    "CBAM_CARBON": "CBAM",
    "ORIGIN_FTA": "FTA/원산지",
    "HS_CLASSIFICATION": "HS/품목분류",
    "TARIFF": "관세정책",
    "CUSTOMS": "통관/세관",
    "TRADE_GENERAL": "무역일반",
}

OUTPUT_COLS = [
    "rank", "Date", "Headline", "URL", "GoogleURL", "OriginalURLCandidate", "BestLinkURL", "URL_Quality",
    "Country", "Agency", "Publisher", "priority_group", "mail_section", "selected", "Risk", "final_score",
    "topic", "topic_score", "samsung_impact", "samsung_impact_score", "subsidiary_score", "action_score", "urgency_score",
    "topic_keyword", "topic_reason", "issue_type", "cluster_key", "RegulationRelated", "RegulationTransferType",
    "affected_subsidiary", "affected_subsidiaries", "affected_products", "subsidiary_products", "subsidiary_reason",
    "impact_production_subsidiaries", "impact_sales_subsidiaries", "impact_products", "fta_impact", "export_control_impact",
    "hs_impact", "tariff_impact", "RequiredAction", "ActionOwner", "ExecutiveMessage", "samsung_score", "samsung_reason",
    "Summary", "AI Analysis", "Action Plan", "KeywordMatches", "SelectReason", "RejectReason", "Source", "SourceFile",
    "original_url", "article_body", "ClusterHeadlines", "article_extract_status", "article_source_type",
    "effective_date_hint", "change_detail_hint", "hs_hint", "tariff_rate_hint", "last_checked",
]

LEGACY_COLS = [
    "No", "Content Type", "Mail Group", "Samsung Impact", "Affected Subsidiary", "Impact Reason", "Date", "Headline",
    "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Risk", "Importance Score", "Priority Group",
    "Issue", "Cluster", "URL", "Source", "Source File",
]


def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")


def clean(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def contains_any(text: str, terms: list[str]) -> bool:
    t = str(text or "").lower()
    return any(term.lower() in t for term in terms)


def has_policy_combo(text: str) -> bool:
    t = str(text or "").lower()
    for left_terms, right_terms in MUST_KEEP_POLICY_COMBOS:
        if contains_any(t, left_terms) and contains_any(t, right_terms):
            return True
    return False


def is_must_keep_policy_news(text: str, topic: str) -> bool:
    if topic not in {"EXPORT_CONTROL", "AD_CVD", "CBAM_CARBON", "ORIGIN_FTA", "HS_CLASSIFICATION", "TARIFF", "CUSTOMS"}:
        return False
    return contains_any(text, MUST_KEEP_POLICY_TERMS) or has_policy_combo(text)


def has_direct_trade_policy_signal(text: str) -> bool:
    return contains_any(text, TRADE_POLICY_DIRECT_TERMS) or has_policy_combo(text)


def is_real_event_noise(text: str) -> bool:
    return contains_any(text, REAL_EVENT_NOISE_TERMS)


def is_samsung_general_noise(text: str, direct_policy_signal: bool, must_keep_policy: bool) -> bool:
    if direct_policy_signal or must_keep_policy:
        if contains_any(text, CONCRETE_TRADE_POLICY_TERMS):
            return False
    return contains_any(text, SAMSUNG_GENERAL_NOISE_TERMS)


def is_financial_or_industry_noise(text: str, direct_policy_signal: bool, must_keep_policy: bool) -> bool:
    if direct_policy_signal or must_keep_policy:
        return False
    return contains_any(text, FINANCIAL_INDUSTRY_NOISE_TERMS)


def is_bilateral_industry_noise(text: str, direct_policy_signal: bool, must_keep_policy: bool) -> bool:
    if direct_policy_signal or must_keep_policy:
        return False
    has_bilateral = contains_any(text, ["summit", " 정상회담", "정상 회담", "cooperation", "협력", "economic security", "경제 안보"])
    has_industry = contains_any(text, PRODUCT_TERMS)
    return has_bilateral and has_industry


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df.loc[:, ~pd.Index(df.columns).duplicated()]


def parse_dt(v):
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt):
            return pd.NaT
        if getattr(dt, "tzinfo", None) is not None:
            dt = dt.tz_convert(None)
        return dt
    except Exception:
        return pd.NaT


def normalize_title(s: str) -> str:
    t = clean(s).lower()
    t = re.sub(r"\[[^\]]+\]|\([^\)]+\)", " ", t)
    t = re.sub(r"[^a-z0-9가-힣]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def is_generic_google_main(url: str) -> bool:
    u = clean(url).lower().rstrip("/")
    if u in {"https://news.google.com", "http://news.google.com", "https://www.google.com", "http://www.google.com"}:
        return True
    try:
        p = urlparse(u)
        if "news.google" in p.netloc and p.path in {"", "/", "/home"}:
            return True
        if "google." in p.netloc and p.path in {"", "/", "/search"}:
            return True
    except Exception:
        pass
    return False


def is_google_article_redirect(url: str) -> bool:
    u = clean(url).lower()
    if not u.startswith(("http://", "https://")):
        return False
    p = urlparse(u)
    return "news.google" in p.netloc and ("/rss/articles/" in p.path or "/articles/" in p.path)


def safe_url(url: str) -> str:
    u = clean(url).replace("\r", "").replace("\n", "").strip()
    if not u:
        return ""
    # Encode spaces and non-ASCII safely without damaging normal URL separators.
    return quote(unquote(u), safe=":/?#[]@!$&'()*+,;=%")


def google_news_token(url: str) -> str:
    try:
        p = urlparse(url)
        parts = [x for x in p.path.split("/") if x]
        if len(parts) >= 2 and parts[-2] in {"articles", "read"}:
            return parts[-1]
    except Exception:
        pass
    return ""


def load_google_resolve_cache() -> None:
    global GOOGLE_RESOLVE_CACHE_LOADED
    if GOOGLE_RESOLVE_CACHE_LOADED:
        return
    GOOGLE_RESOLVE_CACHE_LOADED = True
    if not GOOGLE_RESOLVE_CACHE_FILE.exists():
        return
    try:
        df = pd.read_csv(GOOGLE_RESOLVE_CACHE_FILE)
        for _, row in df.iterrows():
            google_url = safe_url(row.get("google_url", ""))
            resolved_url = safe_url(row.get("resolved_url", ""))
            if google_url and resolved_url:
                GOOGLE_RESOLVE_CACHE[google_url] = resolved_url
    except Exception:
        return


def save_google_resolve_cache() -> None:
    try:
        rows = [
            {"google_url": google_url, "resolved_url": resolved_url, "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            for google_url, resolved_url in GOOGLE_RESOLVE_CACHE.items()
            if google_url and resolved_url
        ]
        if rows:
            pd.DataFrame(rows).drop_duplicates(subset=["google_url"], keep="last").to_csv(
                GOOGLE_RESOLVE_CACHE_FILE, index=False, encoding="utf-8-sig"
            )
    except Exception:
        return


def fetch_google_decode_params(token: str) -> tuple[str, str]:
    ctx = ssl.create_default_context()
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/129 Safari/537.36",
    }
    for prefix in ("https://news.google.com/articles/", "https://news.google.com/rss/articles/"):
        req = urllib.request.Request(prefix + token, headers=headers)
        with urllib.request.urlopen(req, timeout=GOOGLE_RESOLVE_TIMEOUT, context=ctx) as resp:
            html = resp.read().decode("utf-8", "ignore")
        sig = re.search(r'data-n-a-sg="([^"]+)"', html)
        ts = re.search(r'data-n-a-ts="([^"]+)"', html)
        if sig and ts:
            return sig.group(1), ts.group(1)
    return "", ""


def decode_google_news_token(token: str, signature: str, timestamp: str) -> str:
    endpoint = "https://news.google.com/_/DotsSplashUi/data/batchexecute"
    payload = [
        "Fbv4je",
        (
            '["garturlreq",[["X","X",["X","X"],null,null,1,1,"US:en",null,1,'
            'null,null,null,null,null,0,1],"X","X",1,[1,1,1],1,1,null,0,0,null,0],'
            f'"{token}",{timestamp},"{signature}"]'
        ),
    ]
    body = "f.req=" + quote(json.dumps([[payload]], separators=(",", ":")))
    headers = {
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/129 Safari/537.36",
    }
    ctx = ssl.create_default_context()
    req = urllib.request.Request(endpoint, data=body.encode("utf-8"), headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=GOOGLE_RESOLVE_TIMEOUT, context=ctx) as resp:
        text = resp.read().decode("utf-8", "ignore")
    parsed = json.loads(text.split("\n\n", 1)[1])[:-2]
    return json.loads(parsed[0][2])[1]


def resolve_google_news_url(url: str) -> str:
    u = safe_url(url)
    if not GOOGLE_RESOLVE_ENABLED or not is_google_article_redirect(u):
        return ""
    load_google_resolve_cache()
    if u in GOOGLE_RESOLVE_CACHE:
        return GOOGLE_RESOLVE_CACHE[u]
    token = google_news_token(u)
    if not token:
        GOOGLE_RESOLVE_CACHE[u] = ""
        return ""
    for _ in range(2):
        try:
            signature, timestamp = fetch_google_decode_params(token)
            if signature and timestamp:
                resolved = safe_url(decode_google_news_token(token, signature, timestamp))
                if resolved and is_valid_link(resolved) and not is_google_article_redirect(resolved):
                    GOOGLE_RESOLVE_CACHE[u] = resolved
                    save_google_resolve_cache()
                    if GOOGLE_RESOLVE_INTERVAL > 0:
                        time.sleep(GOOGLE_RESOLVE_INTERVAL)
                    return resolved
        except Exception:
            if GOOGLE_RESOLVE_INTERVAL > 0:
                time.sleep(GOOGLE_RESOLVE_INTERVAL)
    GOOGLE_RESOLVE_CACHE[u] = ""
    return ""


def is_valid_link(url: str) -> bool:
    u = safe_url(url)
    if not u.lower().startswith(("http://", "https://")):
        return False
    if is_generic_google_main(u):
        return False
    low = u.lower()
    return not any(x in low for x in BAD_URL_PATTERNS)


def choose_best_link(row: pd.Series, resolve_google: bool = False) -> tuple[str, str]:
    # Prefer real URL. Allow Google article redirect. Never allow script/font/ad URLs.
    for col, status in [
        ("BestLinkURL", "BEST_LINK"),
        ("OriginalURLCandidate", "ORIGINAL_CANDIDATE"),
        ("URL", "URL"),
        ("GoogleURL", "GOOGLE_ARTICLE_REDIRECT"),
        ("original_url", "ORIGINAL_URL"),
    ]:
        v = safe_url(row.get(col, ""))
        if is_valid_link(v):
            if is_google_article_redirect(v) and resolve_google:
                resolved = resolve_google_news_url(v)
                if resolved:
                    return resolved, "GOOGLE_NEWS_RESOLVED"
            if is_google_article_redirect(v):
                return v, "GOOGLE_ARTICLE_REDIRECT"
            return v, status
    return "", "EMPTY_OR_BAD_LINK"


def content_text(row: pd.Series) -> str:
    # Article-facing text only. Do NOT include Step3 metadata like SelectReason or SamsungSignal,
    # because strings such as "samsung=PRODUCTION_COUNTRY" falsely trigger Samsung relevance.
    # Source/Category can contain collector keywords such as "Google Alert - Export Control";
    # do not treat those as article body signals.
    cols = ["Headline", "Summary", "AI Analysis", "ClusterHeadlines"]
    return " ".join(clean(row.get(c, "")) for c in cols).lower()


def topic_text(row: pd.Series) -> str:
    # Topic can use keyword metadata, but Samsung relevance cannot.
    cols = ["Headline", "Summary", "AI Analysis", "KeywordMatches", "IssueKey", "ClusterHeadlines", "Agency", "Publisher", "Source", "Category"]
    return " ".join(clean(row.get(c, "")) for c in cols).lower()


def row_text(row: pd.Series) -> str:
    return topic_text(row)


def detect_products(text: str) -> list[str]:
    products = []
    if contains_any(text, SEMICONDUCTOR_TERMS): products.append("Semiconductor")
    if contains_any(text, MOBILE_TERMS): products.append("Mobile")
    if contains_any(text, BATTERY_TERMS): products.append("Battery")
    if contains_any(text, DISPLAY_TERMS): products.append("Display")
    if contains_any(text, SAMSUNG_EXACT_TERMS): products.append("Samsung mentioned")
    return sorted(set(products))


def detect_topic(row: pd.Series, text: str) -> str:
    # Re-classify from text. Do not trust previous IssueKey blindly.
    for topic, terms in TOPIC_RULES:
        if contains_any(text, terms):
            return topic
    issue = clean(row.get("IssueKey", "")).upper()
    if issue in TOPIC_KR:
        return issue
    return "TRADE_GENERAL"


def is_official_source(row: pd.Series) -> bool:
    blob = " ".join(clean(row.get(c, "")) for c in ["URL", "BestLinkURL", "Agency", "Source", "Publisher"]).lower()
    return any(x in blob for x in [".gov", "europa.eu", "ustr.gov", "cbp.gov", "bis.gov", "usitc.gov", "wto.org", "wcoomd.org", "customs", "관세청"])


def action_for_topic(topic: str) -> tuple[str, str]:
    if topic == "EXPORT_CONTROL":
        return "수출통제팀", "BIS/Entity List/ECCN/고객·거래 제한 여부를 확인하고 관련 법인에 스크리닝을 요청하십시오."
    if topic == "AD_CVD":
        return "통관운영/관세팀", "대상 HS·공급국·공급자·가격자료를 확인하고 AD/CVD 적용 가능성과 신고가격 영향을 점검하십시오."
    if topic == "CBAM_CARBON":
        return "ESG/구매/통관", "CBAM 대상 품목 및 EU 수출입 법인의 배출량·공급사 자료 제출 의무를 확인하십시오."
    if topic == "ORIGIN_FTA":
        return "FTA팀", "원산지 기준·CO 발급·수입 FTA 적용 영향 여부를 확인하고 증빙자료를 점검하십시오."
    if topic == "HS_CLASSIFICATION":
        return "HS/통관팀", "품목분류 기준 변경 여부와 주요 제품 HS Master 영향 여부를 확인하십시오."
    if topic == "TARIFF":
        return "통관운영/FTA팀", "관세율·시행일·대상국·대상품목을 확인하고 수입원가 및 가격 영향을 점검하십시오."
    return "통관운영", "업무 관련성이 있는지 확인 후 모니터링하십시오."


def score_row(row: pd.Series) -> dict:
    text = topic_text(row)
    ctext = content_text(row)
    headline = clean(row.get("Headline", ""))
    link, link_status = choose_best_link(row)
    dt = parse_dt(row.get("Date", ""))
    cdt = parse_dt(row.get("CollectedAt", ""))
    basis_dt = dt if not pd.isna(dt) else cdt
    now = pd.Timestamp(datetime.now())
    age_hours = None if pd.isna(basis_dt) else (now - basis_dt).total_seconds() / 3600

    topic = detect_topic(row, text)
    products = detect_products(ctext)
    official = is_official_source(row)
    strong_policy = topic in {"EXPORT_CONTROL", "AD_CVD", "CBAM_CARBON", "ORIGIN_FTA", "HS_CLASSIFICATION", "TARIFF"}
    must_keep_policy = is_must_keep_policy_news(ctext, topic)
    direct_policy_signal = has_direct_trade_policy_signal(ctext)
    ai_chip_control_signal = (
        topic == "EXPORT_CONTROL"
        and contains_any(ctext, ["ai chip", "ai chips", "ai칩", "ai 칩", "ai 반도체"])
        and contains_any(ctext, ["china", "chinese", "중국", "中"])
        and contains_any(ctext, ["control", "restriction", "restrict", "ban", "export", "sale", "통제", "제한", "차단", "수출", "판매"])
    )
    if ai_chip_control_signal:
        direct_policy_signal = True
    samsung_mention = contains_any(ctext, SAMSUNG_EXACT_TERMS)
    product_policy = bool([p for p in products if p != "Samsung mentioned"]) and strong_policy

    rejects = []
    reasons = []
    if not link:
        rejects.append("no_valid_url")
    if age_hours is not None and age_hours > MAX_AGE_HOURS:
        rejects.append(f"old_news>{MAX_AGE_HOURS}h")
    if age_hours is not None and age_hours < -12:
        rejects.append("future_date_abnormal")
    if is_real_event_noise(text):
        rejects.append("event_training_tender_noise")
    if contains_any(text, LOW_VALUE_TERMS) and not (must_keep_policy or (strong_policy and (samsung_mention or product_policy))):
        rejects.append("low_value_general_news")
    if contains_any(text, GENERAL_ECONOMY_TERMS) and not (must_keep_policy or (strong_policy and (samsung_mention or product_policy))):
        rejects.append("general_economy_without_samsung_policy")
    if is_financial_or_industry_noise(ctext, direct_policy_signal, must_keep_policy):
        rejects.append("financial_industry_noise_without_trade_policy")
    if is_samsung_general_noise(ctext, direct_policy_signal, must_keep_policy):
        rejects.append("samsung_general_business_noise")
    if is_bilateral_industry_noise(ctext, direct_policy_signal, must_keep_policy):
        rejects.append("bilateral_industry_news_without_trade_policy")
    if topic == "EXPORT_CONTROL" and contains_any(ctext, ["ai chip", "ai chips", "ai칩", "ai 칩"]) and not direct_policy_signal:
        rejects.append("ai_chip_industry_without_control_signal")
    if topic == "EXPORT_CONTROL" and product_policy and not direct_policy_signal:
        rejects.append("export_control_industry_without_control_signal")
    if topic == "TRADE_GENERAL":
        rejects.append("trade_general_not_selected")
    # If only metadata created a policy topic but article text lacks concrete policy terms, reject.
    if strong_policy and not contains_any(text, [term for _, terms in TOPIC_RULES for term in terms]):
        rejects.append("weak_policy_text")

    # Samsung impact: Direct only when Samsung is actually mentioned. Country alone is never Direct.
    if samsung_mention:
        impact = "Direct"
        samsung_score = 100
        reasons.append("samsung_exact_mention")
    elif product_policy:
        impact = "Indirect"
        samsung_score = 78
        reasons.append("product_policy_indirect")
    elif official and topic in {"EXPORT_CONTROL", "AD_CVD", "CBAM_CARBON", "ORIGIN_FTA", "HS_CLASSIFICATION", "TARIFF"}:
        impact = "Watch"
        samsung_score = 58
        reasons.append("official_policy_watch")
    elif must_keep_policy:
        impact = "Watch"
        samsung_score = 62
        reasons.append("policy_watch_must_keep")
    else:
        impact = "Reference"
        samsung_score = 20
        rejects.append("weak_samsung_relevance")

    topic_score = {"EXPORT_CONTROL":100, "AD_CVD":96, "CBAM_CARBON":90, "ORIGIN_FTA":88, "HS_CLASSIFICATION":86, "TARIFF":84, "CUSTOMS":65, "TRADE_GENERAL":25}.get(topic, 25)
    action_score = 90 if topic in {"EXPORT_CONTROL", "AD_CVD", "CBAM_CARBON", "TARIFF"} else 78 if topic in {"ORIGIN_FTA", "HS_CLASSIFICATION"} else 45
    urgency_score = 80 if contains_any(text, ["effective", "takes effect", "시행", "발효", "deadline", "due date", "immediate", "즉시"]) else 55
    recency_score = 100 if age_hours is not None and age_hours <= 24 else 85 if age_hours is not None and age_hours <= 48 else 70 if age_hours is not None and age_hours <= MAX_AGE_HOURS else 0
    final_score = round(topic_score*0.35 + samsung_score*0.30 + action_score*0.15 + urgency_score*0.10 + recency_score*0.10)

    if "event_training_tender_noise" in rejects:
        final_score = min(final_score, 40)
    if "old_news" in " ".join(rejects):
        final_score = min(final_score, 20)
    if "weak_samsung_relevance" in rejects:
        final_score = min(final_score, 55)
    if "financial_industry_noise_without_trade_policy" in rejects:
        final_score = min(final_score, 50)
    if "samsung_general_business_noise" in rejects:
        final_score = min(final_score, 48)
    if "bilateral_industry_news_without_trade_policy" in rejects or "ai_chip_industry_without_control_signal" in rejects or "export_control_industry_without_control_signal" in rejects:
        final_score = min(final_score, 50)

    if must_keep_policy:
        final_score = max(final_score, POLICY_WATCH_MIN_SCORE)
        rejects = [r for r in rejects if r not in {
            "weak_samsung_relevance",
            "general_economy_without_samsung_policy",
            "low_value_general_news",
            "financial_industry_noise_without_trade_policy",
            "samsung_general_business_noise",
        }]

    selected = (not rejects) and final_score >= MIN_SELECT_SCORE
    owner, action = action_for_topic(topic)
    products_text = "; ".join(products) if products else "본문에서 확인 불가"
    issue_kr = TOPIC_KR.get(topic, topic)
    risk = "상" if final_score >= 85 else "중" if final_score >= MIN_SELECT_SCORE else "하"
    summary = headline
    ai_analysis = f"{issue_kr} 관련 뉴스입니다. 삼성 영향도는 {impact}로 분류했습니다. 관련 제품/키워드: {products_text}."
    executive = f"{issue_kr} 이슈입니다. {action}"
    priority_group = "CORE" if selected and impact == "Direct" and final_score >= 85 else "POLICY_WATCH" if selected and must_keep_policy and impact == "Watch" else "USABLE" if selected else "EXCLUDED"

    return {
        "URL": link, "BestLinkURL": link, "original_url": link, "URL_Quality": link_status, "topic": topic, "topic_score": topic_score,
        "samsung_impact": impact, "samsung_impact_score": samsung_score, "subsidiary_score": 0,
        "action_score": action_score, "urgency_score": urgency_score, "final_score": final_score, "Risk": risk,
        "selected": "Y" if selected else "N", "priority_group": priority_group,
        "mail_section": "News Core" if priority_group == "CORE" else "Policy Watch" if priority_group == "POLICY_WATCH" else "News Usable" if priority_group == "USABLE" else "Excluded",
        "topic_keyword": issue_kr, "topic_reason": "; ".join(reasons), "issue_type": topic,
        "cluster_key": clean(row.get("IssueClusterKey", normalize_title(headline))),
        "affected_subsidiary": "SEC/HQ" if impact == "Direct" else "관련 법인 검토" if impact in {"Indirect", "Watch"} else "",
        "affected_subsidiaries": "SEC/HQ" if impact == "Direct" else "관련 법인 검토" if impact in {"Indirect", "Watch"} else "",
        "affected_products": products_text, "subsidiary_products": products_text, "subsidiary_reason": "; ".join(reasons),
        "impact_production_subsidiaries": "관련 법인 검토" if impact in {"Direct", "Indirect", "Watch"} else "",
        "impact_sales_subsidiaries": "관련 법인 검토" if impact in {"Direct", "Indirect", "Watch"} else "",
        "impact_products": products_text, "fta_impact": "검토 필요" if topic == "ORIGIN_FTA" else "본문에서 확인 불가",
        "export_control_impact": "검토 필요" if topic == "EXPORT_CONTROL" else "본문에서 확인 불가",
        "hs_impact": "검토 필요" if topic == "HS_CLASSIFICATION" else "본문에서 확인 불가",
        "tariff_impact": "검토 필요" if topic in {"TARIFF", "AD_CVD"} else "본문에서 확인 불가",
        "RequiredAction": action, "ActionOwner": owner, "ExecutiveMessage": executive,
        "samsung_score": samsung_score, "samsung_reason": "; ".join(reasons) if reasons else "weak_or_reference",
        "Summary": summary, "AI Analysis": ai_analysis, "Action Plan": action,
        "RejectReason": "; ".join(sorted(set(rejects))), "original_url": link,
        "article_extract_status": clean(row.get("article_extract_status", "NOT_FETCHED_STEP4_GUARDRAIL")),
        "article_source_type": clean(row.get("article_source_type", "STEP3_SUMMARY")),
        "effective_date_hint": "본문에서 확인 불가", "change_detail_hint": "본문에서 확인 불가", "hs_hint": "본문에서 확인 불가",
        "tariff_rate_hint": "본문에서 확인 불가", "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def read_input() -> pd.DataFrame:
    candidates = []
    if INPUT_FILE:
        candidates.append(Path(INPUT_FILE))
    for path in [INPUT_ARTICLE_FILE, INPUT_SUMMARY_FILE]:
        if path not in candidates:
            candidates.append(path)

    errors = []
    for path in candidates:
        if not path.exists():
            errors.append(f"{path}: missing")
            continue
        try:
            df = normalize_columns(pd.read_excel(path))
        except Exception as exc:
            errors.append(f"{path}: read_failed:{type(exc).__name__}")
            continue
        if df.empty:
            errors.append(f"{path}: empty")
            continue
        log(f"LOAD {path}: {len(df)} rows")
        if path.name == "3-2.news_article_summary.xlsx":
            log("INPUT MODE: article_summary body-enriched")
        else:
            log("INPUT MODE: summary fallback")
        return df

    raise FileNotFoundError("no valid news input file found / " + " | ".join(errors))


def ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(df)
    defaults = {
        "Date":"", "CollectedAt":"", "Headline":"", "URL":"", "GoogleURL":"", "OriginalURLCandidate":"", "BestLinkURL":"",
        "Country":"Global", "Agency":"", "Publisher":"", "Source":"", "SourceFile":"", "KeywordMatches":"", "ClusterHeadlines":"",
        "RegulationRelated":"N", "RegulationTransferType":"None", "IssueClusterKey":"", "SelectReason":"", "article_body":"",
    }
    for k,v in defaults.items():
        if k not in df.columns:
            df[k] = v
    return df


def append_reason(existing: str, reason: str) -> str:
    parts = [x.strip() for x in clean(existing).split(";") if x.strip()]
    if reason not in parts:
        parts.append(reason)
    return "; ".join(parts)


def report_issue_key(row: pd.Series) -> str:
    topic = clean(row.get("topic", "TRADE_GENERAL")).upper()
    blob = " ".join(clean(row.get(c, "")) for c in [
        "Headline", "Summary", "AI Analysis", "ClusterHeadlines", "topic_reason", "KeywordMatches"
    ]).lower()

    if contains_any(blob, ["ai chip", "ai chips", "ai칩", "ai 칩"]) and contains_any(blob, ["taiwan", "대만"]) and contains_any(blob, ["china", "중국"]):
        return "EXPORT_CONTROL:taiwan_ai_chip_china"
    if contains_any(blob, ["belgium", "belgian", "벨기에"]) and contains_any(blob, ["semiconductor", "반도체", "battery", "배터리"]):
        return "BILATERAL:belgium_semiconductor_battery"
    if contains_any(blob, ["uk", "britain", "영국"]) and contains_any(blob, ["steel", "철강"]) and contains_any(blob, ["tariff", "관세", "quota", "쿼터"]):
        return "TARIFF:uk_steel_tariff_quota"
    if contains_any(blob, ["eu", "european union", "유럽연합"]) and contains_any(blob, ["steel", "철강"]) and contains_any(blob, ["cbam", "carbon", "탄소", "tariff", "관세", "quota", "쿼터"]):
        return "CBAM_TARIFF:eu_steel_cbam_tariff"
    if contains_any(blob, ["cbam certificate", "cbam certificate price", "certificate price"]):
        return "CBAM_CARBON:certificate_price"
    if contains_any(blob, ["section 301", "301조"]):
        return "TARIFF:section_301"
    if contains_any(blob, ["section 232", "232조"]):
        return "TARIFF:section_232"
    if contains_any(blob, ["forced labor", "uflpa", "강제노동"]):
        return "EXPORT_CONTROL:forced_labor"
    cluster = normalize_title(clean(row.get("cluster_key", "")))
    if cluster and len(cluster) >= 8:
        return f"{topic}:{cluster[:80]}"

    title = normalize_title(clean(row.get("Headline", "")))
    title = re.sub(r"\b(reuters|bloomberg|guardian|financial times|news|뉴스|단독|속보)\b", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return f"{topic}:{title[:80]}"


def compress_report_duplicates(audit: pd.DataFrame) -> pd.DataFrame:
    if audit.empty or "selected" not in audit.columns:
        return audit

    audit = audit.copy()
    selected_mask = audit["selected"].eq("Y")
    if not selected_mask.any():
        return audit

    selected = audit[selected_mask].copy()
    selected["_report_issue_key"] = selected.apply(report_issue_key, axis=1)
    selected = selected.sort_values(
        ["final_score", "topic_score", "samsung_impact_score", "Date"],
        ascending=[False, False, False, False],
    )
    duplicate_idx = selected[selected.duplicated("_report_issue_key", keep="first")].index
    if len(duplicate_idx) == 0:
        return audit

    audit.loc[duplicate_idx, "selected"] = "N"
    audit.loc[duplicate_idx, "priority_group"] = "EXCLUDED"
    audit.loc[duplicate_idx, "mail_section"] = "Excluded"
    audit.loc[duplicate_idx, "Risk"] = "하"
    audit.loc[duplicate_idx, "final_score"] = audit.loc[duplicate_idx, "final_score"].apply(lambda v: min(int(v or 0), MIN_SELECT_SCORE - 1))
    audit.loc[duplicate_idx, "RejectReason"] = audit.loc[duplicate_idx, "RejectReason"].apply(
        lambda v: append_reason(v, "report_issue_duplicate_compressed")
    )
    return audit


def resolve_selected_google_links(audit: pd.DataFrame) -> pd.DataFrame:
    if audit.empty or "selected" not in audit.columns:
        return audit
    audit = audit.copy()
    selected_idx = audit.index[audit["selected"].eq("Y")].tolist()
    for idx in selected_idx:
        current = safe_url(audit.at[idx, "BestLinkURL"] if "BestLinkURL" in audit.columns else audit.at[idx, "URL"])
        if not is_google_article_redirect(current):
            fixed = safe_url(current)
            audit.at[idx, "URL"] = fixed
            audit.at[idx, "BestLinkURL"] = fixed
            audit.at[idx, "original_url"] = fixed
            continue

        resolved = resolve_google_news_url(current)
        if resolved:
            audit.at[idx, "URL"] = resolved
            audit.at[idx, "BestLinkURL"] = resolved
            audit.at[idx, "original_url"] = resolved
            audit.at[idx, "URL_Quality"] = "GOOGLE_NEWS_RESOLVED"
        else:
            audit.at[idx, "selected"] = "N"
            audit.at[idx, "priority_group"] = "EXCLUDED"
            audit.at[idx, "mail_section"] = "Excluded"
            audit.at[idx, "Risk"] = "하"
            audit.at[idx, "final_score"] = min(int(audit.at[idx, "final_score"] or 0), MIN_SELECT_SCORE - 1)
            audit.at[idx, "RejectReason"] = append_reason(audit.at[idx, "RejectReason"], "google_news_original_url_unresolved")
    return audit


POLICY_EXPAND_TOPICS = {
    "EXPORT_CONTROL", "AD_CVD", "CBAM_CARBON", "ORIGIN_FTA",
    "HS_CLASSIFICATION", "TARIFF", "CUSTOMS",
}

POLICY_EXPAND_HARD_REJECTS = {
    "no_valid_url",
    "event_training_tender_noise",
    "financial_industry_noise_without_trade_policy",
    "samsung_general_business_noise",
    "general_economy_without_samsung_policy",
    "low_value_general_news",
    "bilateral_industry_news_without_trade_policy",
    "ai_chip_industry_without_control_signal",
    "export_control_industry_without_control_signal",
    "google_news_original_url_unresolved",
    "future_date_abnormal",
}


def policy_expand_text(row: pd.Series) -> str:
    cols = [
        "Headline", "Summary", "ClusterHeadlines",
    ]
    return " ".join(clean(row.get(c, "")) for c in cols).lower()


POLICY_EXPAND_DIRECT_TERMS = [
    "tariff", "customs", "customs duty", "import duty", "anti-dumping", "antidumping",
    "countervailing", "ad/cvd", "section 301", "section 232", "quota", "tariff quota",
    "cbam", "carbon border", "origin", "rules of origin", "fta", "cepa", "usmca",
    "export control", "entity list", "forced labor", "uflpa", "hs code",
    "classification", "clearance", "declaration",
    "관세", "통관", "수입관세", "덤핑방지", "반덤핑", "상계관세", "할당관세",
    "쿼터", "원산지", "자유무역협정", "수출통제", "강제노동", "품목분류",
    "신고", "보세", "관세율",
]


POLICY_EXPAND_CONTEXT_NOISE = [
    "iran", "tehran", "hormuz", "oil facility", "oilfield", "missile", "bombing",
    "attack", "war", "military", "ceasefire", "crude oil", "oil price",
    "이란", "테헤란", "호르무즈", "하르그", "원유", "유전", "석유시설",
    "폭격", "공격", "전쟁", "군사", "미사일", "휴전",
]


def has_direct_policy_terms(blob: str) -> bool:
    return contains_any(blob, POLICY_EXPAND_DIRECT_TERMS)


def is_context_noise_without_policy(row: pd.Series) -> bool:
    blob = policy_expand_text(row)
    return contains_any(blob, POLICY_EXPAND_CONTEXT_NOISE) and not has_direct_policy_terms(blob)


def has_policy_expand_signal(row: pd.Series) -> bool:
    blob = policy_expand_text(row)
    topic = clean(row.get("topic")).upper()
    if topic not in POLICY_EXPAND_TOPICS:
        return False
    if not has_direct_policy_terms(blob):
        return False
    if contains_any(blob, MUST_KEEP_POLICY_TERMS) or contains_any(blob, CONCRETE_TRADE_POLICY_TERMS):
        return True
    if topic == "EXPORT_CONTROL" and contains_any(blob, ["export control", "수출통제", "entity list", "forced labor", "uflpa"]):
        return True
    if topic == "AD_CVD" and contains_any(blob, ["anti-dumping", "antidumping", "countervailing", "반덤핑", "상계관세"]):
        return True
    if topic == "ORIGIN_FTA" and contains_any(blob, ["fta", "usmca", "cepa", "origin", "원산지", "trade agreement", "통상협정"]):
        return True
    if topic == "CBAM_CARBON" and contains_any(blob, ["cbam", "carbon border", "탄소국경", "탄소세"]):
        return True
    if topic == "TARIFF" and contains_any(blob, ["tariff", "customs duty", "import duty", "quota", "관세", "쿼터"]):
        return True
    if topic == "CUSTOMS" and contains_any(blob, ["customs", "clearance", "declaration", "통관", "신고", "보세"]):
        return True
    return False


def can_expand_policy_watch(row: pd.Series) -> bool:
    if clean(row.get("selected")).upper() == "Y":
        return False
    score = int(float(row.get("final_score", 0) or 0))
    if score < NEWS_EXPAND_MIN_SCORE:
        return False
    rr = clean(row.get("RejectReason"))
    reasons = {x.strip() for x in rr.split(";") if x.strip()}
    if "report_issue_duplicate_compressed" in reasons:
        return False
    if reasons & POLICY_EXPAND_HARD_REJECTS:
        return False
    if not is_valid_link(row.get("URL", "")):
        return False
    if is_context_noise_without_policy(row):
        return False
    return has_policy_expand_signal(row)


def expand_policy_watch_selection(audit: pd.DataFrame) -> pd.DataFrame:
    if audit.empty or "selected" not in audit.columns:
        return audit
    audit = audit.copy()
    selected_count = int(audit["selected"].eq("Y").sum())
    target = max(0, min(NEWS_TARGET_MIN, NEWS_TARGET_MAX, TOP_N_MAX))
    if selected_count >= target:
        return audit

    candidates = audit[audit.apply(can_expand_policy_watch, axis=1)].copy()
    if candidates.empty:
        return audit
    candidates = candidates.sort_values(
        ["final_score", "topic_score", "Date"],
        ascending=[False, False, False],
    )
    need = max(0, min(target - selected_count, NEWS_TARGET_MAX - selected_count, TOP_N_MAX - selected_count))
    add_idx = candidates.head(need).index
    if len(add_idx) == 0:
        return audit

    audit.loc[add_idx, "selected"] = "Y"
    audit.loc[add_idx, "priority_group"] = "POLICY_WATCH"
    audit.loc[add_idx, "mail_section"] = "Policy Watch"
    audit.loc[add_idx, "samsung_impact"] = audit.loc[add_idx, "samsung_impact"].replace({"Reference": "Watch", "": "Watch"})
    audit.loc[add_idx, "Risk"] = audit.loc[add_idx, "final_score"].apply(lambda v: "상" if int(v or 0) >= 85 else "중")
    audit.loc[add_idx, "RejectReason"] = audit.loc[add_idx, "RejectReason"].apply(
        lambda v: append_reason(v, "expanded_policy_watch")
    )
    audit.loc[add_idx, "final_score"] = audit.loc[add_idx, "final_score"].apply(lambda v: max(int(v or 0), NEWS_EXPAND_MIN_SCORE))
    return audit


def build(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = ensure_cols(df)
    rows = []
    for _, row in df.iterrows():
        d = row.to_dict()
        d.update(score_row(row))
        rows.append(d)
    audit = normalize_columns(pd.DataFrame(rows))
    # remove cross-source duplicates by title after scoring
    audit["_title_norm"] = audit["Headline"].apply(normalize_title)
    audit = audit.sort_values(["selected", "final_score", "Date"], ascending=[False, False, False])
    audit = audit.drop_duplicates(subset=["_title_norm"], keep="first").drop(columns=["_title_norm"], errors="ignore")
    audit = compress_report_duplicates(audit)
    audit = expand_policy_watch_selection(audit)
    audit = resolve_selected_google_links(audit)
    audit = compress_report_duplicates(audit)

    daily = audit[audit["selected"].eq("Y")].copy()
    daily = daily.sort_values(["final_score", "topic_score", "samsung_impact_score"], ascending=[False, False, False]).reset_index(drop=True)
    if len(daily) > NEWS_TARGET_MAX:
        keep_keys = set(daily.head(NEWS_TARGET_MAX)["Headline"].astype(str))
        drop_mask = audit["selected"].eq("Y") & ~audit["Headline"].astype(str).isin(keep_keys)
        audit.loc[drop_mask, "selected"] = "N"
        audit.loc[drop_mask, "priority_group"] = "EXCLUDED"
        audit.loc[drop_mask, "mail_section"] = "Excluded"
        audit.loc[drop_mask, "RejectReason"] = audit.loc[drop_mask, "RejectReason"].apply(
            lambda v: append_reason(v, "over_news_target_max")
        )
        daily = audit[audit["selected"].eq("Y")].copy().sort_values(
            ["final_score", "topic_score", "samsung_impact_score"],
            ascending=[False, False, False],
        ).reset_index(drop=True)
    daily["rank"] = range(1, len(daily)+1)
    audit = audit.sort_values(["selected", "final_score"], ascending=[False, False]).reset_index(drop=True)
    audit["rank"] = range(1, len(audit)+1)
    excluded = audit[audit["selected"].ne("Y")].copy()

    for frame in [daily, audit, excluded]:
        for col in OUTPUT_COLS:
            if col not in frame.columns:
                frame[col] = ""
    return daily[OUTPUT_COLS], audit[OUTPUT_COLS], excluded[OUTPUT_COLS]


def merge_cumulative(daily: pd.DataFrame) -> pd.DataFrame:
    if OUT_CUMULATIVE.exists():
        try:
            old = normalize_columns(pd.read_excel(OUT_CUMULATIVE))
            log(f"cumulative existing load: {len(old)} rows")
        except Exception as exc:
            log(f"cumulative load failed -> new create: {type(exc).__name__}")
            old = pd.DataFrame(columns=OUTPUT_COLS)
    else:
        log("cumulative file missing -> new create")
        old = pd.DataFrame(columns=OUTPUT_COLS)
    for col in OUTPUT_COLS:
        if col not in old.columns: old[col] = ""
        if col not in daily.columns: daily[col] = ""
    combined = pd.concat([old[OUTPUT_COLS], daily[OUTPUT_COLS]], ignore_index=True, sort=False)
    combined = normalize_columns(combined)
    key = combined["BestLinkURL"].fillna("").astype(str).str.lower().str.strip()
    title = combined["Headline"].fillna("").astype(str).str.lower().str.strip()
    combined["_key"] = key.where(key.ne(""), title)
    combined = combined.drop_duplicates(subset=["_key"], keep="last").drop(columns=["_key"], errors="ignore")
    return combined[OUTPUT_COLS].reset_index(drop=True)


def to_legacy(daily: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for i, r in daily.reset_index(drop=True).iterrows():
        rows.append({
            "No": i+1,
            "Content Type": "News",
            "Mail Group": "News - 핵심" if clean(r.get("priority_group")) == "CORE" else "News - 주요/참고",
            "Samsung Impact": clean(r.get("samsung_impact")),
            "Affected Subsidiary": clean(r.get("affected_subsidiary")),
            "Impact Reason": clean(r.get("samsung_reason")),
            "Date": clean(r.get("Date")),
            "Headline": clean(r.get("Headline")),
            "Summary": clean(r.get("Summary")),
            "AI Analysis": clean(r.get("AI Analysis")),
            "Action Plan": clean(r.get("Action Plan")),
            "Country": clean(r.get("Country")),
            "Agency": clean(r.get("Agency")),
            "Risk": clean(r.get("Risk")),
            "Importance Score": int(r.get("final_score", 0) or 0),
            "Priority Group": clean(r.get("priority_group")),
            "Issue": clean(r.get("topic_keyword")),
            "Cluster": clean(r.get("cluster_key")),
            "URL": clean(r.get("BestLinkURL")),
            "Source": clean(r.get("Source")),
            "Source File": clean(r.get("SourceFile")),
        })
    return pd.DataFrame(rows, columns=LEGACY_COLS)


def write_excel(df: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(path, index=False)
    except PermissionError:
        alt = path.with_name(path.stem + f"_{datetime.now():%Y%m%d_%H%M%S}" + path.suffix)
        df.to_excel(alt, index=False)
        log(f"SAVE fallback due to file lock: {alt}")


def main() -> None:
    print("[STEP4-2] News analysis start - GUARDRAIL v4.1")
    df = read_input()
    daily, audit, excluded = build(df)
    cumulative = merge_cumulative(daily)
    legacy = to_legacy(daily)
    write_excel(daily, OUT_SUMMARY)
    write_excel(cumulative, OUT_CUMULATIVE)
    write_excel(audit, OUT_AUDIT)
    write_excel(excluded, OUT_EXCLUDED)
    write_excel(legacy, OUT_LEGACY)
    print(f"[DONE] Daily: {OUT_SUMMARY}")
    print(f"[DONE] Cumulative: {OUT_CUMULATIVE}")
    print(f"[DONE] Audit: {OUT_AUDIT}")
    print(f"[DONE] Excluded: {OUT_EXCLUDED}")
    print(f"[DONE] Legacy: {OUT_LEGACY}")
    print(f"[ROWS] selected={len(daily)} / audit={len(audit)} / excluded={len(excluded)} / cumulative={len(cumulative)}")


if __name__ == "__main__":
    main()
