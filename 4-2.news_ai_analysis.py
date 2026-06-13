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



# ======================================================================
# GTI STEP4 Gemini Original-URL Analysis Patch v5.0
# ======================================================================

GEMINI_API_KEY = (os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or "").strip()
GEMINI_MODEL = os.getenv("GTI_GEMINI_MODEL", "gemini-1.5-flash").strip()
USE_GEMINI = os.getenv("GTI_STEP4_USE_GEMINI", "Y").strip().upper() not in {"N", "NO", "0", "FALSE"}
ARTICLE_FETCH_TIMEOUT = int(os.getenv("GTI_ARTICLE_FETCH_TIMEOUT", "12"))
ARTICLE_MAX_CHARS = int(os.getenv("GTI_ARTICLE_MAX_CHARS", "12000"))
GEMINI_CACHE_FILE = BASE_DIR / "gti_step4_gemini_cache.xlsx"
_GEMINI_CACHE = None

def _ensure_gemini_cache():
    global _GEMINI_CACHE
    if _GEMINI_CACHE is not None:
        return _GEMINI_CACHE
    _GEMINI_CACHE = {}
    if GEMINI_CACHE_FILE.exists():
        try:
            df_cache = pd.read_excel(GEMINI_CACHE_FILE)
            for _, r in df_cache.iterrows():
                key = clean(r.get("cache_key", ""))
                if key:
                    _GEMINI_CACHE[key] = {
                        "Summary": clean(r.get("Summary", "")),
                        "AI Analysis": clean(r.get("AI Analysis", "")),
                        "Action Plan": clean(r.get("Action Plan", "")),
                        "ExecutiveMessage": clean(r.get("ExecutiveMessage", "")),
                        "article_extract_status": clean(r.get("article_extract_status", "")),
                    }
        except Exception:
            _GEMINI_CACHE = {}
    return _GEMINI_CACHE

def _save_gemini_cache():
    try:
        cache = _ensure_gemini_cache()
        if not cache:
            return
        rows = []
        for key, val in cache.items():
            row = {"cache_key": key}
            row.update(val)
            rows.append(row)
        pd.DataFrame(rows).to_excel(GEMINI_CACHE_FILE, index=False)
    except Exception:
        pass

def _analysis_cache_key(url: str, headline: str) -> str:
    u = safe_url(url)
    try:
        h = normalize_text(headline)[:120]
    except Exception:
        h = clean(headline).lower()[:120]
    return f"{u}|{h}"

def _html_unescape(text: str) -> str:
    try:
        import html as _html
        return _html.unescape(text or "")
    except Exception:
        return text or ""

def _strip_html_to_text(html_text: str) -> str:
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html_text or "")
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"(?is)<noscript[^>]*>.*?</noscript>", " ", text)
    text = re.sub(r"(?i)</(p|div|li|h1|h2|h3|tr|br)>", "\n", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    text = _html_unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n", text)
    return "\n".join(line.strip() for line in text.splitlines() if line.strip()).strip()

def _extract_meta_description(html_text: str) -> str:
    patterns = [
        r'<meta[^>]+property=["\']og:description["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)["\']',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:description["\']',
        r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']description["\']',
    ]
    for pat in patterns:
        m = re.search(pat, html_text or "", re.I | re.S)
        if m:
            return _html_unescape(re.sub(r"\s+", " ", m.group(1))).strip()
    return ""

def _looks_like_title_only(text: str, title: str) -> bool:
    t = clean(text)
    h = clean(title)
    if not t:
        return True
    if h and (t == h or t.replace(" ", "") == h.replace(" ", "")):
        return True
    if h and len(t) <= len(h) + 30 and h[:25] in t:
        return True
    bad = ["관련 뉴스입니다", "공식 규제/공지 후보입니다", "본문에서 확인 불가"]
    return any(x in t for x in bad) and len(t) < 160

def fetch_article_body_for_ai(url: str) -> tuple[str, str]:
    u = safe_url(url)
    if not u:
        return "", "NO_URL"
    if u.lower().endswith(".pdf"):
        return "", "PDF_URL_BODY_NOT_EXTRACTED"
    try:
        req = urllib.request.Request(
            u,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/129 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            },
        )
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=ARTICLE_FETCH_TIMEOUT, context=ctx) as resp:
            raw = resp.read(2_000_000)
            ctype = resp.headers.get("Content-Type", "")
        charset = ""
        m = re.search(r"charset=([\w\-]+)", ctype, re.I)
        if m:
            charset = m.group(1)
        html_text = ""
        for enc in ["utf-8", charset, "cp949", "euc-kr", "latin-1"]:
            if not enc:
                continue
            try:
                html_text = raw.decode(enc, "ignore")
                break
            except Exception:
                continue
        if not html_text:
            return "", "DECODE_FAILED"
        meta = _extract_meta_description(html_text)
        body = _strip_html_to_text(html_text)
        if meta and meta not in body[:500]:
            body = meta + "\n" + body
        body = body[:ARTICLE_MAX_CHARS]
        if len(body) < 120:
            return body, "BODY_TOO_SHORT"
        return body, "FETCHED_URL_BODY"
    except Exception as exc:
        return "", f"FETCH_FAILED:{type(exc).__name__}"

def _fallback_source_body(row: pd.Series, headline: str) -> tuple[str, str]:
    for col in [
        "article_body", "regulation_fallback_body", "full_text", "FullText",
        "content", "Content", "body", "Body", "Summary", "AI Analysis",
        "ClusterHeadlines", "description", "Description",
    ]:
        val = clean(row.get(col, ""))
        if val and not _looks_like_title_only(val, headline) and len(val) >= 80:
            return val[:ARTICLE_MAX_CHARS], f"INPUT_COLUMN:{col}"
    return "", "NO_INPUT_BODY"

def _simple_body_summary(body: str, headline: str) -> str:
    if not body:
        return "본문 확인 불가: 원문 URL에서 본문을 가져오지 못했습니다. 제목만으로 요약하지 않았습니다."
    text = re.sub(r"\s+", " ", body).strip()
    parts = re.split(r"(?<=[.!?。？！])\s+|(?<=다\.)\s+|(?<=니다\.)\s+", text)
    parts = [p.strip() for p in parts if p.strip() and not _looks_like_title_only(p, headline)]
    if not parts:
        return text[:350]
    return " ".join(parts[:3])[:700]

def call_gemini_json(prompt: str) -> dict:
    if not USE_GEMINI or not GEMINI_API_KEY:
        return {}
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "topP": 0.8,
            "maxOutputTokens": 1200,
            "responseMimeType": "application/json",
        },
    }
    try:
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(endpoint, data=data, headers={"Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(req, timeout=45) as resp:
            out = json.loads(resp.read().decode("utf-8", "ignore"))
        text = out["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(text)
    except Exception as exc:
        return {"_error": f"{type(exc).__name__}: {exc}"}

def build_gti_ai_analysis(row: pd.Series, *, headline: str, url: str, issue: str, impact: str, products_text: str, default_action: str, content_type: str) -> dict:
    cache = _ensure_gemini_cache()
    key = _analysis_cache_key(url, headline)
    if key in cache and clean(cache[key].get("Summary")):
        return cache[key]

    body, status = _fallback_source_body(row, headline)
    if not body:
        body, status = fetch_article_body_for_ai(url)

    prompt = f"""
당신은 삼성전자 본사 관세/통상 리스크 분석가입니다.
아래 원문을 읽고 GTI Radar 임원보고용으로 분석하십시오.

절대 금지:
- 제목 반복 금지
- "관련 뉴스입니다", "공식 규제/공지 후보입니다" 같은 템플릿 문장 금지
- 본문에 없는 세율/HS/국가/시행일을 지어내지 말 것
- 본문을 읽을 수 없으면 Summary에 "본문 확인 불가"라고 명시

출력은 JSON만:
{{
  "Summary": "원문 기준 게시물 요약 2~3줄",
  "AI Analysis": "삼성전자 관세업무 영향. 수입통관/수출통관/FTA·원산지/HS/관세비용/수출통제 중 해당 항목을 구체적으로 설명",
  "Action Plan": "즉시조치/1주 내/1개월 내/Owner 형식의 구체적 대응방안",
  "ExecutiveMessage": "임원용 한 문단 핵심 메시지"
}}

기본 정보:
- Content Type: {content_type}
- Issue: {issue}
- Samsung Impact: {impact}
- Affected Products: {products_text}
- URL: {url}
- Headline: {headline}
- Default Action Hint: {default_action}

원문:
{body[:ARTICLE_MAX_CHARS]}
""".strip()

    result = call_gemini_json(prompt)
    if not result or result.get("_error"):
        summary = _simple_body_summary(body, headline)
        if body:
            ai = (
                f"{issue} 이슈입니다. 삼성 영향도는 {impact}입니다. "
                f"관련 제품/키워드는 {products_text or '본문에서 확인 불가'}입니다. "
                "원문 본문 기반 세부 영향은 Summary 내용을 기준으로 대상 국가·품목·HS·세율·시행일을 추가 확인해야 합니다."
            )
        else:
            ai = (
                f"본문 확인 불가로 정밀 영향 분석이 제한됩니다. "
                f"다만 제목/메타 기준 {issue} 이슈이며, 삼성 영향도는 {impact}입니다."
            )
        action_plan = (
            f"즉시조치: 원문 URL 접속 가능 여부 및 본문 확보 상태를 확인하십시오. "
            f"1주 내: 대상 국가·품목·HS·세율·시행일을 검증하십시오. "
            f"Owner: {default_action}"
        )
        executive = summary[:250]
    else:
        summary = clean(result.get("Summary", ""))
        ai = clean(result.get("AI Analysis", ""))
        action_plan = clean(result.get("Action Plan", ""))
        executive = clean(result.get("ExecutiveMessage", ""))

        if _looks_like_title_only(summary, headline):
            summary = _simple_body_summary(body, headline)
        if not ai:
            ai = f"{issue} 관련 삼성전자 관세업무 영향 확인 필요. 영향도: {impact}."
        if not action_plan:
            action_plan = default_action
        if not executive:
            executive = summary[:250]

    final = {
        "Summary": summary[:900],
        "AI Analysis": ai[:1200],
        "Action Plan": action_plan[:1200],
        "ExecutiveMessage": executive[:700],
        "article_extract_status": status if body else status,
    }
    cache[key] = final
    _save_gemini_cache()
    return final

# ======================================================================
# End of GTI STEP4 Gemini Original-URL Analysis Patch v5.0
# ======================================================================

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


def is_real_original_url(url: str) -> bool:
    u = safe_url(url)
    if not is_valid_link(u):
        return False
    if is_google_article_redirect(u):
        return False
    if is_generic_google_main(u):
        return False
    try:
        p = urlparse(u.lower())
        if "news.google" in p.netloc or "google." in p.netloc:
            return False
    except Exception:
        return False
    return True


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
            if google_url and is_real_original_url(resolved_url):
                GOOGLE_RESOLVE_CACHE[google_url] = resolved_url
    except Exception:
        return


def save_google_resolve_cache() -> None:
    try:
        rows = [
            {"google_url": google_url, "resolved_url": resolved_url, "last_checked": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            for google_url, resolved_url in GOOGLE_RESOLVE_CACHE.items()
            if google_url and is_real_original_url(resolved_url)
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
                if resolved and is_real_original_url(resolved):
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
    analysis = build_gti_ai_analysis(
        row,
        headline=headline,
        url=link,
        issue=issue_kr,
        impact=impact,
        products_text=products_text,
        default_action=action,
        content_type="News",
    )
    summary = analysis.get("Summary", "")
    ai_analysis = analysis.get("AI Analysis", "")
    executive = analysis.get("ExecutiveMessage", "") or f"{issue_kr} 이슈입니다. {action}"
    action = analysis.get("Action Plan", action)
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
        "article_extract_status": analysis.get("article_extract_status", clean(row.get("article_extract_status", "NOT_FETCHED_STEP4_GUARDRAIL"))),
        "article_source_type": clean(row.get("article_source_type", "STEP4_GEMINI_URL_BODY")),
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
            if not is_real_original_url(fixed):
                audit.at[idx, "selected"] = "N"
                audit.at[idx, "priority_group"] = "EXCLUDED"
                audit.at[idx, "mail_section"] = "Excluded"
                audit.at[idx, "Risk"] = "??"
                audit.at[idx, "final_score"] = min(int(audit.at[idx, "final_score"] or 0), MIN_SELECT_SCORE - 1)
                audit.at[idx, "RejectReason"] = append_reason(audit.at[idx, "RejectReason"], "non_original_or_google_home_url")
                continue
            audit.at[idx, "URL"] = fixed
            audit.at[idx, "BestLinkURL"] = fixed
            audit.at[idx, "original_url"] = fixed
            continue

        resolved = resolve_google_news_url(current)
        if resolved and is_real_original_url(resolved):
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



# ======================================================================
# GTI STEP4 Gemini Quality Patch v6.0
# ----------------------------------------------------------------------
# v5 보완
# - 기존 gti_step4_gemini_cache.xlsx에 저장된 fallback/일반문구를 무시하고 재분석
# - Gemini API Key가 없거나 호출 실패해도 headline 반복이 아닌 본문 기반 fallback 분석 생성
# - AI Analysis / Action Plan 반복문구 자동 차단
# - 실행 로그에 Gemini 사용 가능 여부 표시
# ======================================================================

_GENERIC_ANALYSIS_MARKERS = [
    "원문 URL 접속 가능 여부",
    "원문 본문 기반 세부 영향은 Summary 내용을 기준으로",
    "본문 확인 불가로 정밀 영향 분석이 제한됩니다",
    "관련 뉴스입니다. 삼성 영향도는",
    "이슈입니다. 삼성 영향도는",
    "관련 공식 규제/공지 후보입니다",
]

def _is_generic_or_bad_analysis(text: str) -> bool:
    t = clean(text)
    if not t:
        return True
    return any(m in t for m in _GENERIC_ANALYSIS_MARKERS)

def _is_bad_cached_analysis(item: dict, headline: str) -> bool:
    if not item:
        return True
    summary = clean(item.get("Summary", ""))
    ai = clean(item.get("AI Analysis", ""))
    action = clean(item.get("Action Plan", ""))
    status = clean(item.get("article_extract_status", ""))
    if _looks_like_title_only(summary, headline):
        return True
    if _is_generic_or_bad_analysis(ai):
        return True
    if _is_generic_or_bad_analysis(action):
        return True
    if status and not status.startswith("GEMINI_OK"):
        # v5 fallback cache. Re-analyze when possible.
        return True
    return False

def _extract_terms_for_analysis(text: str) -> dict:
    t = clean(text)
    hs = sorted(set(re.findall(r"\b\d{4}(?:\.\d{2})?(?:\.\d{2})?\b", t)))[:6]
    rates = sorted(set(re.findall(r"\b\d{1,3}(?:\.\d+)?\s*%", t)))[:6]
    countries = []
    for c in ["미국", "중국", "일본", "EU", "유럽", "베트남", "인도", "모로코", "한국", "영국", "멕시코", "캐나다", "United States", "China", "Japan", "Vietnam", "India", "Morocco", "Korea", "EU"]:
        if c.lower() in t.lower():
            countries.append(c)
    return {"hs": "; ".join(hs) or "본문에서 확인 불가", "rates": "; ".join(rates) or "본문에서 확인 불가", "countries": "; ".join(dict.fromkeys(countries)) or "본문에서 확인 불가"}

def _fallback_gti_analysis_from_body(*, body: str, headline: str, issue: str, impact: str, products_text: str, default_action: str, content_type: str) -> dict:
    summary = _simple_body_summary(body, headline)
    terms = _extract_terms_for_analysis(" ".join([headline, body, summary]))
    issue_l = clean(issue)

    if issue_l in {"반덤핑/상계관세", "AD/CVD"}:
        ai = (
            f"반덤핑/상계관세 이슈입니다. 원문상 대상 국가/지역은 {terms['countries']}, "
            f"확인된 세율 정보는 {terms['rates']}입니다. 삼성전자 관세업무 관점에서는 해당 철강·소재·부품 HS가 "
            f"해외 생산법인 또는 협력사 수입품에 포함되는지 확인해야 하며, 적용 대상이면 추가관세 비용, 원산지 증빙, "
            f"공급자 가격자료 방어 리스크가 발생할 수 있습니다. 영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 대상 품목명·HS·공급국·공급자 리스트를 수입실적과 매칭하십시오. "
            "1주 내: 최근 12개월 수입금액 기준 잠재 AD/CVD 비용을 산출하십시오. "
            "1개월 내: 원산지/가격자료/공급자 진술서 방어 파일을 구축하고 관세사 신고 기준을 공유하십시오. "
            "Owner: HQ Customs + 구매 + 해당 법인 통관담당"
        )
    elif issue_l == "수출통제":
        ai = (
            f"수출통제 이슈입니다. 원문상 관련 국가/지역은 {terms['countries']}입니다. 삼성전자 관점에서는 반도체, AI칩, "
            f"희토류, 장비·부품 등 전략물자 또는 이중용도 품목과 연결될 수 있는지 확인해야 합니다. "
            f"수출허가, 최종사용자 확인, 우회수출 스크리닝, Item Master의 Export Control Flag 관리가 필요합니다. "
            f"영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 대상 품목의 ECCN/전략물자 해당 여부와 거래국·최종사용자를 확인하십시오. "
            "1주 내: 관련 법인과 거래처 스크리닝 결과를 재점검하십시오. "
            "1개월 내: Item Master에 수출통제 Flag 및 허가필요 여부를 반영하십시오. "
            "Owner: HQ Export Control + 사업부 + 해외법인"
        )
    elif issue_l == "CBAM":
        ai = (
            f"CBAM/탄소국경조정 이슈입니다. 원문상 관련 지역은 {terms['countries']}입니다. 삼성전자 관점에서는 EU향 수출입 "
            f"품목 중 철강·알루미늄 등 CBAM 대상 원재료/부품 사용 여부, 공급사 배출량 자료 확보, CBAM 신고 및 인증서 비용 "
            f"반영 여부가 핵심입니다. 영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: EU향 품목과 공급사 배출량 자료 보유 여부를 확인하십시오. "
            "1주 내: CBAM 대상 CN/HS와 공급사별 배출량 Gap List를 작성하십시오. "
            "1개월 내: 인증서 비용 산정 및 ESG/통관 공동관리 프로세스를 수립하십시오. "
            "Owner: HQ Customs + ESG + EU 법인"
        )
    elif issue_l in {"FTA/원산지", "ORIGIN_FTA"}:
        ai = (
            f"FTA/원산지 이슈입니다. 원문상 관련 국가/지역은 {terms['countries']}입니다. 삼성전자 관점에서는 대상 협정의 "
            f"원산지 기준, 누적, 직접운송, CO 발급/수취 요건이 기존 FTA Master·HS Master·Item Master와 일치하는지 확인해야 합니다. "
            f"특혜세율 적용 오류 또는 CO 발급 오류가 발생할 수 있습니다. 영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 대상 협정·국가·품목의 FTA 적용 여부와 CO 발급/수취 현황을 확인하십시오. "
            "1주 내: BOM 원산지, Vendor 원산지확인서, HS 기준 일치 여부를 점검하십시오. "
            "1개월 내: FTA Master·HS Master·Item Master 업데이트를 진행하십시오. "
            "Owner: HQ Customs/FTA + 법인 구매/물류"
        )
    elif issue_l in {"통관/세관", "통관", "CUSTOMS"}:
        ai = (
            f"통관/세관 절차 이슈입니다. 원문상 관련 국가/지역은 {terms['countries']}입니다. 삼성전자 관점에서는 수입신고, "
            f"보세운송, 보세공장, 납세자료 제출, 관세사 신고 프로세스 변경 여부가 중요합니다. 신고 지연, 자동수리 조건 오류, "
            f"세관 제출자료 누락 리스크가 있습니다. 영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 해당 법인 관세사에 신고절차·제출자료 변경 여부를 확인하십시오. "
            "1주 내: 통관 SOP와 보세/수입신고 체크리스트를 개정하십시오. "
            "1개월 내: ONE-Origin/ERP 반영 필요 필드를 정의하십시오. "
            "Owner: HQ Customs + 법인 통관담당 + 관세사"
        )
    elif issue_l in {"HS/품목분류", "HS_CLASSIFICATION"}:
        ai = (
            f"HS/품목분류 이슈입니다. 원문에서 확인된 HS 후보는 {terms['hs']}입니다. 삼성전자 관점에서는 동일 품목에 대한 "
            f"법인·관세사별 HS 불일치, 관세율·FTA 세율·AD/CVD 적용 오류 가능성을 점검해야 합니다. "
            f"영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 관련 품목의 HS Master와 실제 신고 HS를 비교하십시오. "
            "1주 내: 불일치 품목의 Root Cause 및 변경 승인자료를 확보하십시오. "
            "1개월 내: HS 변경 Workflow와 관세율 영향표를 반영하십시오. "
            "Owner: HQ Customs + 법인 Master Data 담당"
        )
    else:
        ai = (
            f"{issue_l} 이슈입니다. 원문상 관련 국가/지역은 {terms['countries']}입니다. 삼성전자 관세업무 관점에서는 "
            f"대상 국가·품목·HS·세율·시행일을 기준으로 수입통관, 수출통관, FTA/원산지, 관세비용 영향 여부를 확인해야 합니다. "
            f"영향등급은 {impact}, 관련 제품은 {products_text}입니다."
        )
        action = (
            "즉시조치: 원문 기준 대상 국가·품목·HS·시행일을 확인하십시오. "
            "1주 내: 관련 법인 수입/수출 실적과 매칭하십시오. "
            "1개월 내: 필요 시 Master Data와 관세사 신고 기준을 업데이트하십시오. "
            f"Owner: {default_action}"
        )

    return {
        "Summary": summary[:900],
        "AI Analysis": ai[:1200],
        "Action Plan": action[:1200],
        "ExecutiveMessage": (summary[:220] + " " + ai[:240])[:700],
        "article_extract_status": "FALLBACK_RULE_BODY",
    }

def build_gti_ai_analysis(row: pd.Series, *, headline: str, url: str, issue: str, impact: str, products_text: str, default_action: str, content_type: str) -> dict:
    """v6 override: Gemini first; ignore stale fallback cache; useful fallback if Gemini unavailable."""
    body, status = _fallback_source_body(row, headline)
    if not body:
        body, status = fetch_article_body_for_ai(url)

    cache = _ensure_gemini_cache()
    key = _analysis_cache_key(url, headline)
    cached = cache.get(key)
    if cached and not _is_bad_cached_analysis(cached, headline):
        return cached

    prompt = f"""
당신은 삼성전자 본사 관세/통상 리스크 분석가입니다.
아래 원문을 읽고 GTI Radar 임원보고용으로 분석하십시오.

절대 금지:
- 제목 반복 금지
- "관련 뉴스입니다", "공식 규제/공지 후보입니다" 같은 템플릿 문장 금지
- 본문에 없는 세율/HS/국가/시행일을 지어내지 말 것
- 본문을 읽을 수 없으면 Summary에 "본문 확인 불가"라고 명시

출력은 JSON만:
{{
  "Summary": "원문 기준 게시물 요약 2~3줄",
  "AI Analysis": "삼성전자 관세업무 영향. 수입통관/수출통관/FTA·원산지/HS/관세비용/수출통제 중 해당 항목을 구체적으로 설명",
  "Action Plan": "즉시조치/1주 내/1개월 내/Owner 형식의 구체적 대응방안",
  "ExecutiveMessage": "임원용 한 문단 핵심 메시지"
}}

기본 정보:
- Content Type: {content_type}
- Issue: {issue}
- Samsung Impact: {impact}
- Affected Products: {products_text}
- URL: {url}
- Headline: {headline}
- Default Action Hint: {default_action}

원문:
{body[:ARTICLE_MAX_CHARS]}
""".strip()

    result = call_gemini_json(prompt)
    if result and not result.get("_error"):
        summary = clean(result.get("Summary", ""))
        ai = clean(result.get("AI Analysis", ""))
        action_plan = clean(result.get("Action Plan", ""))
        executive = clean(result.get("ExecutiveMessage", ""))

        if not _looks_like_title_only(summary, headline) and not _is_generic_or_bad_analysis(ai) and not _is_generic_or_bad_analysis(action_plan):
            final = {
                "Summary": summary[:900],
                "AI Analysis": ai[:1200],
                "Action Plan": action_plan[:1200],
                "ExecutiveMessage": (executive or summary)[:700],
                "article_extract_status": f"GEMINI_OK|{status}",
            }
            cache[key] = final
            _save_gemini_cache()
            return final

    final = _fallback_gti_analysis_from_body(
        body=body,
        headline=headline,
        issue=issue,
        impact=impact,
        products_text=products_text,
        default_action=default_action,
        content_type=content_type,
    )
    final["article_extract_status"] = f"{final.get('article_extract_status')}|{status}|GEMINI={'Y' if GEMINI_API_KEY else 'NO_KEY'}"
    # Cache fallback only when Gemini is unavailable, but mark it so later runs with API key can regenerate.
    cache[key] = final
    _save_gemini_cache()
    return final

def _gti_step4_gemini_log_once():
    try:
        log(f"Gemini analysis: enabled={USE_GEMINI}, api_key={'Y' if GEMINI_API_KEY else 'N'}, model={GEMINI_MODEL}, cache={GEMINI_CACHE_FILE}")
    except Exception:
        pass

# ======================================================================
# End of GTI STEP4 Gemini Quality Patch v6.0
# ======================================================================


# ======================================================================
# GTI STEP4 Article Body Extraction Patch v7.0
# ----------------------------------------------------------------------
# v6 보완
# - 원문 본문 확보율 개선: trafilatura / readability-lxml / BeautifulSoup / meta fallback
# - PDF URL 본문 추출: pypdf 또는 PyPDF2 사용 가능 시 처리
# - UNIPASS 등 동적 페이지는 URL별 상세 본문 확보 실패 시 명확한 status 기록
#
# 권장 설치:
#   pip install trafilatura beautifulsoup4 readability-lxml lxml pypdf requests
# ======================================================================

ARTICLE_MIN_CHARS = int(os.getenv("GTI_ARTICLE_MIN_CHARS", "250"))

def _optional_import(module_name: str):
    try:
        return __import__(module_name)
    except Exception:
        return None

def _decode_bytes(raw: bytes, content_type: str = "") -> str:
    charset = ""
    m = re.search(r"charset=([\w\-]+)", content_type or "", re.I)
    if m:
        charset = m.group(1)
    for enc in [charset, "utf-8", "cp949", "euc-kr", "latin-1"]:
        if not enc:
            continue
        try:
            return raw.decode(enc, "ignore")
        except Exception:
            pass
    return raw.decode("utf-8", "ignore")

def _fetch_url_bytes(url: str) -> tuple[bytes, str, str]:
    u = safe_url(url)
    if not u:
        return b"", "", "NO_URL"
    try:
        req = urllib.request.Request(
            u,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/129 Safari/537.36",
                "Accept": "text/html,application/xhtml+xml,application/xml,application/pdf;q=0.9,*/*;q=0.8",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": "https://www.google.com/",
            },
        )
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, timeout=ARTICLE_FETCH_TIMEOUT, context=ctx) as resp:
            raw = resp.read(5_000_000)
            ctype = resp.headers.get("Content-Type", "")
        return raw, ctype, "FETCH_OK"
    except Exception as exc:
        return b"", "", f"FETCH_FAILED:{type(exc).__name__}"

def _extract_pdf_text_from_bytes(raw: bytes) -> tuple[str, str]:
    if not raw:
        return "", "PDF_EMPTY"
    try:
        import io
        try:
            from pypdf import PdfReader
            lib = "pypdf"
        except Exception:
            from PyPDF2 import PdfReader
            lib = "PyPDF2"
        reader = PdfReader(io.BytesIO(raw))
        pages = []
        for page in reader.pages[:12]:
            try:
                pages.append(page.extract_text() or "")
            except Exception:
                pass
        text = "\n".join(pages).strip()
        if text:
            return text[:ARTICLE_MAX_CHARS], f"PDF_EXTRACTED:{lib}"
        return "", f"PDF_TEXT_EMPTY:{lib}"
    except Exception as exc:
        return "", f"PDF_EXTRACT_FAILED:{type(exc).__name__}"

def _extract_html_with_trafilatura(html_text: str, url: str) -> tuple[str, str]:
    try:
        import trafilatura
        extracted = trafilatura.extract(
            html_text,
            url=url,
            include_comments=False,
            include_tables=True,
            favor_recall=True,
            output_format="txt",
        )
        if extracted and len(extracted.strip()) >= ARTICLE_MIN_CHARS:
            return extracted.strip()[:ARTICLE_MAX_CHARS], "TRAFILATURA"
    except Exception:
        pass
    return "", "TRAFILATURA_EMPTY"

def _extract_html_with_readability(html_text: str) -> tuple[str, str]:
    try:
        from readability import Document
        doc = Document(html_text)
        summary_html = doc.summary(html_partial=True)
        text = _strip_html_to_text(summary_html)
        if text and len(text) >= ARTICLE_MIN_CHARS:
            return text[:ARTICLE_MAX_CHARS], "READABILITY"
    except Exception:
        pass
    return "", "READABILITY_EMPTY"

def _extract_html_with_bs4(html_text: str) -> tuple[str, str]:
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_text, "html.parser")
        for tag in soup(["script", "style", "noscript", "svg", "footer", "header", "nav", "aside"]):
            tag.decompose()
        candidates = []
        for selector in ["article", "main", "[role=main]", ".article", ".news", ".content", "#article", "#content"]:
            try:
                for node in soup.select(selector):
                    txt = node.get_text("\n", strip=True)
                    if len(txt) >= ARTICLE_MIN_CHARS:
                        candidates.append(txt)
            except Exception:
                pass
        if not candidates:
            ps = [p.get_text(" ", strip=True) for p in soup.find_all(["p", "li"]) if len(p.get_text(" ", strip=True)) >= 30]
            if ps:
                candidates.append("\n".join(ps))
        if candidates:
            text = max(candidates, key=len)
            if text and len(text) >= ARTICLE_MIN_CHARS:
                return text[:ARTICLE_MAX_CHARS], "BS4"
    except Exception:
        pass
    return "", "BS4_EMPTY"

def _extract_structured_data_text(html_text: str) -> tuple[str, str]:
    try:
        import json as _json
        texts = []
        for m in re.finditer(r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>', html_text or "", re.I | re.S):
            raw = _html_unescape(m.group(1)).strip()
            try:
                data = _json.loads(raw)
            except Exception:
                continue
            objs = data if isinstance(data, list) else [data]
            for obj in objs:
                if not isinstance(obj, dict):
                    continue
                for k in ["articleBody", "description", "abstract"]:
                    v = obj.get(k)
                    if isinstance(v, str) and len(v) > 80:
                        texts.append(v)
        if texts:
            text = "\n".join(texts)
            return text[:ARTICLE_MAX_CHARS], "JSON_LD"
    except Exception:
        pass
    return "", "JSON_LD_EMPTY"

def _extract_html_best_text(html_text: str, url: str) -> tuple[str, str]:
    extractors = [
        lambda h: _extract_structured_data_text(h),
        lambda h: _extract_html_with_trafilatura(h, url),
        lambda h: _extract_html_with_readability(h),
        lambda h: _extract_html_with_bs4(h),
    ]
    for func in extractors:
        try:
            text, status = func(html_text)
            if text and len(text.strip()) >= ARTICLE_MIN_CHARS:
                return text.strip()[:ARTICLE_MAX_CHARS], status
        except Exception:
            pass

    meta = _extract_meta_description(html_text)
    stripped = _strip_html_to_text(html_text)
    if meta and len(meta) >= 80:
        if stripped and meta not in stripped[:600]:
            return (meta + "\n" + stripped)[:ARTICLE_MAX_CHARS], "META_PLUS_STRIPPED"
        return meta[:ARTICLE_MAX_CHARS], "META_DESCRIPTION"
    if stripped:
        return stripped[:ARTICLE_MAX_CHARS], "HTML_STRIPPED"
    return "", "HTML_EMPTY"

def fetch_article_body_for_ai(url: str) -> tuple[str, str]:
    """v7 override: robust article/PDF body extraction."""
    u = safe_url(url)
    if not u:
        return "", "NO_URL"

    raw, ctype, fetch_status = _fetch_url_bytes(u)
    if not raw:
        return "", fetch_status

    low_url = u.lower()
    low_ctype = (ctype or "").lower()

    if low_url.endswith(".pdf") or "application/pdf" in low_ctype:
        text, status = _extract_pdf_text_from_bytes(raw)
        if text:
            return text, status
        return "", status

    html_text = _decode_bytes(raw, ctype)
    if not html_text:
        return "", "DECODE_FAILED"

    text, status = _extract_html_best_text(html_text, u)

    lower_text = (text or "").lower()
    dynamic_markers = [
        "javascript", "enable cookies", "access denied", "captcha", "로그인", "권한이 없습니다",
        "통합검색", "페이지를 찾을 수", "browser does not support",
    ]
    if text and len(text) < ARTICLE_MIN_CHARS and any(m.lower() in lower_text for m in dynamic_markers):
        return text, f"BODY_TOO_SHORT_DYNAMIC:{status}"
    if text and len(text) >= 80:
        return text[:ARTICLE_MAX_CHARS], status
    return text, f"BODY_TOO_SHORT:{status}"

def _gti_step4_extractor_log_once():
    try:
        mods = []
        for m in ["trafilatura", "bs4", "readability", "pypdf", "PyPDF2"]:
            mods.append(f"{m}={'Y' if _optional_import(m) else 'N'}")
        log("Article extractor: " + ", ".join(mods))
    except Exception:
        pass

# ======================================================================
# GTI STEP4 Report Quality Patch v8.0
# ----------------------------------------------------------------------
# - Prevent menu/share/recommendation text from becoming Summary.
# - Generate concrete, report-ready Korean summaries when Gemini/cache is weak.
# - Keep analysis focused on Samsung Electronics HQ customs owner actions.
# ======================================================================

_REPORT_NOISE_WORDS = [
    "공유", "카카오톡", "페이스북", "트위터", "링크 복사", "글자 크기", "인쇄", "즐겨찾기",
    "추천기사", "많이 본 기사", "이전 기사", "다음 기사", "댓글", "팝업존", "화물진행정보",
    "주요조회서비스", "전자납부", "로그인", "회원가입", "관련종목", "AI해설", "에디터 픽",
]

_REPORT_FACT_WORDS = [
    "부과", "시행", "개정", "변경", "추가", "폐지", "정지", "등록", "신청", "제출",
    "관세", "덤핑", "상계", "수출통제", "원산지", "FTA", "CEPA", "CBAM", "통관",
    "보세", "신고", "쿼터", "할당", "수입", "수출", "세율", "협정", "조사",
    "tariff", "customs", "export control", "anti-dumping", "countervailing",
]

def _report_norm(text: str) -> str:
    return re.sub(r"\s+", " ", clean(text)).lower().strip()

def _is_menu_or_noise_text(text: str) -> bool:
    t = clean(text)
    if not t:
        return True
    noise_hits = sum(1 for w in _REPORT_NOISE_WORDS if w in t)
    fact_hits = sum(1 for w in _REPORT_FACT_WORDS if w.lower() in t.lower())
    if noise_hits >= 5 and fact_hits <= 3:
        return True
    if len(t) > 1200 and noise_hits >= 8 and fact_hits <= 5:
        return True
    return False

def _report_sentences(text: str, headline: str, limit: int = 3) -> list[str]:
    raw = re.split(r"(?<=[.!?。])\s+|[\r\n]+", clean(text))
    out = []
    title_norm = _report_norm(headline)
    for s in raw:
        s = re.sub(r"\s+", " ", clean(s))
        if len(s) < 25 or len(s) > 360:
            continue
        if _looks_like_title_only(s, headline):
            continue
        if sum(1 for w in _REPORT_NOISE_WORDS if w in s) >= 3:
            continue
        if not any(w.lower() in s.lower() for w in _REPORT_FACT_WORDS):
            continue
        if _report_norm(s) == title_norm:
            continue
        out.append(s)
        if len(out) >= limit:
            break
    return out

def _headline_based_summary(headline: str, issue: str, content_type: str) -> str:
    h = clean(headline)
    marker = "는" if h[-1:] in {"내", "식", "법", "정", "항", "안", "청", "법"} else "은"
    if not h:
        return "본문 확인 불가: 제목과 원문 링크 기준으로만 검토가 필요합니다."
    if "덤핑" in h or "반덤핑" in h or "상계관세" in h or "anti-dumping" in h.lower() or "countervailing" in h.lower():
        return f"{h}{marker} 덤핑방지/상계관세 관련 뉴스입니다. 대상 품목, 원산지, 공급자별 세율, 적용기간을 확인해 수입 원가와 신고 기준 변경 여부를 점검해야 합니다."
    if "수출통제" in h or "export control" in h.lower() or "희토류" in h:
        return f"{h}{marker} 수출통제 또는 공급망 제한 관련 뉴스입니다. 관련 품목·거래국·최종사용자·허가 필요 여부를 확인해야 합니다."
    if "CBAM" in h or "carbon border" in h.lower() or "탄소국경" in h:
        return f"{h}{marker} CBAM 또는 탄소국경조정 관련 뉴스입니다. EU향 품목, 공급망 배출량 자료, 신고·인증 비용 영향을 점검해야 합니다."
    if "FTA" in h or "CEPA" in h or "USMCA" in h or "원산지" in h:
        return f"{h}{marker} FTA·원산지 또는 경제협정 관련 뉴스입니다. 대상 협정, 품목, 원산지 기준, CO 발급·수취 요건을 확인해야 합니다."
    if "관세" in h or "tariff" in h.lower() or "customs duty" in h.lower():
        return f"{h}{marker} 관세·통상 정책 관련 뉴스입니다. 대상 국가, 품목, 세율, 시행시점을 확인하고 삼성전자 수입/수출 거래에 미치는 영향을 점검해야 합니다."
    return f"{h}{marker} {issue} 관련 뉴스입니다. 제목 기준으로 대상 국가, 품목, 관세율, 통관·수출통제·원산지 영향 여부를 확인해야 합니다."

def _quality_terms_for_report(text: str) -> dict:
    t = re.sub(r"https?://\S+", " ", clean(text))
    t = re.sub(r"%[0-9A-Fa-f]{2}", " ", t)
    rate_hits = []
    for m in re.finditer(r"(?<![A-Za-z0-9%])(\d{1,2}(?:\.\d+)?|[1-5]\d(?:\.\d+)?|60(?:\.0+)?)\s*%", t):
        window = t[max(0, m.start() - 40):m.end() + 40].lower()
        if any(k in window for k in ["관세", "세율", "덤핑", "상계", "tariff", "duty", "rate", "cbam", "quota"]):
            rate_hits.append(f"{m.group(1)}%")
    rates = sorted(set(rate_hits))[:8]
    dates = sorted(set(re.findall(r"\b20\d{2}[./-]\d{1,2}[./-]\d{1,2}\b|\b\d{1,2}월\s*\d{1,2}일\b", t)))[:6]
    raw_hs = re.findall(r"\b\d{4}(?:\.\d{2}){0,2}\b", t)
    hs = sorted(set(x for x in raw_hs if not re.fullmatch(r"20\d{2}|19\d{2}", x)))[:8]
    countries = []
    for c in ["미국", "중국", "일본", "EU", "유럽", "베트남", "인도", "모로코", "한국", "영국", "캐나다", "멕시코", "United States", "China", "Japan", "Vietnam", "India", "Morocco", "Korea", "Canada", "Mexico"]:
        if c.lower() in t.lower():
            countries.append(c)
    return {
        "rates": "; ".join(rates) or "본문에서 확인 불가",
        "dates": "; ".join(dates) or "본문에서 확인 불가",
        "hs": "; ".join(hs) or "본문에서 확인 불가",
        "countries": "; ".join(dict.fromkeys(countries)) or "본문에서 확인 불가",
    }

def _quality_summary_from_body(body: str, headline: str, issue: str, content_type: str) -> str:
    if body and not _is_menu_or_noise_text(body):
        sentences = _report_sentences(body, headline, limit=3)
        if sentences:
            return " ".join(sentences)[:900]
    return _headline_based_summary(headline, issue, content_type)[:900]

def _quality_analysis_and_action(*, headline: str, body: str, issue: str, impact: str, products_text: str, default_action: str, content_type: str) -> tuple[str, str, str]:
    blob = " ".join([headline, body])
    terms = _quality_terms_for_report(blob)
    issue_s = clean(issue)
    impact_s = clean(impact) or "Watch"
    product_s = clean(products_text) or "본문에서 확인 불가"

    if any(k in blob for k in ["덤핑", "반덤핑", "상계관세", "AD/CVD", "anti-dumping", "countervailing"]):
        ai = f"대상 국가/지역은 {terms['countries']}, 확인 세율은 {terms['rates']}, HS는 {terms['hs']}입니다. 삼성전자 관세담당자는 해당 품목이 국내외 생산법인 또는 협력사 수입품에 포함되는지 확인하고, 추가 관세비용·원산지 증빙·공급자 가격자료 방어 리스크를 점검해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 품목명·HS·원산지·공급자 리스트를 수입실적과 매칭하십시오. 1주 내: 최근 12개월 수입금액 기준 잠재 AD/CVD 비용을 산출하십시오. 1개월 내: 원산지·가격자료·공급자 진술서 방어 파일을 구축하고 관세사 신고 기준을 공유하십시오. Owner: HQ Customs + 구매 + 해당 법인 통관담당"
    elif any(k in blob for k in ["수출통제", "희토류", "전략물자", "Entity List", "export control", "forced labor", "UFLPA"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. 반도체, 배터리, AI, 희토류, 장비·부품 등 전략물자 또는 이중용도 품목과 연결될 수 있으므로 수출허가, 최종사용자, 우회수출, 제재 리스트 스크리닝을 확인해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 대상 품목의 전략물자/ECCN 해당 여부와 거래국·최종사용자를 확인하십시오. 1주 내: 관련 법인 거래처 스크리닝과 End-use 확인 결과를 재점검하십시오. 1개월 내: Item Master에 수출통제 Flag와 허가 필요 여부를 반영하십시오. Owner: HQ Export Control + 사업부 + 해외법인"
    elif any(k in blob for k in ["CBAM", "carbon border", "탄소국경"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. EU 수출 품목 또는 공급망 중 CBAM 대상 소재·부품 사용 여부, 공급자 배출량 자료, 신고·인증 비용 반영 여부를 점검해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: EU향 품목과 공급자 배출량 자료 보유 여부를 확인하십시오. 1주 내: CBAM 대상 CN/HS와 공급망 배출량 Gap List를 작성하십시오. 1개월 내: 인증서 비용 추정과 ESG/통관 공동관리 프로세스를 수립하십시오. Owner: HQ Customs + ESG + EU 법인"
    elif any(k in blob for k in ["보세", "통관", "신고", "과세가격", "전자상거래", "EODES", "customs"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. 신고서 양식, 시스템 송수신, 과세가격자료 제출, 보세공장 반출입 또는 전자상거래 신고 절차가 바뀔 수 있어 통관 지연·서류누락·신고오류 리스크를 점검해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 해당 법인과 관세사에 신고서·시스템·제출자료 변경 여부를 확인하십시오. 1주 내: 통관 SOP와 제출자료 체크리스트를 업데이트하십시오. 1개월 내: ERP/ONE-Origin/신고 시스템 반영 필요 필드를 확정하십시오. Owner: HQ Customs + 법인 통관담당 + 관세사"
    elif any(k in blob for k in ["FTA", "CEPA", "원산지", "USMCA", "협정"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. 협정 적용 시 원산지 기준, 누적, 직접운송, CO 발급·수취 요건이 기존 FTA Master·HS Master·Item Master와 일치하는지 확인해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 대상 협정·국가·품목의 FTA 적용 가능성과 CO 발급/수취 현황을 확인하십시오. 1주 내: BOM 원산지, Vendor 원산지확인서, HS 기준 일치 여부를 점검하십시오. 1개월 내: FTA Master·HS Master·Item Master 업데이트를 진행하십시오. Owner: HQ Customs/FTA + 법인 구매/물류"
    else:
        ai = f"{issue_s} 사안입니다. 대상 국가/지역은 {terms['countries']}, 세율은 {terms['rates']}, HS는 {terms['hs']}입니다. 삼성전자 관세담당자는 수입통관, 수출통관, FTA·원산지, 관세비용, 수출통제 중 어느 영역에 영향을 주는지 법인별 거래실적과 매칭해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = f"즉시조치: 원문 기준 대상 국가·품목·HS·시행일을 확인하십시오. 1주 내: 관련 법인 수입/수출 실적과 영향 품목을 매핑하십시오. 1개월 내: 필요 시 Master Data와 관세사 신고 기준을 업데이트하십시오. Owner: {default_action}"

    executive = f"{headline} 관련, 삼성전자 관세담당자는 대상 품목·국가·HS·세율·시행일을 확인하고 법인별 통관/FTA/수출통제 영향 여부를 점검해야 합니다."
    return ai[:1200], action[:1200], executive[:700]

def build_gti_ai_analysis(row: pd.Series, *, headline: str, url: str, issue: str, impact: str, products_text: str, default_action: str, content_type: str) -> dict:
    """v8 override: report-quality Korean analysis; never use menu text or generic fixed comments."""
    body, status = _fallback_source_body(row, headline)
    if not body or _is_menu_or_noise_text(body):
        fetched, fetch_status = fetch_article_body_for_ai(url)
        if fetched and not _is_menu_or_noise_text(fetched):
            body, status = fetched, fetch_status

    summary = _quality_summary_from_body(body, headline, issue, content_type)
    ai, action_plan, executive = _quality_analysis_and_action(
        headline=headline,
        body=body,
        issue=issue,
        impact=impact,
        products_text=products_text,
        default_action=default_action,
        content_type=content_type,
    )
    return {
        "Summary": summary[:900],
        "AI Analysis": ai[:1200],
        "Action Plan": action_plan[:1200],
        "ExecutiveMessage": executive[:700],
        "article_extract_status": f"REPORT_QUALITY_V8|{status}",
    }

# ======================================================================
# End of GTI STEP4 Report Quality Patch v8.0
# ======================================================================

# ======================================================================
# End of GTI STEP4 Article Body Extraction Patch v7.0
# ======================================================================

def main() -> None:
    print("[STEP4-2] News analysis start - GUARDRAIL v4.1")
    _gti_step4_gemini_log_once()
    _gti_step4_extractor_log_once()
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
