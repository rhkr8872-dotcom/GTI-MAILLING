# -*- coding: utf-8 -*-
"""
GTI STEP4-1 REGULATION AI ANALYSIS - GUARDRAIL v4.1

Fixes
- Exclude stale regulations/notices older than GTI_STEP4_REG_MAX_AGE_DAYS (default 90).
- Exclude webinar/seminar/tender/opening ceremony/event notices.
- Exclude bad URLs such as fonts.googleapis / analytics.
- Do not misread arbitrary percentages as tariff rates.
- Keep only customs/trade/FTA/export-control/CBAM/AD-CVD/HS regulation items.
"""
from __future__ import annotations

import os
import re
import json
import ssl
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from urllib.parse import quote, unquote, urlparse

import pandas as pd

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_FILE = BASE_DIR / "3-1.regulation_article_summary.xlsx"
KEYWORD_FILE = BASE_DIR / "keyword.xlsx"
OUT_SUMMARY = BASE_DIR / "4-1.regulation_ai_summary.xlsx"
OUT_CUMULATIVE = BASE_DIR / "4-1.regulation_ai_cumulative.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-1.regulation_ai_excluded.xlsx"

MAX_AGE_DAYS = int(os.getenv("GTI_STEP4_REG_MAX_AGE_DAYS", "90"))
TOP_N_MAX = int(os.getenv("GTI_STEP4_REG_TOP_N_MAX", "9999"))
MIN_SCORE = int(os.getenv("GTI_STEP4_REG_MIN_SCORE", "70"))
KEYWORD_MIN_LEN = int(os.getenv("GTI_STEP4_REG_KEYWORD_MIN_LEN", "2"))

BAD_URL_PATTERNS = ["google-analytics.com", "googletagmanager.com", "doubleclick.net", "analytics.js", "fonts.googleapis.com", "fonts.gstatic.com", "googleusercontent.com", "googleadservices.com"]
EVENT_NOISE_TERMS = [
    "webinar", "seminar", "conference", "summit", "workshop", "training", "education", "lecture", "forum", "symposium",
    "registration", "tender", "call for tender", "rfp", "expo", "opening ceremony", "ceremony", "join the upcoming",
    "live streaming",
    "웨비나", "세미나", "컨퍼런스", "서밋", "워크숍", "교육", "강의", "설명회", "포럼", "입찰", "공모", "행사", "참가신청",
]
TOPIC_RULES = [
    ("AD_CVD", ["anti-dumping", "anti dumping", "antidumping", "countervailing", "countervailing duty", "countervailing duties", "ad/cvd", "cvd", "dumping duties", "반덤핑", "덤핑방지관세", "상계관세", "무역구제"]),
    ("EXPORT_CONTROL", ["export control", "export controls", "entity list", "denied persons", "bureau of industry and security", "수출통제", "전략물자", "제재", "산업안보국", "산업보안국"]),
    ("CBAM_CARBON", ["cbam", "carbon border", "carbon border adjustment", "탄소국경"]),
    ("ORIGIN_FTA", ["fta", "cepa", "usmca", "rules of origin", "origin", "원산지", "자유무역협정", "tepa"]),
    ("HS_CLASSIFICATION", ["hs code", "classification", "tariff classification", "품목분류", "hs코드"]),
    ("TARIFF", ["section 301", "301조", "section 232", "232조", "reciprocal tariff", "tariff", "tariffs", "customs duty", "import duty", "관세", "관세율", "추가관세", "상호관세"]),
    ("CUSTOMS", ["customs", "clearance", "declaration", "통관", "세관", "관세청"]),
]
TOPIC_KR = {"EXPORT_CONTROL":"수출통제", "AD_CVD":"반덤핑/상계관세", "CBAM_CARBON":"CBAM", "ORIGIN_FTA":"FTA/원산지", "HS_CLASSIFICATION":"HS/품목분류", "TARIFF":"관세정책", "CUSTOMS":"통관/세관", "TRADE_GENERAL":"무역일반"}

STRICT_TRADE_REG_TERMS = [
    "관세", "관세율", "관세청", "통관", "세관", "보세", "수입신고", "수출신고",
    "품목분류", "hs code", "hs코드", "원산지", "fta", "자유무역협정", "cepa",
    "anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "덤핑방지관세",
    "상계관세", "무역구제", "수출통제", "전략물자", "entity list", "cbam", "carbon border",
    "customs", "tariff", "tariffs", "customs duty", "import duty", "section 301", "section 232",
]

SOFT_TRADE_REG_TERMS = [
    "import", "importation", "export", "exportation", "exporters", "trade notice", "public notice",
    "trade", "e-commerce exporters", "export obligation", "import and export", "dgft", "cbic",
    "federal register", "notice of request", "information collection", "approval", "regulation",
    "수입", "수출", "무역", "통상", "공고", "고시", "입법예고", "행정예고",
]

CONCRETE_TRADE_REG_TERMS = [
    "import", "importation", "export", "exportation", "exporters", "e-commerce exporters",
    "export obligation", "import and export", "fta", "tepa", "cepa", "safeguard",
    "anti-dumping", "antidumping", "countervailing", "ad/cvd", "tariff", "customs duty",
    "import duty", "rules of origin", "hs code", "classification",
    "수입", "수출", "원산지", "관세", "반덤핑", "상계관세", "무역구제", "세이프가드",
]

GENERIC_NOTICE_ONLY_TERMS = {"notice", "public notice", "regulation", "law", "act", "decree", "공고", "고시"}

OFFICIAL_TRADE_AGENCY_TERMS = [
    "관세청", "관세법령", "유니패스", "customs", "cbp", "ustr", "usitc", "wto", "wco",
    "taxud", "trade", "commerce", "mofcom", "dgft", "cbic", "meti", "gacc",
]

GENERAL_LAW_NOISE_TERMS = [
    "민사소송법", "형사소송법", "도로교통법", "남녀고용평등", "고용보험", "장애인고용",
    "공직선거법", "주택임대차보호법", "자동차관리법", "건설기술 진흥법", "고압가스 안전관리법",
    "전자장치 부착", "제대군인", "농어업인 삶의 질", "가맹사업거래", "국가연구개발혁신법",
]

PURE_REGULATION_TERMS = [
    "regulation", "rule", "rules", "law", "decree", "ordinance", "notice", "public notice",
    "trade notice", "federal register", "determination under", "investigation", "anti-dumping",
    "antidumping", "countervailing", "customs duty", "import duty", "export obligation",
    "법", "법률", "법령", "시행령", "시행규칙", "규칙", "고시", "공고", "훈령", "예규",
    "행정규칙", "입법예고", "행정예고", "덤핑방지관세", "상계관세", "무역구제",
    "관세율", "관세법", "보세", "통관", "수출입고시", "수입규제", "수출규제",
]

LEGAL_FORM_TITLE_TERMS = [
    "regulation", "rule", "rules", "law", "decree", "ordinance", "notice", "public notice",
    "trade notice", "federal register", "determination under", "investigation",
    "법", "법률", "법령", "시행령", "시행규칙", "규칙", "고시", "공고", "훈령", "예규",
    "행정규칙", "입법예고", "행정예고", "덤핑방지관세", "상계관세", "무역구제", "지급요령",
]

POLICY_NOTICE_NOISE_TERMS = [
    "press release", "briefing", "presidentview", "pressreleaseview", "newsid=",
    "speech", "remarks", "interview", "meeting", "delegation", "cooperation",
    "support team", "task force", "one-stop", "statistics", "provisional",
    "보도자료", "브리핑", "정상회담", "주요 성과", "성과", "면담", "대표단",
    "협력", "지원팀", "원스톱", "신설", "수출입 현황", "잠정치", "발표",
    "청장", "대통령", "경제 분야", "관세 행정 지원",
    "안내", "guidelines", "credit assistance", "support for emerging",
]

PURE_REGULATION_SOURCE_TERMS = [
    "law.go.kr", "unipass.customs.go.kr/clip", "federalregister.gov", "dgft.gov.in",
    "content.dgft.gov.in", "customs.go.jp", "mof.go.jp", "world.moleg.go.kr",
    "clhs.co.kr/law", "법령", "행정규칙", "고시", "공고", "입법예고", "행정예고",
]

UNIPASS_NOTICE_FORCE_TERMS = [
    "유니패스", "유니패스(공지사항)", "unipass", "unipass.customs.go.kr",
]

INDIRECT_CUSTOMS_TAX_LAW_TERMS = [
    "조세특례제한법", "조세특례제한법 일부개정법률안",
    "관세감면", "관세 면제", "수입부가세", "수입 부가가치세", "부가가치세 영세율",
    "개별소비세", "농어촌특별세", "세액공제", "면세",
    "tax exemption", "tax incentive", "special taxation", "customs exemption",
    "import vat", "vat exemption", "zero-rated vat",
]

POLICY_BRIEFING_NEWS_TERMS = [
    "정책브리핑", "korea.kr/briefing/pressreleaseview", "pressreleaseview.do",
    "press release", "보도자료",
]

BIS_VALID_CONTEXT = [
    "bis", "bureau of industry and security", "department of commerce", "commerce department",
    "entity list", "denied persons", "export control", "수출통제", "산업안보국", "산업보안국",
]

OUTPUT_COLS = [
    "No", "Content Type", "Mail Group", "Samsung Impact", "Affected Subsidiary", "Impact Reason", "Date", "Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Risk", "Importance Score", "Priority Group", "Issue", "Cluster", "URL", "Source", "Source File", "RejectReason", "KeywordMatches", "effective_date_hint", "hs_hint", "tariff_rate_hint"
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

def log(msg): print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")
def clean(v): return "" if pd.isna(v) else str(v).strip()
def contains_any(text, terms):
    t = str(text or "").lower()
    return any(term.lower() in t for term in terms)

def contains_term(text, term):
    t = normalize_text(text)
    k = normalize_text(term)
    if not k:
        return False
    if re.fullmatch(r"[a-z0-9/.-]{2,5}", k):
        return re.search(rf"(?<![a-z0-9]){re.escape(k)}(?![a-z0-9])", t) is not None
    return k in t

def contains_terms(text, terms):
    return any(contains_term(text, term) for term in terms)

def normalize_text(v):
    return re.sub(r"\s+", " ", clean(v)).lower().strip()

def load_keyword_terms():
    if not KEYWORD_FILE.exists():
        return []
    try:
        df = pd.read_excel(KEYWORD_FILE)
        df = normalize_columns(df)
        active_col = pick_col(df, ["active", "use", "enabled"])
        if active_col:
            active = df[active_col].fillna("Y").astype(str).str.upper().str.strip()
            df = df[active.isin(["Y", "YES", "TRUE", "1"])]

        keyword_cols = [
            col for col in df.columns
            if "keyword" in str(col).lower() or str(col).lower() in ["kr", "en", "cn", "vi", "hi", "tr", "es", "pt"]
        ]
        terms = []
        for col in keyword_cols:
            terms.extend(df[col].dropna().astype(str).str.strip().tolist())

        broad_noise = {"수출", "수입", "무역", "통상", "세관", "customs", "trade", "import", "export", "bis", "aeo", "sta", "epa"}
        cleaned = []
        for term in terms:
            t = normalize_text(term)
            if len(t) < KEYWORD_MIN_LEN:
                continue
            if t in broad_noise:
                continue
            cleaned.append(term.strip())
        return sorted(set(cleaned), key=lambda x: x.lower())
    except Exception as exc:
        log(f"WARN keyword load failed: {KEYWORD_FILE} / {exc}")
        return []

KEYWORD_TERMS = []

def keyword_match_terms(text):
    terms = KEYWORD_TERMS or []
    t = normalize_text(text)
    return [term for term in terms if contains_term(t, term)]

def has_bis_valid_context(text):
    t = normalize_text(text)
    if not re.search(r"\bbis\b", t):
        return False
    return contains_any(t, BIS_VALID_CONTEXT)

def has_strict_trade_reg_signal(text, row=None):
    t = normalize_text(text)
    if contains_terms(t, STRICT_TRADE_REG_TERMS):
        return True
    if has_bis_valid_context(t):
        return True
    if keyword_match_terms(t):
        return True
    if row is not None:
        agency = normalize_text(row.get("Agency", row.get("agency", "")))
        source = normalize_text(row.get("Source", row.get("source", "")))
        if contains_terms(f"{agency} {source}", OFFICIAL_TRADE_AGENCY_TERMS):
            return contains_terms(t, ["notice", "regulation", "law", "act", "decree", "고시", "공고", "예고", "규칙", "법령", "관세", "통관"])
    return False

def source_trade_reg_signal(row, text):
    t = normalize_text(text)
    meta_blob = normalize_text(" ".join(clean(row.get(c, "")) for c in [
        "official_regulation_type",
        "official_regulation_reason",
        "protected_regulation_reason",
        "matched_policy_terms",
        "Agency",
        "agency",
        "Source",
        "source",
    ]))
    official_type = normalize_text(row.get("official_regulation_type", ""))
    protected_score = 0
    try:
        protected_score = int(float(clean(row.get("protected_regulation_score", 0)) or 0))
    except Exception:
        protected_score = 0

    if "official_trade_regulation" in official_type and contains_terms(meta_blob + " " + t, CONCRETE_TRADE_REG_TERMS):
        return True
    if contains_terms(meta_blob, STRICT_TRADE_REG_TERMS):
        return True
    if protected_score >= 80 and contains_terms(meta_blob + " " + t, CONCRETE_TRADE_REG_TERMS):
        return True
    if contains_terms(meta_blob, OFFICIAL_TRADE_AGENCY_TERMS) and contains_terms(t, CONCRETE_TRADE_REG_TERMS):
        return True
    return False

def soft_trade_keyword_hits(row, text):
    hits = keyword_match_terms(text)
    t = normalize_text(text)
    meta = normalize_text(" ".join(clean(row.get(c, "")) for c in [
        "matched_policy_terms",
        "official_regulation_reason",
        "protected_regulation_reason",
    ]))
    for term in CONCRETE_TRADE_REG_TERMS:
        if contains_term(t, term) or contains_term(meta, term):
            hits.append(term)
    return sorted(set(hits), key=lambda x: x.lower())

def is_general_law_noise(text):
    t = normalize_text(text)
    if not contains_terms(t, GENERAL_LAW_NOISE_TERMS):
        return False
    return not has_strict_trade_reg_signal(t)

def is_unipass_notice_candidate(row):
    blob = normalize_text(" ".join(clean(row.get(c, "")) for c in [
        "Agency", "agency", "Source", "source", "site_name", "URL", "url", "original_url",
    ]))
    return contains_terms(blob, UNIPASS_NOTICE_FORCE_TERMS)

def is_indirect_customs_tax_law(row, text):
    blob = normalize_text(" ".join([
        clean(row.get("Headline", row.get("title", ""))),
        clean(row.get("Agency", row.get("agency", ""))),
        clean(row.get("Source", row.get("source", ""))),
        clean(text),
    ]))
    if not contains_terms(blob, INDIRECT_CUSTOMS_TAX_LAW_TERMS):
        return False
    return contains_terms(blob, LEGAL_FORM_TITLE_TERMS) or contains_terms(blob, ["법률안", "일부개정법률안", "개정안"])

def is_policy_briefing_press_release(row):
    blob = normalize_text(" ".join(clean(row.get(c, "")) for c in [
        "Headline", "title", "Agency", "agency", "Source", "source", "URL", "url", "original_url",
    ]))
    return contains_terms(blob, POLICY_BRIEFING_NEWS_TERMS)

def is_pure_regulation_candidate(row, text, topic):
    t = normalize_text(text)
    headline = normalize_text(row.get("Headline", row.get("title", "")))
    url = normalize_text(row.get("URL", row.get("url", row.get("original_url", ""))))
    agency = normalize_text(row.get("Agency", row.get("agency", "")))
    source = normalize_text(row.get("Source", row.get("source", "")))
    official_type = normalize_text(row.get("official_regulation_type", ""))
    meta = normalize_text(" ".join(clean(row.get(c, "")) for c in [
        "official_regulation_reason",
        "protected_regulation_reason",
        "matched_policy_terms",
        "date_status",
    ]))
    blob = " ".join([headline, url, agency, source, official_type, meta, t])

    if is_policy_briefing_press_release(row):
        return False

    if is_unipass_notice_candidate(row):
        return True

    if is_indirect_customs_tax_law(row, text):
        return True

    if contains_terms(blob, POLICY_NOTICE_NOISE_TERMS) and not contains_terms(headline, LEGAL_FORM_TITLE_TERMS):
        return False

    if topic in {"AD_CVD", "ORIGIN_FTA", "HS_CLASSIFICATION"} and contains_terms(blob, PURE_REGULATION_TERMS):
        return True

    if "official_trade_regulation" in official_type and contains_terms(blob, PURE_REGULATION_TERMS):
        return True

    if contains_terms(url + " " + source + " " + agency, PURE_REGULATION_SOURCE_TERMS) and contains_terms(blob, PURE_REGULATION_TERMS):
        return True

    if contains_terms(headline, PURE_REGULATION_TERMS) and has_strict_trade_reg_signal(text, row):
        return True

    return False

def is_old_ad_cvd_review(topic, text, age_days):
    if age_days is None or age_days <= MAX_AGE_DAYS:
        return False
    return topic == "AD_CVD" or contains_terms(text, ["anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "상계관세"])
def normalize_columns(df):
    df = df.copy(); df.columns = [str(c).strip() for c in df.columns]
    return df.loc[:, ~pd.Index(df.columns).duplicated()]
def parse_dt(v):
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.isna(dt): return pd.NaT
        if getattr(dt, "tzinfo", None) is not None: dt = dt.tz_convert(None)
        return dt
    except Exception: return pd.NaT

def is_valid_url(url):
    u = safe_url(url)
    if not u.lower().startswith(("http://", "https://")): return False
    low = u.lower()
    return not any(p in low for p in BAD_URL_PATTERNS)

def safe_url(url):
    u = clean(url).replace("\r", "").replace("\n", "").strip()
    if not u:
        return ""
    return quote(unquote(u), safe=":/?#[]@!$&'()*+,;=%")

def pick_col(df, names):
    lower = {c.lower(): c for c in df.columns}
    for n in names:
        if n.lower() in lower: return lower[n.lower()]
    return None

def row_text(row):
    cols = ["Headline", "title", "Summary", "article_body", "regulation_fallback_body", "Agency", "Source", "matched_policy_terms", "official_regulation_reason"]
    return " ".join(clean(row.get(c, "")) for c in cols).lower()

def detect_topic(text):
    for topic, terms in TOPIC_RULES:
        if contains_terms(text, terms): return topic
    return "TRADE_GENERAL"

def extract_tariff_rate(text):
    # Only accept percentages close to tariff/duty/rate/관세율/세율 context. Avoid CBAM random 98/3/2/5/0 percentages.
    rates = []
    for m in re.finditer(r"(tariff|duty|rate|관세율|세율|관세)[^\n\r]{0,40}?([0-9]{1,2}(?:\.[0-9]+)?\s*%)", text, re.I):
        try:
            num = float(m.group(2).replace('%','').strip())
            if 0 < num <= 50:
                rates.append(m.group(2).replace(' ', ''))
        except Exception:
            pass
    return "; ".join(dict.fromkeys(rates)) if rates else "본문에서 확인 불가"

def action_for(topic):
    if topic == "EXPORT_CONTROL": return "수출통제팀", "BIS/Entity List/ECCN/거래상대방 스크리닝 영향 여부를 확인하십시오."
    if topic == "AD_CVD": return "통관운영/관세팀", "대상 HS·공급국·공급자·가격자료 기준 AD/CVD 적용 가능성을 점검하십시오."
    if topic == "CBAM_CARBON": return "ESG/구매/통관", "CBAM 대상 품목, 공급사 탄소자료, EU 신고 증빙 체계를 점검하십시오."
    if topic == "ORIGIN_FTA": return "FTA팀", "원산지 기준·CO 발급·수입 FTA 적용 및 증빙자료 영향을 확인하십시오."
    if topic == "HS_CLASSIFICATION": return "HS/통관팀", "품목분류 기준 변경 및 HS Master 영향 여부를 확인하십시오."
    if topic == "TARIFF": return "통관운영/FTA팀", "관세율·시행일·대상국·대상품목을 확인하고 원가 영향을 점검하십시오."
    return "통관운영", "업무 관련성 확인 후 모니터링하십시오."

def score_row(row):
    text = row_text(row)
    topic = detect_topic(text)
    headline = clean(row.get("Headline", row.get("title", "")))
    url = safe_url(row.get("URL", row.get("url", row.get("original_url", ""))))
    if not url: url = safe_url(row.get("original_url", ""))
    date_val = row.get("Date", row.get("date", ""))
    dt = parse_dt(date_val)
    now = pd.Timestamp(datetime.now())
    age_days = None if pd.isna(dt) else (now - dt).total_seconds() / 86400
    rejects = []
    keyword_hits = soft_trade_keyword_hits(row, text)
    metadata_trade_signal = source_trade_reg_signal(row, text)
    policy_briefing_news = is_policy_briefing_press_release(row)
    unipass_notice_force = is_unipass_notice_candidate(row)
    indirect_tax_law_force = is_indirect_customs_tax_law(row, text) and not policy_briefing_news
    forced_customs_trade_regulation = unipass_notice_force or indirect_tax_law_force
    strict_trade_signal = has_strict_trade_reg_signal(text, row) or metadata_trade_signal or forced_customs_trade_regulation
    old_ad_cvd_review = is_old_ad_cvd_review(topic, text, age_days)
    pure_regulation = is_pure_regulation_candidate(row, text, topic)

    if not is_valid_url(url): rejects.append("no_valid_url")
    if age_days is not None and age_days > MAX_AGE_DAYS:
        rejects.append(f"old_regulation>{MAX_AGE_DAYS}d")
        if old_ad_cvd_review:
            rejects.append("review_preserve_ad_cvd_old_date")
    if age_days is not None and age_days < -30: rejects.append("future_date_abnormal")
    event_text = (headline + " " + clean(row.get("article_body", ""))[:500] + " " + clean(row.get("regulation_fallback_body", ""))[:500]).lower()
    if contains_any(event_text, EVENT_NOISE_TERMS) and not metadata_trade_signal and not forced_customs_trade_regulation:
        rejects.append("event_training_tender_noise")
    if is_general_law_noise(text) and not metadata_trade_signal and not forced_customs_trade_regulation:
        rejects.append("general_law_not_customs_trade")
    if policy_briefing_news:
        rejects.append("policy_briefing_press_release_to_news")
    if not strict_trade_signal:
        rejects.append("not_customs_trade_keyword")
    if not pure_regulation:
        rejects.append("policy_notice_not_pure_regulation")
    if topic == "TRADE_GENERAL" and not keyword_hits and not metadata_trade_signal:
        rejects.append("weak_trade_policy_signal")

    base_map = {"EXPORT_CONTROL":100,"AD_CVD":96,"CBAM_CARBON":90,"ORIGIN_FTA":88,"HS_CLASSIFICATION":86,"TARIFF":84,"CUSTOMS":74,"TRADE_GENERAL":72 if keyword_hits else 30}
    base = base_map.get(topic, 30)
    if age_days is None and metadata_trade_signal:
        recency = 85
    else:
        recency = 100 if age_days is not None and age_days <= 30 else 85 if age_days is not None and age_days <= 60 else 70 if age_days is not None and age_days <= MAX_AGE_DAYS else 0
    score = round(base*0.75 + recency*0.25)
    if metadata_trade_signal and topic == "TRADE_GENERAL":
        score = max(score, 70)
    if forced_customs_trade_regulation:
        score = max(score, 72)
        if unipass_notice_force:
            keyword_hits.append("UNIPASS_NOTICE_FORCE_INCLUDE")
        if indirect_tax_law_force:
            keyword_hits.append("INDIRECT_CUSTOMS_TAX_LAW")
    if keyword_hits and not rejects:
        score = max(score, 72)
    if rejects:
        if "review_preserve_ad_cvd_old_date" in rejects:
            score = min(score, 55)
        else:
            score = min(score, 45 if "event_training_tender_noise" in rejects else 50)
    selected = not rejects and score >= MIN_SCORE
    owner, action = action_for(topic)
    risk = "상" if score >= 85 else "중" if score >= 70 else "하"
    issue = TOPIC_KR.get(topic, topic)
    impact = "Watch" if selected else "Reference"
    products_text = clean(row.get("affected_products", "")) or "본문에서 확인 불가"
    analysis = build_gti_ai_analysis(
        row,
        headline=headline,
        url=url,
        issue=issue,
        impact=impact,
        products_text=products_text,
        default_action=action,
        content_type="Regulation",
    )
    return {"selected": selected, "RejectReason": "; ".join(rejects), "Issue": issue, "topic": topic, "score": score, "Risk": risk, "URL": url, "Headline": headline, "Date": clean(date_val), "Agency": clean(row.get("Agency", row.get("agency", ""))), "Source": clean(row.get("Source", row.get("source", ""))), "Summary": analysis.get("Summary", ""), "AI Analysis": analysis.get("AI Analysis", ""), "Action Plan": analysis.get("Action Plan", action), "Owner": owner, "KeywordMatches": "; ".join(keyword_hits[:12]), "tariff_rate_hint": extract_tariff_rate(text), "effective_date_hint": clean(row.get("effective_date_hint", "본문에서 확인 불가")) or "본문에서 확인 불가", "hs_hint": clean(row.get("hs_hint", "본문에서 확인 불가")) or "본문에서 확인 불가", "article_extract_status": analysis.get("article_extract_status", "")}

def read_input():
    if not INPUT_FILE.exists(): raise FileNotFoundError(f"input not found: {INPUT_FILE}")
    df = normalize_columns(pd.read_excel(INPUT_FILE))
    log(f"LOAD {INPUT_FILE}: {len(df)} rows")
    # normalize common caps for scoring
    if "Headline" not in df.columns and "title" in df.columns: df["Headline"] = df["title"]
    if "URL" not in df.columns and "url" in df.columns: df["URL"] = df["url"]
    if "Date" not in df.columns and "date" in df.columns: df["Date"] = df["date"]
    if "Agency" not in df.columns and "agency" in df.columns: df["Agency"] = df["agency"]
    if "Source" not in df.columns and "source" in df.columns: df["Source"] = df["source"]
    return df

def build(df):
    rows=[]
    for _, row in df.iterrows():
        s=score_row(row)
        rows.append(s)
    audit=pd.DataFrame(rows)
    selected_all=audit[audit["selected"]].copy().sort_values(["score","Date"], ascending=[False, False]).reset_index(drop=True)
    selected=selected_all.head(TOP_N_MAX).copy().reset_index(drop=True)
    over_top=selected_all.iloc[TOP_N_MAX:].copy()
    if not over_top.empty:
        over_top["selected"] = False
        over_top["RejectReason"] = over_top["RejectReason"].fillna("").astype(str).map(lambda x: "over_top_n" if not x else f"{x}; over_top_n")
    excluded=pd.concat([audit[~audit["selected"]].copy(), over_top], ignore_index=True, sort=False).reset_index(drop=True)
    return selected, excluded, audit

def to_output(df, content_type="Regulation"):
    rows=[]
    for i,r in df.reset_index(drop=True).iterrows():
        impact = "Watch"
        rows.append({
            "No": i+1, "Content Type": content_type, "Mail Group": "Regulation" if content_type=="Regulation" else "News - 주요/참고",
            "Samsung Impact": impact, "Affected Subsidiary": "관련 법인 검토", "Impact Reason": "official_trade_regulation_watch",
            "Date": r["Date"], "Headline": r["Headline"], "Summary": r["Summary"], "AI Analysis": r["AI Analysis"], "Action Plan": r["Action Plan"],
            "Country": "", "Agency": r["Agency"], "Risk": r["Risk"], "Importance Score": int(r["score"]), "Priority Group": "CORE" if int(r["score"])>=85 else "USABLE",
            "Issue": r["Issue"], "Cluster": r["Headline"], "URL": r["URL"], "Source": r["Source"], "Source File": "3-1.regulation_article_summary.xlsx",
            "RejectReason": r.get("RejectReason", ""), "KeywordMatches": r.get("KeywordMatches", ""), "effective_date_hint": r.get("effective_date_hint", "본문에서 확인 불가"), "hs_hint": r.get("hs_hint", "본문에서 확인 불가"), "tariff_rate_hint": r.get("tariff_rate_hint", "본문에서 확인 불가")
        })
    return pd.DataFrame(rows, columns=OUTPUT_COLS)

def normalize_cum_cols(df):
    df=normalize_columns(df)
    for c in OUTPUT_COLS:
        if c not in df.columns: df[c]=""
    return df[OUTPUT_COLS]

def merge_cumulative(daily):
    if OUT_CUMULATIVE.exists():
        try:
            old=normalize_cum_cols(pd.read_excel(OUT_CUMULATIVE)); log(f"cumulative existing load: {len(old)} rows")
        except Exception: old=pd.DataFrame(columns=OUTPUT_COLS)
    else:
        old=pd.DataFrame(columns=OUTPUT_COLS); log("cumulative file missing -> new create")
    daily=normalize_cum_cols(daily)
    combined=pd.concat([old,daily], ignore_index=True, sort=False)
    key=combined["URL"].fillna("").astype(str).str.lower().str.strip()
    title=combined["Headline"].fillna("").astype(str).str.lower().str.strip()
    combined["_key"]=key.where(key.ne(""), title)
    combined=combined.drop_duplicates(subset=["_key"], keep="last").drop(columns=["_key"], errors="ignore")
    return normalize_cum_cols(combined)

def write_excel(df,path):
    path.parent.mkdir(parents=True, exist_ok=True)
    try: df.to_excel(path,index=False)
    except PermissionError:
        alt=path.with_name(path.stem+f"_{datetime.now():%Y%m%d_%H%M%S}"+path.suffix); df.to_excel(alt,index=False); log(f"SAVE fallback: {alt}")


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
]

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
    title_norm = normalize_text(headline)
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
        if normalize_text(s) == title_norm:
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
    if "덤핑" in h or "반덤핑" in h or "상계관세" in h:
        return f"{h}{marker} 덤핑방지/상계관세 관련 조치입니다. 대상 품목, 원산지, 공급자별 세율, 적용기간을 확인해 수입 원가와 신고 기준 변경 여부를 점검해야 합니다."
    if "보세공장" in h or "보세운송" in h:
        return f"{h}{marker} 보세공장·보세운송 관련 신청서 또는 반출입 신고 처리 조건 변경 안내입니다. 보세공장 운영 법인은 신고서 양식, 자동수리 조건, 시스템 반영 시점을 확인해야 합니다."
    if "전자상거래" in h and ("등록" in h or "신고서" in h or "서식" in h):
        return f"{h}{marker} 전자상거래 수출입 신고 또는 사업자 등록 절차 관련 안내입니다. 전자상거래 물류·수출입 신고 법인은 등록 요건과 제출 서류 변경 여부를 확인해야 합니다."
    if "EODES" in h or "송수신" in h:
        return f"{h}{marker} 원산지자료교환 또는 통관 시스템 송수신 중단/변경 안내입니다. 해당 기간 원산지 증빙 수취, FTA 적용, 통관 지연 가능성을 점검해야 합니다."
    if "과세가격" in h:
        return f"{h}{marker} 과세가격 결정자료 제출 방식 관련 안내입니다. 수입신고 가격자료, 특수관계자 거래자료, 관세평가 대응자료의 제출 경로와 책임자를 확인해야 합니다."
    if "FTA" in h or "CEPA" in h or "원산지" in h:
        return f"{h}은 FTA·원산지 또는 경제협정 관련 사안입니다. 대상 협정, 품목, 원산지 기준, CO 발급·수취 요건을 확인해야 합니다."
    if "export" in h.lower() or "수출" in h:
        return f"{h}은 수출 절차 또는 수출의무 관련 공식 공지입니다. 대상 품목, 허가·신고 요건, 선적 가능 기간, 법인별 적용 여부를 확인해야 합니다."
    return f"{h}은 {issue} 관련 공식 게시물입니다. 제목 기준으로 대상 국가, 품목, 신고·허가·세율·시행일 변경 여부를 확인해야 합니다."

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
    elif any(k in blob for k in ["보세", "통관", "신고", "과세가격", "전자상거래", "EODES"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. 신고서 양식, 시스템 송수신, 과세가격자료 제출, 보세공장 반출입 또는 전자상거래 신고 절차가 바뀔 수 있어 통관 지연·서류누락·신고오류 리스크를 점검해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 해당 법인과 관세사에 신고서·시스템·제출자료 변경 여부를 확인하십시오. 1주 내: 통관 SOP와 제출자료 체크리스트를 업데이트하십시오. 1개월 내: ERP/ONE-Origin/신고 시스템 반영 필요 필드를 확정하십시오. Owner: HQ Customs + 법인 통관담당 + 관세사"
    elif any(k in blob for k in ["FTA", "CEPA", "원산지", "USMCA", "협정"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. 협정 적용 시 원산지 기준, 누적, 직접운송, CO 발급·수취 요건이 기존 FTA Master·HS Master·Item Master와 일치하는지 확인해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: 대상 협정·국가·품목의 FTA 적용 가능성과 CO 발급/수취 현황을 확인하십시오. 1주 내: BOM 원산지, Vendor 원산지확인서, HS 기준 일치 여부를 점검하십시오. 1개월 내: FTA Master·HS Master·Item Master 업데이트를 진행하십시오. Owner: HQ Customs/FTA + 법인 구매/물류"
    elif any(k in blob for k in ["CBAM", "carbon border", "탄소국경"]):
        ai = f"대상 국가/지역은 {terms['countries']}입니다. EU 수출 품목 또는 공급망 중 CBAM 대상 소재·부품 사용 여부, 공급자 배출량 자료, 신고·인증 비용 반영 여부를 점검해야 합니다. 영향등급은 {impact_s}, 관련 제품은 {product_s}입니다."
        action = "즉시조치: EU향 품목과 공급자 배출량 자료 보유 여부를 확인하십시오. 1주 내: CBAM 대상 CN/HS와 공급망 배출량 Gap List를 작성하십시오. 1개월 내: 인증서 비용 추정과 ESG/통관 공동관리 프로세스를 수립하십시오. Owner: HQ Customs + ESG + EU 법인"
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

def main():
    print("[STEP4-1] Regulation analysis start - GUARDRAIL v4.1")
    _gti_step4_gemini_log_once()
    _gti_step4_extractor_log_once()
    global KEYWORD_TERMS
    KEYWORD_TERMS = load_keyword_terms()
    log(f"keyword guardrail loaded: {len(KEYWORD_TERMS)} terms")
    df=read_input()
    selected, excluded_raw, audit_raw=build(df)
    daily=to_output(selected)
    excluded=to_output(excluded_raw)
    cumulative=merge_cumulative(daily)
    write_excel(daily, OUT_SUMMARY); write_excel(cumulative, OUT_CUMULATIVE); write_excel(excluded, OUT_EXCLUDED)
    print(f"[DONE] Daily: {OUT_SUMMARY}")
    print(f"[DONE] Cumulative: {OUT_CUMULATIVE}")
    print(f"[DONE] Excluded: {OUT_EXCLUDED}")
    print(f"[ROWS] daily={len(daily)}, cumulative={len(cumulative)}, excluded={len(excluded)}")
if __name__ == "__main__": main()
