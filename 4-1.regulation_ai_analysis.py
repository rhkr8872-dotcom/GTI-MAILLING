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

BIS_VALID_CONTEXT = [
    "bis", "bureau of industry and security", "department of commerce", "commerce department",
    "entity list", "denied persons", "export control", "수출통제", "산업안보국", "산업보안국",
]

OUTPUT_COLS = [
    "No", "Content Type", "Mail Group", "Samsung Impact", "Affected Subsidiary", "Impact Reason", "Date", "Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Risk", "Importance Score", "Priority Group", "Issue", "Cluster", "URL", "Source", "Source File", "RejectReason", "KeywordMatches", "effective_date_hint", "hs_hint", "tariff_rate_hint"
]


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
    unipass_notice_force = is_unipass_notice_candidate(row)
    indirect_tax_law_force = is_indirect_customs_tax_law(row, text)
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
    summary = f"{headline} 건은 {issue} 관련 공식 규제/공지 후보입니다."
    ai = f"{issue} 이슈입니다. 시행일/대상국/대상품목은 원문 기준 확인이 필요합니다."
    return {"selected": selected, "RejectReason": "; ".join(rejects), "Issue": issue, "topic": topic, "score": score, "Risk": risk, "URL": url, "Headline": headline, "Date": clean(date_val), "Agency": clean(row.get("Agency", row.get("agency", ""))), "Source": clean(row.get("Source", row.get("source", ""))), "Summary": summary, "AI Analysis": ai, "Action Plan": action, "Owner": owner, "KeywordMatches": "; ".join(keyword_hits[:12]), "tariff_rate_hint": extract_tariff_rate(text), "effective_date_hint": clean(row.get("effective_date_hint", "본문에서 확인 불가")) or "본문에서 확인 불가", "hs_hint": clean(row.get("hs_hint", "본문에서 확인 불가")) or "본문에서 확인 불가"}

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

def main():
    print("[STEP4-1] Regulation analysis start - GUARDRAIL v4.1")
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
