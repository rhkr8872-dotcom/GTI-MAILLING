# -*- coding: utf-8 -*-
"""
GTI STEP4-2 NEWS AI v25 CLEAN
- Input: 3-2.news_summary.xlsx
- Strict published-date 24h guard
- No legacy v18/v20/v23/v24 override chain
- Gemini: customs/trade YES/NO + Samsung customs impact analysis
- Final: maximum 30 news items
"""

from __future__ import annotations
import os, re, json, time, html as html_lib
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse
import pandas as pd
import requests

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_FILE = BASE_DIR / "3-2.news_summary.xlsx"
OUT_SUMMARY = BASE_DIR / "4-2.news_ai_summary.xlsx"
OUT_CUMULATIVE = BASE_DIR / "4-2.news_ai_cumulative.xlsx"
OUT_AUDIT = BASE_DIR / "4-2.news_ai_audit_candidates.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-2.news_ai_excluded.xlsx"
OUT_LEGACY = BASE_DIR / "4.news_ai_analysis.xlsx"

MAX_AGE_HOURS = int(os.getenv("GTI_STEP4_NEWS_MAX_AGE_HOURS", "24"))
TARGET_MAX = int(os.getenv("GTI_STEP4_NEWS_TARGET_MAX", "0"))  # 0 = quality-based, no fixed count
AI_REVIEW_MAX = int(os.getenv("GTI_STEP4_AI_REVIEW_MAX", "120"))
REPORT_TARGET = int(os.getenv("GTI_STEP4_NEWS_REPORT_TARGET", "30"))
WATCH_MIN_RELEVANCE = int(os.getenv("GTI_STEP4_WATCH_MIN_RELEVANCE", "5"))
GEMINI_MODEL = os.getenv("GTI_GEMINI_MODEL", "gemini-2.5-flash-lite").strip()
GEMINI_TIMEOUT = int(os.getenv("GTI_GEMINI_TIMEOUT", "20"))
GEMINI_API_KEY = (
    os.getenv("GEMINI_API_KEY", "").strip()
    or os.getenv("GOOGLE_API_KEY", "").strip()
)
USE_GEMINI = bool(GEMINI_API_KEY)

ISSUE_WEIGHT = {
    "EXPORT_CONTROL": 25,
    "AD_CVD": 24,
    "TARIFF": 22,
    "ORIGIN_FTA": 20,
    "HS_CLASSIFICATION": 19,
    "CUSTOMS": 18,
    "CBAM": 17,
    "TRADE_POLICY": 12,
}

SAMSUNG_TERMS = [
    "samsung", "삼성전자", "삼성", "semiconductor", "반도체", "chip",
    "smartphone", "스마트폰", "display", "oled", "battery", "배터리",
    "steel", "철강", "aluminum", "알루미늄", "polysilicon", "폴리실리콘",
]

OUTPUT_COLS = [
    "No", "Content Type", "Mail Group", "Samsung Impact", "Affected Subsidiary",
    "Impact Reason", "Date", "Publish Date", "Headline", "Summary", "AI Analysis",
    "Action Plan", "Country", "Agency", "Risk", "Importance Score",
    "Priority Group", "Issue", "Cluster", "URL", "Source", "Source File",
    "RejectReason", "AIRelevant", "AIRelevanceScore", "AIReason",
    "Top3 Eligible", "Body Verified", "Direct Evidence", "Missing Facts",
    "Policy Event", "Official Evidence",
]


def log(msg: str) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}")


def clean(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(v)).strip()


def as_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    return clean(value).lower() in {"1", "true", "yes", "y"}


EVENT_ONLY_TERMS = [
    "경진대회", "품목분류 경진", "실력 겨루", "겨룬다", "퀴즈대회", "공모전", "시상식",
    "세미나", "웨비나", "설명회", "포럼", "컨퍼런스", "워크숍",
    "contest", "competition", "seminar", "webinar", "conference", "workshop",
]
CONCRETE_MEASURE_TERMS = [
    "시행", "발효", "부과", "인상", "인하", "폐지", "개정", "공포",
    "조사 개시", "예비판정", "최종판정", "결정 고시", "명령", "고시",
    "effective", "entered into force", "imposed", "increased", "reduced",
    "amended", "investigation initiated", "preliminary determination",
    "final determination", "order", "notice",
]
OPINION_TITLE_TERMS = [
    "[세풍", "[사설", "[칼럼", "[기고", "오피니언", "column", "opinion", "editorial",
]


def event_only_noise(title: object, body: object = "") -> bool:
    title_text = clean(title).lower()
    full_text = f"{title_text} {clean(body).lower()}"
    return (
        any(term in title_text for term in EVENT_ONLY_TERMS)
        and not any(term in full_text for term in CONCRETE_MEASURE_TERMS)
    )


def opinion_article(title: object) -> bool:
    return any(term in clean(title).lower() for term in OPINION_TITLE_TERMS)


def unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out.loc[:, ~out.columns.duplicated(keep="first")].copy()


def pick_col(df: pd.DataFrame, names: list[str]):
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n.lower() in lookup:
            return lookup[n.lower()]
    return None


def normalize_url(v) -> str:
    u = clean(v)
    return u if u.startswith(("http://", "https://")) else ""


def domain(u: str) -> str:
    try:
        return urlparse(u).netloc.lower()
    except Exception:
        return ""


def load_input() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"input not found: {INPUT_FILE}")
    raw = unique_columns(pd.read_excel(INPUT_FILE))

    aliases = {
        "Date": ["Date", "Publish Date", "date", "published", "pubDate"],
        "Headline": ["Headline", "Title", "title"],
        "Summary": ["Summary", "summary", "description"],
        "URL": ["BestURL", "URL", "OriginalURL", "original_url", "url", "link"],
        "Source": ["Source", "source"],
        "Publisher": ["Publisher", "publisher", "Agency", "agency"],
        "Issue": ["IssueKey", "Issue", "issue_type", "topic"],
        "Gate": ["CandidateGate", "Gate", "priority_group"],
        "Score": ["FinalScore", "final_score", "RuleScore", "score"],
        "Cluster": ["EventKey", "Cluster", "cluster_key"],
    }
    out = pd.DataFrame(index=raw.index)
    for target, names in aliases.items():
        c = pick_col(raw, names)
        out[target] = raw[c] if c else ""

    # audit fields if available
    for c in ["InputKeyword", "InputFile", "GateReason", "URLStatus"]:
        src = pick_col(raw, [c])
        out[c] = raw[src] if src else ""

    out["Date"] = pd.to_datetime(out["Date"], errors="coerce")
    out["Headline"] = out["Headline"].fillna("").astype(str).map(clean)
    out["Summary"] = out["Summary"].fillna("").astype(str).map(clean)
    out["URL"] = out["URL"].fillna("").astype(str).map(normalize_url)
    out["Source"] = out["Source"].fillna("").astype(str).map(clean)
    out["Publisher"] = out["Publisher"].fillna("").astype(str).map(clean)
    out["Issue"] = out["Issue"].fillna("").astype(str).map(lambda x: clean(x).upper())
    out["Gate"] = out["Gate"].fillna("").astype(str).map(lambda x: clean(x).upper())
    out["Score"] = pd.to_numeric(out["Score"], errors="coerce").fillna(0)
    out["Cluster"] = out["Cluster"].fillna("").astype(str).map(clean)
    out = out[out["Headline"].ne("")].copy()
    log(f"LOAD {INPUT_FILE}: {len(out)} rows")
    return out.reset_index(drop=True)


def strict_24h(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    now = pd.Timestamp.now()
    cutoff = now - pd.Timedelta(hours=MAX_AGE_HOURS)
    dt = pd.to_datetime(df["Date"], errors="coerce")
    keep = dt.notna() & (dt >= cutoff) & (dt <= now + pd.Timedelta(hours=2))
    stale = df.loc[~keep].copy()
    stale["RejectReason"] = "STRICT_24H_PUBLISHED_DATE"
    fresh = df.loc[keep].copy().reset_index(drop=True)
    log(f"24H GUARD: {len(df)} -> {len(fresh)} / removed={len(stale)} / cutoff={cutoff}")
    return fresh, stale


def pre_score(row: pd.Series) -> int:
    base = int(float(row.get("Score", 0) or 0))
    issue = clean(row.get("Issue")).upper()
    gate = clean(row.get("Gate")).upper()
    text = f"{clean(row.get('Headline'))} {clean(row.get('Summary'))}".lower()
    s = base + ISSUE_WEIGHT.get(issue, 5)
    if gate == "CORE":
        s += 10
    if any(x in text for x in SAMSUNG_TERMS):
        s += 10
    d = domain(row.get("URL", ""))
    if any(x in d for x in [".gov", ".go.kr", "europa.eu", "wto.org", "ustr.gov", "cbp.gov", "usitc.gov"]):
        s += 6
    return min(150, s)


def gemini_json(prompt: str) -> dict:
    if not USE_GEMINI:
        return {}
    endpoint = (
        f"https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMINI_MODEL}:generateContent"
    )
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.15,
            "maxOutputTokens": 4096,
            "responseMimeType": "application/json",
        },
    }
    try:
        r = requests.post(
            endpoint,
            headers={"x-goog-api-key": GEMINI_API_KEY},
            json=payload,
            timeout=GEMINI_TIMEOUT,
        )
        if r.status_code >= 400:
            return {"_error": f"HTTP_{r.status_code}: {clean(r.text)[:500]}"}
        data = r.json()
        txt = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(txt)
    except Exception as exc:
        message = clean(exc)
        if GEMINI_API_KEY:
            message = message.replace(GEMINI_API_KEY, "***REDACTED***")
        message = re.sub(r"([?&]key=)[^&\s]+", r"\1***REDACTED***", message, flags=re.I)
        return {"_error": f"{type(exc).__name__}:{message[:180]}"}


def fetch_article_text(url: str) -> tuple[str, str]:
    """Best-effort original body fetch. Short/blocked pages never qualify for Top3."""
    if not url.startswith(("http://", "https://")):
        return "", "NO_URL"
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0 Chrome/124 Safari/537.36"}, timeout=12, allow_redirects=True)
        if r.status_code >= 400:
            return "", f"HTTP_{r.status_code}"
        text = r.text
        text = re.sub(r"(?is)<(script|style|nav|footer|header).*?>.*?</\1>", " ", text)
        text = re.sub(r"(?s)<[^>]+>", " ", text)
        text = clean(html_lib.unescape(text))
        if len(text) < 350:
            return text, "BODY_TOO_SHORT"
        return text[:12000], "BODY_OK"
    except Exception as exc:
        return "", f"FETCH_{type(exc).__name__}"


def fallback_relevant(row: pd.Series) -> tuple[bool, int, str]:
    issue = clean(row.get("Issue")).upper()
    gate = clean(row.get("Gate")).upper()
    score = pre_score(row)
    relevant = gate == "CORE" or issue in {
        "TARIFF", "CUSTOMS", "ORIGIN_FTA", "EXPORT_CONTROL", "AD_CVD",
        "HS_CLASSIFICATION", "CBAM"
    }
    return relevant, min(100, max(50 if relevant else 0, score - 35)), "RULE_FALLBACK"


def concrete_customs_signal(text: str) -> bool:
    t = clean(text).lower()
    subject = any(x in t for x in [
        "관세율", "추가관세", "반덤핑", "상계관세", "세이프가드", "원산지 규정",
        "수입신고", "수출신고", "품목분류", "전략물자", "수출통제", "제재 대상",
        "tariff rate", "additional tariff", "anti-dumping", "countervailing duty",
        "rules of origin", "customs declaration", "hs code", "export control", "sanctions",
        "section 232", "section 301", "cbam"
    ])
    action = any(x in t for x in [
        "시행", "발효", "개정", "공포", "고시", "공고", "조사 개시", "예비판정", "최종판정",
        "법원", "판결", "행정명령", "적용", "유예", "철회", "면제", "환급",
        "effective", "entered into force", "amend", "notice", "investigation", "determination",
        "court", "ruling", "executive order", "exemption", "refund"
    ])
    return subject and action


def analyze_row(row: pd.Series) -> dict:
    title = clean(row.get("Headline"))
    source_summary = clean(row.get("Summary"))
    issue = clean(row.get("Issue"))
    url = clean(row.get("URL"))
    body, body_status = fetch_article_text(url)
    evidence_text = body if len(body) >= 350 else source_summary
    body_verified = body_status == "BODY_OK" and len(body) >= 350

    prompt = f"""
당신은 삼성전자 본사 관세·통상 및 관세컴플라이언스 책임자입니다.
아래 기사 원문을 근거로 먼저 보고 대상 여부를 판정한 뒤 의사결정용 분석을 작성하십시오.

1단계 관련성 판정:
- YES: 관세율·Section 232/301·AD/CVD·세이프가드·통관·HS·과세가격·FTA/원산지·수출통제·제재·CBAM·수입규제의 구체적 조치가 핵심인 기사.
- NO: 단순 산업동향, 정치 발언, 행사/세미나, 기업실적, 주가, 일반 공급망, 관세 단어가 부수적으로만 등장하는 기사.
- 주요 글로벌 관세정책은 삼성 직접영향이 없어도 YES/Watch 가능.

2단계 분석 순서(반드시 준수):
- [사실관계] 발표 주체·국가, 조치 단계, 발표/시행일, 대상 품목·HS, 세율/쿼터, 원산지·신고·증빙 요건.
- [삼성전자 관세업무 직접영향] 영향 법인·제품·거래흐름과 수입/수출통관, HS, 과세가격, FTA/원산지, 관세비용, 조사대응 중 바뀌는 업무.
- [대응] 즉시(오늘~3영업일), 1개월 내, 상시 모니터링, Owner.

Direct/Top3 조건:
- 삼성전자가 직접 언급되고 구체적 관세조치가 연결되면 Direct 후보.
- 삼성전자 명칭이 없으면 원칙적으로 Indirect/Watch. 단, 한국 반도체(HS 854239 포함)가 중국산 제품의 환적·원산지 위험 경로로 공식 지목된 사건은 Direct 후보.
- 생산국·제품명·관세 단어가 각각 등장하는 것만으로 연결관계를 추정하지 말 것.
- 국가명, 삼성/반도체 단어, 일반 공급망 언급만으로 Direct 금지.
- 원문 본문을 확보하지 못했으면 body_verified=false, Direct 금지, Top3 Eligible=false.
- 사실이 불명확하면 추정하지 말고 missing_facts에 기록.

JSON만 출력:
{{
 "relevant": true,
 "relevance_score": 0,
 "reason": "YES/NO의 원문 근거",
 "samsung_impact": "Direct|Indirect|Watch|None",
 "top3_eligible": false,
 "body_verified": {str(body_verified).lower()},
 "direct_evidence": ["Direct 판정 원문 근거"],
 "affected_subsidiary": "영향 법인/지역 또는 관련 법인 검토",
 "risk": "상|중|하",
 "summary_ko": "[사실관계] 원문 근거 3~5문장",
 "analysis_ko": "[삼성전자 관세업무 직접영향] 영향 경로와 업무를 4~7문장",
 "action_ko": "[즉시] ... | [1개월 내] ... | [상시] ... | [Owner] ...",
 "country": "발표국/영향국",
 "agency": "발표기관/매체",
 "issue": "TARIFF|AD_CVD|EXPORT_CONTROL|SANCTIONS|CUSTOMS|HS_CLASSIFICATION|ORIGIN_FTA|CBAM_CARBON|OTHER",
 "policy_event": true,
 "official_evidence": "정부·세관·법원·공식문서와 구체적 조치 근거",
 "missing_facts": ["확인되지 않은 HS·세율·시행일 등"]
}}

Issue: {issue}
Headline: {title}
URL: {url}
Body status: {body_status}
원문/본문:
{evidence_text[:12000]}
""".strip()

    result = gemini_json(prompt)
    if not result or result.get("_error"):
        return {
            "relevant": False,
            "relevance_score": 0,
            "reason": "AI_ERROR_FAIL_CLOSED" + (f"; {result.get('_error')}" if result else ""),
            "analysis_ok": False,
            "samsung_impact": "None",
            "top3_eligible": False,
            "body_verified": body_verified,
            "direct_evidence": [],
            "affected_subsidiary": "관련 법인 검토",
            "risk": "하",
            "summary_ko": source_summary or title,
            "analysis_ko": "",
            "action_ko": "",
            "country": "",
            "agency": clean(row.get("Publisher")) or clean(row.get("Source")),
            "missing_facts": ["Gemini 분석 실패 또는 원문 근거 부족"],
            "policy_event": False,
            "official_evidence": "",
        }

    impact = clean(result.get("samsung_impact"))
    if impact not in {"Direct", "Indirect", "Watch", "None"}:
        impact = "Watch" if as_bool(result.get("relevant")) else "None"
    verified = as_bool(result.get("body_verified")) and body_verified
    evidence = [clean(x) for x in result.get("direct_evidence", []) if clean(x)]
    policy_event = (
        as_bool(result.get("policy_event"))
        and concrete_customs_signal(evidence_text)
    )
    issue_out = clean(result.get("issue")).upper()
    allowed_issues = {
        "TARIFF", "AD_CVD", "EXPORT_CONTROL", "SANCTIONS",
        "CUSTOMS", "HS_CLASSIFICATION", "ORIGIN_FTA",
        "CBAM_CARBON",
    }
    if issue_out not in allowed_issues:
        policy_event = False
        issue_out = "OTHER"

    country = clean(result.get("country"))

    if event_only_noise(title, evidence_text):
        policy_event = False
        impact = "None"
        result["relevance_score"] = 0
        result["reason"] = "EVENT_ONLY_NO_CONCRETE_CUSTOMS_CHANGE"

    # HQ customs decision rule (fail closed): do not promote an article merely
    # because product/country/customs words occur somewhere in the body.
    # Direct requires either explicit Samsung evidence, or the narrowly defined
    # Korea-semiconductor transshipment/origin route requested by the business.
    route_text = f"{title} {evidence_text} {country}".lower()
    product_terms = [
        "반도체", "semiconductor", "memory chip", "메모리칩", "스마트폰",
        "smartphone", "display", "디스플레이", "television", "tv", "가전",
        "appliance", "network equipment", "배터리", "battery",
    ]
    route_customs_terms = [
        "관세", "tariff", "환적", "transshipment", "원산지", "origin",
        "section 232", "section 301", "반덤핑", "anti-dumping", "통관",
        "customs", "수출통제", "export control", "제재", "sanction", "fta",
    ]
    samsung_named = any(term in route_text for term in [
        "삼성전자", "삼성 전자", "samsung electronics", "samsung semiconductor",
    ])
    explicit_samsung_direct = (
        verified
        and policy_event
        and samsung_named
        and any(term in route_text for term in product_terms)
        and any(term in route_text for term in route_customs_terms)
    )
    korea_semicon_transshipment_direct = (
        verified
        and policy_event
        and any(term in route_text for term in ["환적", "transshipment", "원산지 세탁", "origin laundering", "관세 회피"])
        and any(term in route_text for term in ["한국", "korea", "경기", "반도체벨트", "semiconductor belt"])
        and any(term in route_text for term in ["반도체", "semiconductor", "854239", "8542.39"])
        and any(term in route_text for term in ["중국", "china", "중국산"])
    )
    route_direct = explicit_samsung_direct or korea_semicon_transshipment_direct

    # AI prose that explicitly denies direct impact takes precedence over a
    # broad categorical label. It may still be retained as Indirect/Watch.
    analysis_lower = clean(result.get("analysis_ko")).lower()
    denies_direct = any(term in analysis_lower for term in [
        "직접적인 영향은 없", "직접 영향은 없", "직접적인 관련성은 없",
        "직접 관련성은 없", "직접적인 관세 업무 영향은 없", "미치는 영향은 없",
        "no direct impact", "not directly related",
    ])
    if denies_direct and not korea_semicon_transshipment_direct:
        if impact == "Direct":
            impact = "Watch"
        route_direct = False
    if route_direct:
        impact = "Direct"
        try:
            result["relevance_score"] = max(8, int(float(result.get("relevance_score", 0) or 0)))
        except Exception:
            result["relevance_score"] = 8
        if not evidence and korea_semicon_transshipment_direct:
            evidence = [
                "한국 반도체(HS 854239 포함)의 중국 연계 환적·원산지 위험 경로를 원문에서 확인"
            ]

    # Enforce the same evidence gate on Gemini's original Direct label. A
    # concrete global measure without a proven Samsung route is Indirect; an
    # article explicitly denying impact is Watch.
    if impact == "Direct" and not route_direct:
        impact = "Watch" if denies_direct else "Indirect"

    # 사설·칼럼은 공식 조치의 보조 해설로만 사용하며 Direct/Top3로 올리지 않는다.
    if opinion_article(title) and impact == "Direct":
        impact = "Watch"
        route_direct = False

    third_party_case = (
        not samsung_named
        and any(term in route_text for term in ["한국타이어", "hankook tire", "타이어", "tire"])
        and any(term in route_text for term in ["불복", "소송", "심판청구", "appeal", "lawsuit", "court challenge"])
    )
    if third_party_case and impact in {"Direct", "Indirect"}:
        impact = "Watch"
        route_direct = False

    relevant = (as_bool(result.get("relevant")) or route_direct) and verified and policy_event and bool(country)
    top3 = (
        (as_bool(result.get("top3_eligible")) or route_direct)
        and verified and impact == "Direct" and len(evidence) >= 1
    )
    if not relevant:
        impact = "None"
        top3 = False
    if not verified and impact == "Direct":
        impact = "Watch"
    result.update({
        "samsung_impact": impact,
        "top3_eligible": top3,
        "body_verified": verified,
        "direct_evidence": evidence,
        "missing_facts": [clean(x) for x in result.get("missing_facts", []) if clean(x)],
        "relevant": relevant,
        "analysis_ok": True,
        "policy_event": policy_event,
        "country": country,
        "issue": issue_out,
    })
    return result


def build() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = load_input()
    fresh, stale = strict_24h(df)
    fresh["PreScore"] = fresh.apply(pre_score, axis=1)
    fresh = fresh.sort_values(["PreScore", "Date"], ascending=[False, False], kind="stable").reset_index(drop=True)

    review = fresh.head(AI_REVIEW_MAX).copy()
    tail = fresh.iloc[AI_REVIEW_MAX:].copy()
    if not tail.empty:
        tail["RejectReason"] = "OUTSIDE_AI_REVIEW_POOL"

    audited = []
    for i, (_, row) in enumerate(review.iterrows(), 1):
        a = analyze_row(row)
        if i == 1 and "AI_ERROR_FAIL_CLOSED" in clean(a.get("reason")):
            raise RuntimeError(
                "Gemini first request failed: " + clean(a.get("reason")))
        r = row.to_dict()
        r["AIRelevant"] = "Y" if as_bool(a.get("relevant")) else "N"
        try:
            r["AIRelevanceScore"] = int(float(a.get("relevance_score", 0) or 0))
        except Exception:
            r["AIRelevanceScore"] = 0
        r["AIReason"] = clean(a.get("reason"))
        r["Samsung Impact"] = clean(a.get("samsung_impact")) or ("Watch" if r["AIRelevant"] == "Y" else "None")
        r["Top3 Eligible"] = "Y" if as_bool(a.get("top3_eligible")) else "N"
        r["Body Verified"] = "Y" if as_bool(a.get("body_verified")) else "N"
        r["Direct Evidence"] = " | ".join(clean(x) for x in a.get("direct_evidence", []) if clean(x))
        r["Missing Facts"] = " | ".join(clean(x) for x in a.get("missing_facts", []) if clean(x))
        r["Policy Event"] = "Y" if as_bool(a.get("policy_event")) else "N"
        r["Official Evidence"] = clean(a.get("official_evidence"))
        r["Analysis OK"] = "Y" if as_bool(a.get("analysis_ok")) else "N"
        r["Affected Subsidiary"] = clean(a.get("affected_subsidiary")) or "관련 법인 검토"
        r["Risk"] = clean(a.get("risk")) or "중"
        r["SummaryAI"] = clean(a.get("summary_ko")) or clean(row.get("Summary"))
        r["AnalysisAI"] = clean(a.get("analysis_ko"))
        r["ActionAI"] = clean(a.get("action_ko"))
        r["Country"] = clean(a.get("country"))
        r["Issue"] = clean(a.get("issue")) or "OTHER"
        r["Agency"] = clean(a.get("agency")) or clean(row.get("Publisher"))
        audited.append(r)
        if i % 10 == 0 or i == len(review):
            log(f"AI REVIEW {i}/{len(review)}")

    audit = pd.DataFrame(audited)
    if audit.empty:
        audit = fresh.iloc[0:0].copy()
        audit["AIRelevant"] = ""
        audit["AIRelevanceScore"] = 0
        audit["_EventOnly"] = False

    if not audit.empty:
        audit["SelectionScore"] = (
            pd.to_numeric(audit["PreScore"], errors="coerce").fillna(0) * 0.45
            + pd.to_numeric(audit["AIRelevanceScore"], errors="coerce").fillna(0) * 0.55
        ).round(1)
        audit["_EventOnly"] = audit.apply(
            lambda r: event_only_noise(r.get("Headline"), r.get("SummaryAI")), axis=1
        )

    ai_failures = int((audit.get("Analysis OK", pd.Series(index=audit.index, dtype=str)) != "Y").sum())
    if len(audit) and ai_failures / len(audit) >= 0.20:
        raise RuntimeError(
            f"Gemini analysis failed for {ai_failures}/{len(audit)} rows. "
            "Fail-closed: existing summary/mail inputs were not overwritten. Check API key, quota and model access."
        )

    # 메일 후보는 원문·정책사건·AI 관련성 최소기준을 모두 충족해야 한다.
    min_relevance = int(os.getenv("GTI_STEP4_MIN_RELEVANCE", "3"))
    selected = audit[
        audit["AIRelevant"].eq("Y")
        & audit["Body Verified"].eq("Y")
        & audit["Policy Event"].eq("Y")
        & ~audit["_EventOnly"].fillna(False)
        & pd.to_numeric(audit["AIRelevanceScore"], errors="coerce").fillna(0).ge(min_relevance)
    ].copy()
    selected = selected.sort_values(
        ["SelectionScore", "PreScore", "Date"], ascending=[False, False, False], kind="stable"
    )
    selected["_supplemental_watch"] = False

    # Quality-preserving minimum report size. Items that are verified and
    # clearly belong to a customs issue, but lack a finalized measure or a
    # proven Samsung route, may fill the report only as Watch. They can never
    # become Direct or Top3 through this supplement path.
    if REPORT_TARGET > 0 and len(selected) < REPORT_TARGET:
        used_idx = set(selected.index)
        watch_pool = audit[
            ~audit.index.isin(used_idx)
            & audit["Body Verified"].eq("Y")
            & audit["Analysis OK"].eq("Y")
            & audit["Issue"].ne("OTHER")
            & ~audit["_EventOnly"].fillna(False)
            & pd.to_numeric(audit["AIRelevanceScore"], errors="coerce").fillna(0).ge(WATCH_MIN_RELEVANCE)
        ].copy()
        watch_pool = watch_pool.sort_values(
            ["SelectionScore", "PreScore", "Date"], ascending=[False, False, False], kind="stable"
        )
        watch_pool["Samsung Impact"] = "Watch"
        watch_pool["Top3 Eligible"] = "N"
        watch_pool["_supplemental_watch"] = True
        selected = pd.concat(
            [selected, watch_pool.head(max(0, REPORT_TARGET - len(selected)))],
            axis=0, sort=False,
        )

    def semantic_event_key(row: pd.Series) -> str:
        title_text = clean(row.get("Headline")).lower()
        text = " ".join([
            clean(row.get("Headline")), clean(row.get("SummaryAI")),
            clean(row.get("Country")), clean(row.get("Issue")),
        ]).lower()
        if (
            any(x in text for x in ["미국", "트럼프", "백악관", "washington", "u.s."])
            and any(x in title_text for x in ["삼성전자", "삼전", "samsung", "sk하이닉스", "하이닉스", "반도체", "semiconductor", "chip"])
            and any(x in text for x in ["관세", "tariff", "투자 압박", "투자 이행", "미국 투자", "유치 압박", "양자택일", "편 서지 마", "pressure"])
        ):
            return "US_SEMICON_TARIFF_INVESTMENT_PRESSURE"
        if (
            any(x in text for x in ["미국", "트럼프", "u.s."])
            and any(x in title_text for x in ["드론", "drone", "무인기", "uas"])
            and any(x in text for x in ["100%", "100％", "최대 100", "section 232", "232조"])
            and any(x in text for x in ["관세", "tariff"])
        ):
            return "US_DRONE_232_TARIFF"
        rules = [
            ("US_SECTION232_DRONE_COMPONENTS_TARIFF", [["드론", "drone", "무인기", "uas"], ["section 232", "무역확장법 232", "232조"], ["관세", "tariff"]]),
            ("US_CHINA_TRANSSHIPMENT_KOREA_SEMICON", [["환적", "transshipment", "원산지 세탁", "관세 회피"], ["한국", "korea", "경기", "반도체벨트"], ["중국", "china", "중국산"]]),
            ("US_CHINA_AUTO_TARIFF_BLOCK", [["중국차", "중국산 자동차", "chinese car", "chinese vehicle"], ["127.5%", "딜러", "dealer", "봉쇄", "진입 반대"]]),
            ("KR_STEEL_TRADE_REMEDY", [["철강", "steel"], ["반덤핑", "anti-dumping", "232", "쿼터"]]),
            ("BUSAN_EXPORT_ECONOMY", [["부산"], ["제조업", "실물경제", "수출 중심"]]),
            ("EU_CHINA_AUTO_SUPPLY_CHAIN", [["유럽", "europe", "eu"], ["중국차", "chinese car", "부품망", "공급망"]]),
        ]
        for name, groups in rules:
            if all(any(term in text for term in group) for group in groups):
                return name
        words = re.findall(r"[a-z0-9가-힣]+", clean(row.get("Headline")).lower())
        stop = {"속보", "종합", "단독", "오늘", "뉴스", "the", "a", "an", "and", "of", "to", "in"}
        return clean(row.get("Issue")) + "|" + " ".join(w for w in words if w not in stop)[:90]

    selected["SemanticEventKey"] = selected.apply(semantic_event_key, axis=1)
    selected = selected.drop_duplicates("SemanticEventKey", keep="first")
    # Dedup may reduce the count; refill once from the remaining verified Watch
    # pool using a different semantic event.
    if REPORT_TARGET > 0 and len(selected) < REPORT_TARGET:
        existing_keys = set(selected["SemanticEventKey"].astype(str))
        refill = audit[
            ~audit.index.isin(set(selected.index))
            & audit["Body Verified"].eq("Y")
            & audit["Analysis OK"].eq("Y")
            & audit["Issue"].ne("OTHER")
            & ~audit["_EventOnly"].fillna(False)
            & pd.to_numeric(audit["AIRelevanceScore"], errors="coerce").fillna(0).ge(WATCH_MIN_RELEVANCE)
        ].copy()
        if not refill.empty:
            refill["Samsung Impact"] = "Watch"
            refill["Top3 Eligible"] = "N"
            refill["_supplemental_watch"] = True
            refill["SemanticEventKey"] = refill.apply(semantic_event_key, axis=1)
            refill = refill[~refill["SemanticEventKey"].isin(existing_keys)]
            refill = refill.drop_duplicates("SemanticEventKey", keep="first").sort_values(
                ["SelectionScore", "PreScore", "Date"], ascending=[False, False, False], kind="stable"
            )
            selected = pd.concat([selected, refill.head(REPORT_TARGET-len(selected))], axis=0, sort=False)
    if TARGET_MAX > 0:
        selected = selected.head(TARGET_MAX)
    selected_audit_indices = set(selected.index)
    selected = selected.reset_index(drop=True)

    selected_rows = []
    for i, r in selected.iterrows():
        supplemental_watch = bool(r.get("_supplemental_watch", False))
        impact = "Watch" if supplemental_watch else clean(r.get("Samsung Impact"))
        risk_map = {"High": "상", "Medium": "중", "Low": "하", "HIGH": "상", "MEDIUM": "중", "LOW": "하"}
        normalized_risk = risk_map.get(clean(r.get("Risk")), clean(r.get("Risk")) or "중")
        mail_group = "News - 핵심" if impact == "Direct" else "News - 주요/참고"
        selected_rows.append({
            "No": i + 1,
            "Content Type": "News",
            "Mail Group": mail_group,
            "Samsung Impact": impact or "Watch",
            "Affected Subsidiary": clean(r.get("Affected Subsidiary")) or "관련 법인 검토",
            "Impact Reason": clean(r.get("AIReason")),
            "Date": pd.to_datetime(r.get("Date"), errors="coerce"),
            "Publish Date": pd.to_datetime(r.get("Date"), errors="coerce"),
            "Headline": clean(r.get("Headline")),
            "Summary": clean(r.get("SummaryAI")),
            "AI Analysis": clean(r.get("AnalysisAI")),
            "Action Plan": clean(r.get("ActionAI")),
            "Country": clean(r.get("Country")),
            "Agency": clean(r.get("Agency")),
            "Risk": normalized_risk,
            "Importance Score": int(round(float(r.get("SelectionScore", 0) or 0))),
            "Priority Group": "CORE" if mail_group == "News - 핵심" else "USABLE",
            "Issue": clean(r.get("Issue")),
            "Cluster": clean(r.get("SemanticEventKey")) or clean(r.get("Cluster")),
            "URL": clean(r.get("URL")),
            "Source": clean(r.get("Source")),
            "Source File": str(INPUT_FILE),
            "RejectReason": "",
            "AIRelevant": clean(r.get("AIRelevant")) or "N",
            "AIRelevanceScore": int(float(r.get("AIRelevanceScore", 0) or 0)),
            "AIReason": clean(r.get("AIReason")),
            "Top3 Eligible": "N" if supplemental_watch else (clean(r.get("Top3 Eligible")) or "N"),
            "Body Verified": clean(r.get("Body Verified")) or "N",
            "Direct Evidence": clean(r.get("Direct Evidence")),
            "Missing Facts": clean(r.get("Missing Facts")),
            "Policy Event": clean(r.get("Policy Event")) or "N",
            "Official Evidence": clean(r.get("Official Evidence")),
        })
    daily = pd.DataFrame(selected_rows, columns=OUTPUT_COLS)

    rejected_ai = audit[~audit.index.isin(selected_audit_indices) & ~audit["AIRelevant"].eq("Y")].copy()
    if not rejected_ai.empty:
        rejected_ai["RejectReason"] = "AI_NOT_CUSTOMS_TRADE"
    excluded = pd.concat([stale, tail, rejected_ai], ignore_index=True, sort=False)

    return daily, audit, excluded


def merge_cumulative(daily: pd.DataFrame) -> pd.DataFrame:
    if OUT_CUMULATIVE.exists():
        try:
            old = unique_columns(pd.read_excel(OUT_CUMULATIVE))
        except Exception:
            old = pd.DataFrame()
    else:
        old = pd.DataFrame()
    combined = pd.concat([old, daily], ignore_index=True, sort=False)
    if combined.empty:
        return combined
    combined["_key"] = combined["Headline"].fillna("").astype(str).str.lower().str.strip()
    combined = combined.drop_duplicates("_key", keep="last").drop(columns="_key")
    return combined.reset_index(drop=True)


def safe_write(path: Path, df: pd.DataFrame) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(path, index=False)
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        df.to_excel(alt, index=False)
        log(f"FILE LOCKED -> {alt}")


def main() -> int:
    log("GTI STEP4-2 NEWS AI v30 ISSUE-VERIFIED + WATCH-HANDOFF START")
    log(f"MODEL={GEMINI_MODEL} / Gemini={'Y' if USE_GEMINI else 'N'} / 24h / max={TARGET_MAX}")
    daily, audit, excluded = build()
    cumulative = merge_cumulative(daily)
    safe_write(OUT_SUMMARY, daily)
    safe_write(OUT_CUMULATIVE, cumulative)
    safe_write(OUT_AUDIT, audit)
    safe_write(OUT_EXCLUDED, excluded)
    safe_write(OUT_LEGACY, daily)
    log(f"DONE selected={len(daily)} / audit={len(audit)} / excluded={len(excluded)} / cumulative={len(cumulative)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
