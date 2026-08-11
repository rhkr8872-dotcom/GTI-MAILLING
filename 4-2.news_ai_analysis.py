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
import os, re, json, time
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
TARGET_MAX = min(int(os.getenv("GTI_STEP4_NEWS_TARGET_MAX", "30")), 30)
AI_REVIEW_MAX = int(os.getenv("GTI_STEP4_AI_REVIEW_MAX", "80"))
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
        f"{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}"
    )
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.15,
            "maxOutputTokens": 1500,
            "responseMimeType": "application/json",
        },
    }
    try:
        r = requests.post(endpoint, json=payload, timeout=GEMINI_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        txt = data["candidates"][0]["content"]["parts"][0]["text"]
        return json.loads(txt)
    except Exception as exc:
        return {"_error": f"{type(exc).__name__}:{clean(exc)[:180]}"}


def fallback_relevant(row: pd.Series) -> tuple[bool, int, str]:
    issue = clean(row.get("Issue")).upper()
    gate = clean(row.get("Gate")).upper()
    score = pre_score(row)
    relevant = gate == "CORE" or issue in {
        "TARIFF", "CUSTOMS", "ORIGIN_FTA", "EXPORT_CONTROL", "AD_CVD",
        "HS_CLASSIFICATION", "CBAM"
    }
    return relevant, min(100, max(50 if relevant else 0, score - 35)), "RULE_FALLBACK"


def analyze_row(row: pd.Series) -> dict:
    title = clean(row.get("Headline"))
    summary = clean(row.get("Summary"))
    issue = clean(row.get("Issue"))
    url = clean(row.get("URL"))

    prompt = f"""
당신은 삼성전자 본사 관세/통상 담당자입니다.
아래 기사가 삼성전자 관세업무 관점에서 보고할 가치가 있는지 판정하고 분석하십시오.

판정 기준:
- YES: 관세, 통관, HS/품목분류, FTA/원산지, AD/CVD/세이프가드,
  수출통제/제재, CBAM, 무역장벽/수입규제처럼 관세·통상 실무에 영향을 주는 기사.
- NO: 환율, 일반 수출실적, 주가, 고용, 산업동향, 행사, 지역경제 등 일반 뉴스.
- 제목에 '수입/수출/무역' 단어가 있다는 이유만으로 YES 금지.
- 삼성전자에 직접 영향이 없어도 주요 글로벌 관세정책이면 YES 가능.

JSON만 출력:
{{
 "relevant": true,
 "relevance_score": 0,
 "reason": "판정 근거",
 "samsung_impact": "Direct|Indirect|Watch|None",
 "affected_subsidiary": "영향 법인/지역 또는 관련 법인 검토",
 "risk": "상|중|하",
 "summary_ko": "기사 핵심 요약 2~3문장",
 "analysis_ko": "삼성전자 관세업무 영향 분석 2~4문장",
 "action_ko": "즉시/1주내/1개월내 조치와 담당",
 "country": "국가/지역",
 "agency": "기관/매체"
}}

Issue: {issue}
Headline: {title}
Source summary: {summary}
URL: {url}
""".strip()

    result = gemini_json(prompt)
    if not result or result.get("_error"):
        rel, rs, reason = fallback_relevant(row)
        return {
            "relevant": rel,
            "relevance_score": rs,
            "reason": reason + (f"; {result.get('_error')}" if result else ""),
            "samsung_impact": "Watch" if rel else "None",
            "affected_subsidiary": "관련 법인 검토",
            "risk": "중" if rel else "하",
            "summary_ko": summary or title,
            "analysis_ko": f"{issue or '관세·통상'} 관련성 확인 대상입니다." if rel else "",
            "action_ko": "대상국·품목·HS·시행일 및 법인 실적 영향을 확인하십시오." if rel else "",
            "country": "",
            "agency": clean(row.get("Publisher")) or clean(row.get("Source")),
        }
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
        r = row.to_dict()
        r["AIRelevant"] = "Y" if bool(a.get("relevant")) else "N"
        try:
            r["AIRelevanceScore"] = int(float(a.get("relevance_score", 0) or 0))
        except Exception:
            r["AIRelevanceScore"] = 0
        r["AIReason"] = clean(a.get("reason"))
        r["Samsung Impact"] = clean(a.get("samsung_impact")) or ("Watch" if r["AIRelevant"] == "Y" else "None")
        r["Affected Subsidiary"] = clean(a.get("affected_subsidiary")) or "관련 법인 검토"
        r["Risk"] = clean(a.get("risk")) or "중"
        r["SummaryAI"] = clean(a.get("summary_ko")) or clean(row.get("Summary"))
        r["AnalysisAI"] = clean(a.get("analysis_ko"))
        r["ActionAI"] = clean(a.get("action_ko"))
        r["Country"] = clean(a.get("country"))
        r["Agency"] = clean(a.get("agency")) or clean(row.get("Publisher"))
        audited.append(r)
        if i % 10 == 0 or i == len(review):
            log(f"AI REVIEW {i}/{len(review)}")

    audit = pd.DataFrame(audited)
    if audit.empty:
        audit = fresh.iloc[0:0].copy()
        audit["AIRelevant"] = ""
        audit["AIRelevanceScore"] = 0

    if not audit.empty:
        audit["SelectionScore"] = (
            pd.to_numeric(audit["PreScore"], errors="coerce").fillna(0) * 0.45
            + pd.to_numeric(audit["AIRelevanceScore"], errors="coerce").fillna(0) * 0.55
        ).round(1)

    selected = audit[audit["AIRelevant"].eq("Y")].copy()
    selected = selected.sort_values(
        ["SelectionScore", "PreScore", "Date"], ascending=[False, False, False], kind="stable"
    ).head(TARGET_MAX).reset_index(drop=True)

    selected_rows = []
    for i, r in selected.iterrows():
        impact = clean(r.get("Samsung Impact"))
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
            "Risk": clean(r.get("Risk")) or "중",
            "Importance Score": int(round(float(r.get("SelectionScore", 0) or 0))),
            "Priority Group": "CORE" if mail_group == "News - 핵심" else "USABLE",
            "Issue": clean(r.get("Issue")),
            "Cluster": clean(r.get("Cluster")),
            "URL": clean(r.get("URL")),
            "Source": clean(r.get("Source")),
            "Source File": str(INPUT_FILE),
            "RejectReason": "",
            "AIRelevant": "Y",
            "AIRelevanceScore": int(float(r.get("AIRelevanceScore", 0) or 0)),
            "AIReason": clean(r.get("AIReason")),
        })
    daily = pd.DataFrame(selected_rows, columns=OUTPUT_COLS)

    rejected_ai = audit[~audit["AIRelevant"].eq("Y")].copy()
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
    log("GTI STEP4-2 NEWS AI v25 CLEAN START")
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
