# -*- coding: utf-8 -*-
"""
GTI STEP5 Mail Engine - report quality form v2

Report form
-----------
1. 총평
2. Top3 Deep Analysis
3. Regulation
4. 주요뉴스

This step does not reselect STEP4 results. It keeps all selected regulation/news
items, then rewrites weak STEP4 text into an executive report style.
"""

from __future__ import annotations

import argparse
import html
import os
import re
import smtplib
import ssl
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


REGULATION_INPUT_FILE = Path(os.getenv("GTI_REGULATION_INPUT_FILE", r"C:\Temp\4-1.regulation_ai_summary.xlsx"))
NEWS_INPUT_FILE = Path(os.getenv("GTI_NEWS_INPUT_FILE", r"C:\Temp\4-2.news_ai_summary.xlsx"))
OUTPUT_DIR = Path(os.getenv("GTI_OUTPUT_DIR", r"C:\Temp\12345\c_type_outputs"))
RUN_DATE = os.getenv("GTI_RUN_DATE", datetime.now().strftime("%Y-%m-%d"))

NEWS_MAX_ROWS = int(os.getenv("GTI_NEWS_MAX_ROWS", "0"))  # 0 = no cap
SEND_EMAIL = os.getenv("GTI_SEND_EMAIL", "Y").strip().upper() in {"Y", "YES", "TRUE", "1"}
SMTP_HOST = os.getenv("GTI_SMTP_HOST", "smtp.naver.com")
SMTP_PORT = int(os.getenv("GTI_SMTP_PORT", "465"))
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or "").strip()
MAIL_TO = os.getenv("GTI_MAIL_TO", "").strip()
MAIL_FROM_NAME = os.getenv("GTI_MAIL_FROM_NAME", "GTI Radar").strip()
RECIPIENT_FILE = Path(os.getenv("GTI_RECIPIENT_FILE", r"C:\Temp\00.xlsx"))


OUTPUT_COLUMNS = [
    "No", "Content Type", "Mail Group", "Samsung Impact", "Affected Subsidiary", "Impact Reason",
    "Date", "Headline", "Major Changes", "Summary", "AI Analysis", "Action Plan", "Country", "Agency",
    "Risk", "Importance Score", "Priority Group", "Issue", "Cluster", "URL", "Source", "Source File",
]

GROUP_REGULATION = "Regulation"
GROUP_NEWS = "주요뉴스"


def output_paths() -> dict[str, Path]:
    return {
        "analysis": OUTPUT_DIR / "4.news_ai_analysis.xlsx",
        "mail_xlsx": OUTPUT_DIR / f"[GTI Radar] Global Trade Intelligence({RUN_DATE}).xlsx",
        "mail_html": OUTPUT_DIR / f"[GTI Radar] Global Trade Intelligence({RUN_DATE}).html",
        "cumulative": OUTPUT_DIR / "gti_news_cumulative.xlsx",
    }


def clean(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value)).strip()


def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lookup:
            return lookup[cand.lower()]
    return None


def safe_num(value) -> float:
    try:
        if value is None or pd.isna(value):
            return 0.0
    except Exception:
        pass
    text = re.sub(r"[^0-9.\-]", "", str(value))
    try:
        return float(text) if text else 0.0
    except Exception:
        return 0.0


def normalize_risk(value) -> str:
    raw = clean(value)
    low = raw.lower()
    if raw in {"상", "중", "하"}:
        return raw
    if low in {"high", "h", "red"}:
        return "상"
    if low in {"medium", "med", "m", "orange"}:
        return "중"
    if low in {"low", "l", "blue"}:
        return "하"
    return "중"


def risk_weight(value) -> int:
    return {"상": 300, "중": 150, "하": 0}.get(normalize_risk(value), 0)


def priority_weight(value) -> int:
    p = clean(value).upper()
    return {"CORE": 1000, "POLICY_WATCH": 850, "USABLE": 650, "REFERENCE": 300, "WATCH": 250}.get(p, 200)


def parse_date(value):
    dt = pd.to_datetime(value, errors="coerce")
    return pd.Timestamp.min if pd.isna(dt) else dt


def display_date(value) -> str:
    dt = pd.to_datetime(value, errors="coerce")
    if pd.isna(dt):
        return clean(value)[:16]
    if dt.hour == 0 and dt.minute == 0:
        return dt.strftime("%Y-%m-%d")
    return dt.strftime("%Y-%m-%d %H:%M")


def best_url_from_values(values) -> str:
    invalid = {
        "", "nan", "none", "null", "new", "https://new", "http://new",
        "https://news", "http://news", "https://news.google.com", "https://news.google.com/",
    }
    candidates: list[str] = []
    for value in values:
        text = clean(value)
        if not text:
            continue
        for item in [text] + re.findall(r"https?://[^'\"),\s]+", text):
            url = html.unescape(item).strip().strip("<>'\"").rstrip(".,);]}")
            if url.lower() in invalid:
                continue
            if re.match(r"^https?://", url, re.I) and url not in candidates:
                candidates.append(url)
    for url in candidates:
        low = url.lower()
        if "news.google.com/rss/articles/" not in low and "news.google.com/articles/" not in low:
            return url
    return candidates[0] if candidates else ""


def non_empty_hint(value: str) -> str:
    text = clean(value)
    if not text or text in {"본문에서 확인 불가", "nan", "None"}:
        return ""
    return text


def normalize_input(df: pd.DataFrame, content_type: str, source_file: Path) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_date = pick_col(df, ["Date", "date"])
    col_headline = pick_col(df, ["Headline", "Title", "headline"])
    col_country = pick_col(df, ["Country", "country"])
    col_agency = pick_col(df, ["Agency", "Publisher", "agency", "source"])
    col_risk = pick_col(df, ["Risk", "risk"])
    col_score = pick_col(df, ["Importance Score", "final_score", "samsung_score", "Score", "Importance"])
    col_priority = pick_col(df, ["Priority Group", "priority_group", "mail_section", "Tier"])
    col_issue = pick_col(df, ["Issue", "issue_type", "topic_keyword", "topic", "IssueKey"])
    col_cluster = pick_col(df, ["Cluster", "cluster_key", "ClusterHeadlines"])
    col_summary = pick_col(df, ["Summary", "summary", "ExecutiveMessage"])
    col_analysis = pick_col(df, ["AI Analysis", "analysis", "samsung_reason"])
    col_action = pick_col(df, ["Action Plan", "RequiredAction", "action"])
    col_source = pick_col(df, ["Source", "SourceFile", "source"])
    col_impact = pick_col(df, ["Samsung Impact", "samsung_impact"])
    col_subs = pick_col(df, ["Affected Subsidiary", "affected_subsidiary", "affected_subsidiaries"])
    col_reason = pick_col(df, ["Impact Reason", "subsidiary_reason", "samsung_reason", "SelectReason"])

    hint_cols = {
        "effective_date_hint": pick_col(df, ["effective_date_hint"]),
        "change_detail_hint": pick_col(df, ["change_detail_hint"]),
        "hs_hint": pick_col(df, ["hs_hint"]),
        "tariff_rate_hint": pick_col(df, ["tariff_rate_hint"]),
        "KeywordMatches": pick_col(df, ["KeywordMatches"]),
        "affected_products": pick_col(df, ["affected_products", "impact_products", "subsidiary_products"]),
        "fta_impact": pick_col(df, ["fta_impact"]),
        "export_control_impact": pick_col(df, ["export_control_impact"]),
        "hs_impact": pick_col(df, ["hs_impact"]),
        "tariff_impact": pick_col(df, ["tariff_impact"]),
    }

    url_cols = [
        pick_col(df, ["BestLinkURL"]),
        pick_col(df, ["OriginalURLCandidate"]),
        pick_col(df, ["original_url"]),
        pick_col(df, ["URL", "url", "Link"]),
        pick_col(df, ["GoogleURL"]),
        col_source,
    ]
    url_cols = [c for c in url_cols if c]

    out = pd.DataFrame()
    out["Date"] = df[col_date].apply(display_date) if col_date else ""
    out["_sort_date"] = df[col_date].apply(parse_date) if col_date else pd.Timestamp.min
    out["Headline"] = df[col_headline].apply(clean) if col_headline else ""
    out["Country"] = df[col_country].apply(clean) if col_country else ""
    out["Agency"] = df[col_agency].apply(clean) if col_agency else ""
    out["Risk"] = df[col_risk].apply(normalize_risk) if col_risk else "중"
    out["Importance Score"] = df[col_score].apply(safe_num) if col_score else 0
    out["Priority Group"] = df[col_priority].apply(lambda v: clean(v).upper()) if col_priority else ("CORE" if content_type == "Regulation" else "USABLE")
    out["Issue"] = df[col_issue].apply(clean) if col_issue else ""
    out["Issue"] = out["Issue"].replace({
        "TARIFF": "관세정책", "SECTION_301_232": "관세정책",
        "CUSTOMS": "통관", "CUSTOMS_CLEARANCE": "통관",
        "ORIGIN_FTA": "FTA/원산지", "CBAM_CARBON": "CBAM",
        "HS_CLASSIFICATION": "HS/품목분류", "AD_CVD": "AD/CVD",
        "EXPORT_CONTROL": "수출통제",
    })
    out["Cluster"] = df[col_cluster].apply(clean) if col_cluster else ""
    out["Summary"] = df[col_summary].apply(clean) if col_summary else ""
    out["AI Analysis"] = df[col_analysis].apply(clean) if col_analysis else ""
    out["Action Plan"] = df[col_action].apply(clean) if col_action else ""
    out["Samsung Impact"] = df[col_impact].apply(lambda v: clean(v).title() if clean(v).lower() in {"direct", "indirect", "watch"} else clean(v)) if col_impact else "Watch"
    out["Samsung Impact"] = out["Samsung Impact"].replace({"": "Watch", "직접": "Direct", "간접": "Indirect", "모니터링": "Watch"})
    out["Affected Subsidiary"] = df[col_subs].apply(clean) if col_subs else ""
    out["Impact Reason"] = df[col_reason].apply(clean) if col_reason else ""
    out["Source"] = df[col_source].apply(clean) if col_source else ""
    out["Source File"] = str(source_file)
    out["Content Type"] = content_type
    out["Mail Group"] = GROUP_REGULATION if content_type == "Regulation" else GROUP_NEWS
    out["URL"] = df.apply(lambda r: best_url_from_values([r.get(c, "") for c in url_cols]), axis=1) if len(df) else ""

    for out_col, src_col in hint_cols.items():
        out[out_col] = df[src_col].apply(clean) if src_col else ""

    out = out[out["Headline"].astype(str).str.strip().ne("")]
    return out.reset_index(drop=True)


def read_step4_results() -> pd.DataFrame:
    frames = []
    if REGULATION_INPUT_FILE.exists():
        frames.append(normalize_input(pd.read_excel(REGULATION_INPUT_FILE), "Regulation", REGULATION_INPUT_FILE))
    if NEWS_INPUT_FILE.exists():
        news = normalize_input(pd.read_excel(NEWS_INPUT_FILE), "News", NEWS_INPUT_FILE)
        if NEWS_MAX_ROWS > 0:
            news = news.head(NEWS_MAX_ROWS)
        frames.append(news)
    if not frames:
        raise FileNotFoundError(f"STEP4 outputs not found: {REGULATION_INPUT_FILE}, {NEWS_INPUT_FILE}")

    rows = pd.concat(frames, ignore_index=True)
    rows["_dedup_key"] = rows.apply(
        lambda r: clean(r.get("URL")) or (clean(r.get("Headline"))[:160] + "|" + clean(r.get("Agency")) + "|" + clean(r.get("Date"))),
        axis=1,
    )
    rows = rows.drop_duplicates(subset=["_dedup_key"], keep="first").drop(columns=["_dedup_key"], errors="ignore")
    rows["_integrated_score"] = rows.apply(
        lambda r: priority_weight(r["Priority Group"]) + risk_weight(r["Risk"]) + (180 if r["Content Type"] == "Regulation" else 0) + safe_num(r["Importance Score"]),
        axis=1,
    )
    return rows.reset_index(drop=True)


def dedup_report_rows(rows: pd.DataFrame) -> pd.DataFrame:
    """Remove near-duplicate report items after STEP4 merge.

    This intentionally catches cases where the same policy appears through two
    official pages, such as CBAM certificate price "to be published" and
    "now available", or the same bonded warehouse notice from two boards.
    """
    if rows.empty:
        return rows
    rows = rows.copy()
    rows["_report_dedup_key"] = rows.apply(report_dedup_key, axis=1)
    rows["_dedup_rank"] = rows.apply(dedup_rank, axis=1)
    rows = rows.sort_values(["_dedup_rank", "_integrated_score", "_sort_date"], ascending=[False, False, False])
    rows = rows.drop_duplicates(subset=["_report_dedup_key"], keep="first")
    return rows.drop(columns=["_report_dedup_key", "_dedup_rank"], errors="ignore").reset_index(drop=True)


def report_dedup_key(row: pd.Series) -> str:
    issue = clean(row.get("Issue")) or issue_for(row)
    title = clean(row.get("Headline")).lower()
    source = clean(row.get("Agency")).lower()
    normalized = re.sub(r"\([^)]*\)|\[[^]]*\]", " ", title)
    normalized = re.sub(r"제출기한[:：]?\s*\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}.*", " ", normalized)
    normalized = re.sub(r"\b(to be published|now available|published|available|first)\b", " ", normalized)
    normalized = re.sub(r"[^0-9a-z가-힣]+", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()

    if "cbam" in normalized and "certificate price" in normalized:
        return "REG:CBAM_CERTIFICATE_PRICE"
    if "보세창고" in normalized and "특허" in normalized and "운영" in normalized:
        return "REG:BONDED_WAREHOUSE_LICENSE_OPERATION"
    if "환전영업자" in normalized and "관리" in normalized:
        return "REG:FX_BUSINESS_OPERATOR_MANAGEMENT"
    if "수입신고" in normalized and "가산세" in normalized:
        return "REG:IMPORT_DECLARATION_DELAY_SURCHARGE"

    tokens = [t for t in normalized.split() if len(t) >= 2]
    return f"{clean(row.get('Content Type'))}:{issue}:{' '.join(tokens[:9])}:{source[:24]}"


def dedup_rank(row: pd.Series) -> float:
    title = clean(row.get("Headline")).lower()
    rank = safe_num(row.get("Importance Score")) + priority_weight(row.get("Priority Group")) + risk_weight(row.get("Risk"))
    if "now available" in title or "successfully entered into force" in title:
        rank += 150
    if "to be published" in title or "reminder" in title:
        rank -= 80
    if "제출기한" in title:
        rank += 60
    if clean(row.get("Agency")).startswith("Korea Customs") or "관세청" in clean(row.get("Agency")):
        rank += 40
    return rank


def issue_for(row) -> str:
    issue = clean(row.get("Issue"))
    if issue and issue.lower() not in {"watch", "policy_watch", "usable", "core"}:
        return issue
    text = " ".join(clean(row.get(c)) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "KeywordMatches"]).lower()
    if any(k in text for k in ["section 301", "section 232", "tariff", "quota", "duty", "관세", "쿼터"]):
        return "관세정책"
    if any(k in text for k in ["anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "상계관세"]):
        return "AD/CVD"
    if any(k in text for k in ["cbam", "carbon border"]):
        return "CBAM"
    if any(k in text for k in ["fta", "cepa", "origin", "원산지"]):
        return "FTA/원산지"
    if any(k in text for k in ["export control", "entity list", "uflpa", "forced labor", "수출통제"]):
        return "수출통제"
    if any(k in text for k in ["hs code", "classification", "품목분류"]):
        return "HS/품목분류"
    if clean(row.get("Content Type")) == "Regulation":
        return "법규"
    return "Watch"


def prepare_rows(rows: pd.DataFrame) -> pd.DataFrame:
    rows = rows.copy()
    rows["Issue"] = rows.apply(issue_for, axis=1)
    rows = dedup_report_rows(rows)
    rows["Mail Group"] = rows["Content Type"].map({"Regulation": GROUP_REGULATION}).fillna(GROUP_NEWS)
    rows["Major Changes"] = rows.apply(major_changes, axis=1)
    rows["Summary"] = rows.apply(report_summary, axis=1)
    rows["AI Analysis"] = rows.apply(report_impact, axis=1)
    rows["Action Plan"] = rows.apply(report_action, axis=1)
    rows["_report_score"] = rows.apply(report_score, axis=1)
    rows = rows.sort_values(["_integrated_score", "_sort_date"], ascending=[False, False]).reset_index(drop=True)
    rows["No"] = range(1, len(rows) + 1)
    return rows


def choose_top3(rows: pd.DataFrame) -> pd.DataFrame:
    pool = rows.copy()
    pool["_top3_score"] = pool.apply(top3_deep_score, axis=1)
    pool = pool.sort_values(["_top3_score", "_sort_date"], ascending=[False, False])
    selected = []
    used_issues = set()
    for _, row in pool.iterrows():
        issue = clean(row.get("Issue"))
        if issue in used_issues and len(selected) < 3:
            continue
        selected.append(row)
        used_issues.add(issue)
        if len(selected) == 3:
            break
    if len(selected) < 3:
        for _, row in pool.iterrows():
            if any(clean(row.get("Headline")) == clean(x.get("Headline")) for x in selected):
                continue
            selected.append(row)
            if len(selected) == 3:
                break
    out = pd.DataFrame(selected).reset_index(drop=True)
    if not out.empty:
        out["No"] = range(1, len(out) + 1)
    return out


def top3_deep_score(row: pd.Series) -> float:
    text = " ".join(clean(row.get(c)) for c in [
        "Headline", "Major Changes", "Summary", "AI Analysis", "Action Plan", "Issue"
    ]).lower()
    score = report_score(row)

    high_terms = [
        "anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "상계관세",
        "cbam", "carbon border", "탄소국경", "탄소세",
        "export control", "entity list", "forced labor", "uflpa", "수출통제", "강제노동",
        "section 301", "section 232", "tariff quota", "duty-free quota", "customs duty",
        "관세", "무관세", "쿼터", "원산지", "rules of origin", "hs code",
    ]
    medium_terms = [
        "fta", "cepa", "usmca", "통상협정", "fta", "통관", "보세", "신고", "classification",
    ]
    low_terms = [
        "수출 85.9", "수출입 현황", "잠정치", "재정적자", "refunds", "customs revenue",
        "브랜드", "주식", "전략회의", "칼럼", "market outlook",
    ]

    if any(t in text for t in high_terms):
        score += 1200
    if any(t in text for t in medium_terms):
        score += 450
    if clean(row.get("Content Type")) == "Regulation":
        score += 300
    if any(t in text for t in low_terms):
        score -= 900
    return score


def report_score(row: pd.Series) -> float:
    impact_weight = {"Direct": 2200, "Indirect": 900, "Watch": 0}.get(clean(row.get("Samsung Impact")), 0)
    type_weight = 350 if clean(row.get("Content Type")) == "Regulation" else 0
    issue_weight = {
        "관세정책": 500,
        "AD/CVD": 500,
        "반덤핑/상계관세": 500,
        "CBAM": 450,
        "수출통제": 450,
        "FTA/원산지": 350,
        "통관": 300,
        "통관/세관": 300,
        "HS/품목분류": 300,
    }.get(clean(row.get("Issue")), 150)
    return safe_num(row.get("Importance Score")) + priority_weight(row.get("Priority Group")) + risk_weight(row.get("Risk")) + impact_weight + type_weight + issue_weight


def hint_line(label: str, value: str) -> str:
    value = non_empty_hint(value)
    return f"{label}: {value}" if value else ""


def compact_parts(parts: list[str], fallback: str) -> str:
    parts = [p for p in parts if clean(p)]
    return "; ".join(parts) if parts else fallback


def major_changes(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    headline = clean(row.get("Headline"))
    title_l = headline.lower()

    if "보세창고" in headline and "특허" in headline:
        return (
            "개정 사유: 자가용보세창고 특허요건 완화 및 불명확한 규정 보완 필요. "
            "주요 개정 내용: 자가용보세창고 반입 대상에 국제무역선·기 적재 자가화물 외 수리용 예비부분품 및 부속품 장치를 허용하고, "
            "관세법 제178조상 물품반입 정지기간을 오해 없이 적용할 수 있도록 규정을 명확화하는 내용입니다."
        )
    if "환전영업자" in headline and "관리" in headline:
        return (
            "주요 내용: 환전영업자의 등록·관리, 보고·자료제출, 영업장 운영 및 관세청 관리 기준과 관련된 고시입니다. "
            "해외출장·주재원·외환거래 지원 프로세스와 연결될 수 있어 실제 법인 업무 해당 여부 확인이 필요합니다."
        )
    if "cbam" in title_l and "certificate price" in title_l:
        return (
            "주요 내용: EU CBAM 인증서 가격이 공표되었거나 공표 일정이 확정된 사안입니다. "
            "EU 수입품의 내재배출량 신고, 인증서 구매 비용, 공급사 배출량 자료 확보 체계에 영향을 줄 수 있습니다."
        )
    if "customs enforcement" in title_l and "executive order" in title_l:
        return (
            "주요 내용: 미국 세관 집행 강화 행정명령 관련 사안입니다. "
            "수입신고 정확성, 저가신고·우회수입·전자상거래 물품 관리 및 CBP 심사 강화 가능성을 확인해야 합니다."
        )
    if "수입신고" in headline and "가산세" in headline:
        return (
            "주요 내용: 수입신고 지연 가산세 부과 대상이 되는 매점매석 금지 품목의 적용기간 연장 공고입니다. "
            "해당 품목 수입 시 신고 지연, 재고 운영, 통관 일정 관리 기준을 확인해야 합니다."
        )

    parts = [
        hint_line("시행/적용일", row.get("effective_date_hint")),
        hint_line("변경 내용", row.get("change_detail_hint")),
        hint_line("대상 HS", row.get("hs_hint")),
        hint_line("관세율/쿼터", row.get("tariff_rate_hint")),
        hint_line("키워드", row.get("KeywordMatches")),
    ]
    if any(parts):
        return compact_parts(parts, "")

    if issue == "관세정책":
        return "관세율, 쿼터, 면세/환급 또는 Section 301/232 등 관세 비용에 영향을 줄 수 있는 정책 변화입니다."
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return "반덤핑 또는 상계관세 조사·판정·연장 가능성이 있는 사안입니다. 공급국, 대상 품목, 조사 기간과 관세율 확인이 필요합니다."
    if issue == "CBAM":
        return "CBAM 신고, 인증서 가격, 배출량 자료 또는 EU 수입통관 절차와 연결되는 탄소국경조정 변화입니다."
    if issue == "FTA/원산지":
        return "FTA/CEPA 협정, 원산지 기준, CO 발급 또는 특혜관세 적용 가능성에 영향을 주는 변화입니다."
    if issue == "수출통제":
        return "Entity List, ECCN, UFLPA, forced labor 또는 전략물자·제재 스크리닝 관련 변화입니다."
    if issue == "통관":
        return "보세, 통관, 신고, 세관 심사 또는 행정절차 기준에 영향을 줄 수 있는 공식 공지입니다."
    if issue == "HS/품목분류":
        return "HS 분류 기준 또는 품목 해석이 달라질 수 있어 품목 마스터와 신고 기준 점검이 필요한 사안입니다."
    return f"{headline} 관련 관세·통상 모니터링 사안입니다."


def report_summary(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    country = clean(row.get("Country")) or "관련 국가"
    agency = clean(row.get("Agency")) or "관련 기관"
    change = major_changes(row)
    if clean(row.get("Content Type")) == "Regulation":
        return f"{agency}의 공식 법규/공지입니다. 핵심은 {change} 원문 기준으로 시행일, 적용 품목, HS, 세율 또는 신고 절차를 확인해야 합니다."
    return f"{country}에서 포착된 {issue} 뉴스입니다. 핵심은 {change} 삼성전자 관련 법인·품목에 직접 적용되는지 확인할 필요가 있습니다."


def report_impact(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    impact = clean(row.get("Samsung Impact")) or "Watch"
    subs = clean(row.get("Affected Subsidiary")) or "SEC/HQ"
    products = non_empty_hint(row.get("affected_products"))
    product_txt = f" 대상 제품 후보는 {products}입니다." if products else ""
    if issue == "관세정책":
        return f"{subs} 기준 수입가격, 관세환급, 할당관세/쿼터, 공급국 선택에 영향을 줄 수 있습니다.{product_txt} Impact는 {impact}로 분류됩니다."
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return f"{subs}의 철강·부품·원재료 조달에서 AD/CVD 추가관세 또는 조사 대응 자료 부담이 생길 수 있습니다.{product_txt} 공급국과 HS별 노출도를 확인해야 합니다."
    if issue == "CBAM":
        return f"{subs}의 EU향 판매·공급망에서 CBAM 신고자료, 배출량 증빙, 인증서 비용 관리가 필요할 수 있습니다.{product_txt}"
    if issue == "FTA/원산지":
        return f"{subs}의 FTA 활용, 원산지 판정, CO 발급, BOM 원산지 증빙 체계에 영향을 줄 수 있습니다.{product_txt}"
    if issue == "수출통제":
        return f"{subs}의 거래처 스크리닝, ECCN/전략물자 분류, 우회수출 통제와 연결될 수 있습니다.{product_txt}"
    if issue == "통관":
        return f"{subs}의 수입신고, 보세창고, 통관 심사, 세관 제출자료 운영 기준에 반영 여부를 확인해야 합니다.{product_txt}"
    if issue == "HS/품목분류":
        return f"{subs}의 HS 마스터, 품목 설명, 관세율 산정 및 신고 정확성에 영향을 줄 수 있습니다.{product_txt}"
    return f"{subs} 기준 관세·통상 리스크 모니터링 가치가 있습니다. Impact는 {impact}입니다."


def report_action(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    subs = clean(row.get("Affected Subsidiary")) or "SEC/HQ"
    if issue == "관세정책":
        return f"{subs}: 대상 HS·공급국·거래금액을 매핑하고 세율/쿼터/환급 가능성을 산출해 관세비용 영향표에 반영하십시오."
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return f"{subs}: 대상 품목·공급국·벤더를 확인하고 조사대응 자료, 원산지 증빙, 가격자료 보관 필요성을 점검하십시오."
    if issue == "CBAM":
        return f"{subs}: EU향 품목, 공급사 배출량 자료, CBAM 신고·인증서 비용 반영 여부를 ESG/구매/통관 담당과 확인하십시오."
    if issue == "FTA/원산지":
        return f"{subs}: BOM 원산지, CO 발급, 직접운송, 누적기준, 특혜세율 적용 가능성을 FTA 마스터와 대조하십시오."
    if issue == "수출통제":
        return f"{subs}: ECCN/전략물자 분류, 거래처·최종사용자 스크리닝, 제재국 우회거래 가능성을 재점검하십시오."
    if issue == "통관":
        return f"{subs}: 통관 SOP, 보세/신고 체크리스트, 관세사 안내문, 세관 제출자료 양식을 업데이트하십시오."
    if issue == "HS/품목분류":
        return f"{subs}: 관련 제품의 HS 설명, 판정 근거, 해외법인 신고코드와 한국 본사 마스터 간 차이를 점검하십시오."
    return f"{subs}: 원문 기준으로 대상 국가, 품목, 시행일, 담당 부서를 확인하고 후속 모니터링하십시오."


def html_link(title: str, url: str) -> str:
    title_e = html.escape(clean(title))
    url = best_url_from_values([url])
    if not url:
        return title_e
    return f'<a href="{html.escape(url)}" target="_blank">{title_e}</a>'


def risk_color(risk: str) -> str:
    return {"상": "#C00000", "중": "#C55A11", "하": "#4472C4"}.get(normalize_risk(risk), "#555")


def short_text(value, fallback: str, limit: int = 360) -> str:
    text = clean(value) or fallback
    return text[:limit] + ("..." if len(text) > limit else "")


def one_line(row: pd.Series) -> str:
    return f"{clean(row.get('Issue'))} / {clean(row.get('Country')) or '-'} / {clean(row.get('Samsung Impact'))}: {short_text(row.get('Major Changes'), '주요 변경내역 확인 필요', 130)}"


def top3_summary_sentence(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return "반덤핑·상계관세 조사는 대상 품목의 HS, 원산지, 가격자료 방어체계를 우선 점검해야 합니다."
    if issue == "FTA/원산지":
        return "FTA·원산지 변경은 관련 법인의 CO 발급요건과 특혜세율 적용 가능성을 재검토해야 합니다."
    if issue in {"관세정책", "통관", "통관/세관", "HS/품목분류"}:
        return "관세·통상 정책 변화는 관련 품목의 HS, 원산지, 관세율 영향을 확인해야 합니다."
    if issue == "CBAM":
        return "CBAM 변화는 EU향 품목의 배출량 자료, 인증서 비용, 신고의무 반영 여부를 우선 점검해야 합니다."
    if issue == "수출통제":
        return "수출통제 변화는 ECCN·전략물자 분류와 거래처·최종사용자 스크리닝 체계를 우선 확인해야 합니다."
    return short_text(row.get("Major Changes"), "관세·통상 영향 여부를 원문 기준으로 확인해야 합니다.", 150)


def top3_summary_rows(rows: pd.DataFrame, top3: pd.DataFrame) -> pd.DataFrame:
    """Issue-level summary rows for the executive summary.

    The executive summary should be thematic and practical, not just repeat the
    three article titles. Prefer AD/CVD, FTA/origin, tariff/customs if present.
    """
    preferred_groups = [
        {"AD/CVD", "반덤핑/상계관세"},
        {"FTA/원산지", "ORIGIN", "원산지"},
        {"관세정책", "통관", "통관/세관", "HS/품목분류"},
        {"CBAM"},
        {"수출통제"},
    ]
    selected = []
    used = set()
    for group in preferred_groups:
        cand = rows[rows["Issue"].astype(str).isin(group)].copy()
        if cand.empty:
            continue
        cand = cand.sort_values(["_report_score", "_sort_date"], ascending=[False, False])
        row = cand.iloc[0]
        key = clean(row.get("Issue"))
        if key in used:
            continue
        selected.append(row)
        used.add(key)
        if len(selected) >= 3:
            break
    if len(selected) < 3:
        for _, row in top3.iterrows():
            key = clean(row.get("Issue"))
            if key in used:
                continue
            selected.append(row)
            used.add(key)
            if len(selected) >= 3:
                break
    return pd.DataFrame(selected)


def overall_html(rows: pd.DataFrame, top3: pd.DataFrame) -> str:
    reg = rows[rows["Content Type"].eq("Regulation")]
    news = rows[rows["Content Type"].eq("News")]
    direct = rows[rows["Samsung Impact"].eq("Direct")]
    indirect = rows[rows["Samsung Impact"].eq("Indirect")]
    watch = rows[rows["Samsung Impact"].eq("Watch")]
    issues = rows["Issue"].value_counts().head(6)
    issue_txt = ", ".join(f"{k} {v}건" for k, v in issues.items())
    summary_rows = top3_summary_rows(rows, top3)
    top_lines = "".join(f"<li>{html.escape(top3_summary_sentence(r))}</li>" for _, r in summary_rows.iterrows())
    return f"""
    <div style="padding:15px;background:#F4F6F8;border-left:6px solid #1F4E78;margin-bottom:18px;">
      <div style="font-size:14px;color:#555;margin-bottom:8px;">
        금일 선별 결과: 법규 {len(reg)}건, 주요뉴스 {len(news)}건 | Direct {len(direct)}건, Indirect {len(indirect)}건, Watch {len(watch)}건
      </div>
      <div style="font-size:15px;font-weight:bold;line-height:1.8;margin-bottom:8px;">
        금일 GTI Radar는 {html.escape(issue_txt)} 중심으로 관세·통상 변화가 포착되었습니다. 법규는 시행일·HS·세율·신고절차 반영 여부를, 뉴스는 실제 비용·원산지·수출통제 영향 가능성을 우선 확인해야 합니다.
      </div>
      <div style="margin-top:8px;"><b>Top3 요약</b><ol style="margin-top:6px;">{top_lines}</ol></div>
    </div>
    """


def top3_html(top3: pd.DataFrame) -> str:
    blocks = []
    for idx, row in top3.iterrows():
        blocks.append(f"""
        <div style="margin:14px 0 18px 0;padding:15px;border-left:5px solid #C00000;background:#FFF7F7;">
          <div style="font-size:15px;font-weight:bold;margin-bottom:6px;">Top {idx + 1}. {html_link(row.get('Headline'), row.get('URL'))}</div>
          <div style="font-size:12px;color:#555;margin-bottom:9px;">
            Type: {html.escape(clean(row.get('Content Type')))} | Topic: {html.escape(clean(row.get('Issue')))} |
            Samsung Impact: <b>{html.escape(clean(row.get('Samsung Impact')))}</b> |
            Subsidiary: {html.escape(clean(row.get('Affected Subsidiary')) or 'SEC/HQ')} |
            Agency: {html.escape(clean(row.get('Agency')))} | Publish Date: {html.escape(clean(row.get('Date')))} |
            Country: {html.escape(clean(row.get('Country')))} |
            Risk: <span style="color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</span> |
            Score: {safe_num(row.get('Importance Score')):.0f}
          </div>
          <div style="margin-top:8px;"><b>Executive Impact</b><br>{html.escape(one_line(row))}</div>
          <div style="margin-top:8px;"><b>주요 변경내역</b><br>{html.escape(short_text(row.get('Major Changes'), '주요 변경내역 확인 필요', 520))}</div>
          <div style="margin-top:8px;"><b>삼성 영향</b><br>{html.escape(short_text(row.get('AI Analysis'), '삼성 영향 검토 필요', 520))}</div>
          <div style="margin-top:8px;"><b>Action</b><br>{html.escape(short_text(row.get('Action Plan'), '담당 부서 확인 필요', 520))}</div>
        </div>
        """)
    return "".join(blocks)


def table_html(title: str, rows: pd.DataFrame, color: str) -> str:
    if rows.empty:
        return f"<h3 style='color:{color};'>{html.escape(title)} (0건)</h3>"
    trs = []
    for _, row in rows.iterrows():
        trs.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(str(row.get('No')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Issue')))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html_link(row.get('Headline'), row.get('URL'))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('Major Changes'), '주요 변경내역 확인 필요', 260))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('AI Analysis'), '영향 검토 필요', 260))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('Action Plan'), '담당 부서 확인 필요', 260))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Country')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Samsung Impact')))}</td>
        </tr>
        """)
    return f"""
    <h3 style="margin-top:24px;color:{color};">{html.escape(title)} ({len(rows)}건)</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;table-layout:fixed;">
      <colgroup>
        <col style="width:3%;"><col style="width:8%;"><col style="width:22%;">
        <col style="width:22%;"><col style="width:18%;"><col style="width:18%;">
        <col style="width:5%;"><col style="width:4%;"><col style="width:5%;">
      </colgroup>
      <thead>
        <tr style="background:{color};color:white;">
          <th style="padding:7px;border:1px solid #ddd;">No</th>
          <th style="padding:7px;border:1px solid #ddd;">Issue</th>
          <th style="padding:7px;border:1px solid #ddd;">Headline</th>
          <th style="padding:7px;border:1px solid #ddd;">주요 변경내역</th>
          <th style="padding:7px;border:1px solid #ddd;">삼성 영향</th>
          <th style="padding:7px;border:1px solid #ddd;">Action</th>
          <th style="padding:7px;border:1px solid #ddd;">Country</th>
          <th style="padding:7px;border:1px solid #ddd;">Risk</th>
          <th style="padding:7px;border:1px solid #ddd;">Impact</th>
        </tr>
      </thead>
      <tbody>{''.join(trs)}</tbody>
    </table>
    """


def build_html(rows: pd.DataFrame, top3: pd.DataFrame) -> str:
    subject = f"[GTI Radar] Global Trade Intelligence | {RUN_DATE}"
    regulation = rows[rows["Content Type"].eq("Regulation")]
    news = rows[rows["Content Type"].eq("News")]
    return f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"><title>{html.escape(subject)}</title></head>
<body style="font-family:Arial,'Malgun Gothic',sans-serif;font-size:13px;color:#222;line-height:1.55;">
  <div style="max-width:1320px;margin:0 auto;">
    <h2 style="margin-bottom:3px;color:#1F4E78;">[GTI Radar] Global Trade Intelligence</h2>
    <div style="font-size:13px;color:#555;margin-bottom:16px;">{RUN_DATE} | Samsung Electronics Customs & Trade Intelligence</div>

    <h3 style="margin-top:18px;margin-bottom:6px;">1. 총평</h3>
    {overall_html(rows, top3)}

    <h3 style="margin-top:22px;color:#C00000;">2. Top3 Deep Analysis</h3>
    {top3_html(top3)}

    {table_html('3. Regulation', regulation, '#1F4E78')}
    {table_html('4. 주요뉴스', news, '#548235')}

    <p style="margin-top:18px;color:#666;font-size:12px;">첨부 Excel에는 전체 선별 결과와 원문 링크가 포함되어 있습니다.</p>
  </div>
</body>
</html>"""


def style_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
    widths = {
        "A": 5, "B": 13, "C": 13, "D": 13, "E": 18, "F": 30, "G": 15, "H": 48,
        "I": 46, "J": 46, "K": 46, "L": 46, "M": 14, "N": 18, "O": 8, "P": 12,
        "Q": 16, "R": 16, "S": 24, "T": 36, "U": 24, "V": 28,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"


def append_output_row(ws, row: pd.Series) -> None:
    ws.append([row.get(c, "") for c in OUTPUT_COLUMNS])
    headline_col = OUTPUT_COLUMNS.index("Headline") + 1
    cell = ws.cell(row=ws.max_row, column=headline_col)
    url = best_url_from_values([row.get("URL")])
    if url:
        cell.hyperlink = url
        cell.font = Font(color="0563C1", underline="single", bold=True)


def save_excel(rows: pd.DataFrame, top3: pd.DataFrame, paths: dict[str, Path]) -> None:
    wb = Workbook()
    sheets = [
        ("GTI Radar", rows),
        ("Top3 Deep Analysis", top3),
        ("Regulation", rows[rows["Content Type"].eq("Regulation")]),
        ("주요뉴스", rows[rows["Content Type"].eq("News")]),
    ]
    first = True
    for name, frame in sheets:
        ws = wb.active if first else wb.create_sheet(name[:31])
        first = False
        ws.title = name[:31]
        ws.append(OUTPUT_COLUMNS)
        for _, row in frame.iterrows():
            append_output_row(ws, row)
        style_sheet(ws)

    runlog = wb.create_sheet("Run Log")
    runlog.append(["item", "value"])
    runlog.append(["regulation_input", str(REGULATION_INPUT_FILE)])
    runlog.append(["news_input", str(NEWS_INPUT_FILE)])
    runlog.append(["run_date", RUN_DATE])
    runlog.append(["total_rows", len(rows)])
    runlog.append(["regulation_rows", int(rows["Content Type"].eq("Regulation").sum())])
    runlog.append(["news_rows", int(rows["Content Type"].eq("News").sum())])
    runlog.append(["direct_rows", int(rows["Samsung Impact"].eq("Direct").sum())])
    runlog.append(["indirect_rows", int(rows["Samsung Impact"].eq("Indirect").sum())])
    runlog.append(["watch_rows", int(rows["Samsung Impact"].eq("Watch").sum())])
    style_sheet(runlog)

    wb.save(paths["mail_xlsx"])
    wb.save(paths["analysis"])
    rows[OUTPUT_COLUMNS].to_excel(paths["cumulative"], index=False)


def read_recipients() -> list[str]:
    recipients = []
    if MAIL_TO:
        recipients.extend([x.strip() for x in re.split(r"[;,]", MAIL_TO) if x.strip()])
    if RECIPIENT_FILE.exists():
        try:
            df = pd.read_excel(RECIPIENT_FILE)
            for col in df.columns:
                for value in df[col].dropna().astype(str):
                    text = clean(value)
                    if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", text):
                        recipients.append(text)
        except Exception:
            pass
    seen, out = set(), []
    for email in recipients:
        low = email.lower()
        if low not in seen:
            seen.add(low)
            out.append(email)
    return out


def send_email(html_body: str, attachment: Path) -> None:
    if not SEND_EMAIL:
        print("[MAIL SKIP] GTI_SEND_EMAIL=N or --no-email")
        return
    recipients = read_recipients()
    if not recipients:
        print("[MAIL SKIP] recipients missing")
        return
    if not SMTP_USER or not SMTP_PASS:
        print("[MAIL SKIP] SMTP credential missing")
        return

    msg = EmailMessage()
    msg["Subject"] = f"[GTI Radar] Global Trade Intelligence({RUN_DATE})"
    msg["From"] = formataddr((MAIL_FROM_NAME, SMTP_USER))
    msg["To"] = ", ".join(recipients)
    msg.set_content("GTI Radar report is attached. HTML mail requires an HTML-capable client.")
    msg.add_alternative(html_body, subtype="html")
    data = attachment.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment.name,
    )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=30) as server:
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    print(f"[MAIL SENT] {len(recipients)} recipients")



# ======================================================================
# GTI STEP5 Executive Quality Patch v3
# ----------------------------------------------------------------------
# Purpose
# 1) Re-rank Top3 by Samsung relevance + customs actionability, not keywords only
# 2) Auto-demote low-relevance items such as wheat/agriculture/general economy to REFERENCE
# 3) Expand Top3 Deep Analysis into issue summary / Samsung impact / customs impact / risk / action / owner
# 4) Keep STEP4 selected rows, but improve mail report quality at STEP5
# ======================================================================

SAMSUNG_RELEVANCE_TERMS = {
    "high": [
        "semiconductor", "chip", "memory", "dram", "nand", "hbm", "foundry", "wafer",
        "반도체", "메모리", "파운드리", "웨이퍼",
        "display", "oled", "lcd", "디스플레이",
        "battery", "cell", "cathode", "anode", "lithium", "nickel", "cobalt", "graphite",
        "배터리", "양극재", "음극재", "리튬", "니켈", "코발트", "흑연",
        "electronics", "smartphone", "mobile", "tv", "appliance", "home appliance",
        "전자", "스마트폰", "모바일", "가전", "tv",
        "pcb", "substrate", "module", "camera module", "sensor", "mlcc",
        "기판", "모듈", "센서", "mlcc",
        "rare earth", "gallium", "germanium", "gan", "silicon carbide", "sic",
        "희토류", "갈륨", "게르마늄", "전략물자",
        "steel", "aluminum", "copper", "zinc", "cold-rolled", "galvanized",
        "철강", "알루미늄", "구리", "아연", "냉간압연", "도금강판",
    ],
    "medium": [
        "customs", "tariff", "duty", "origin", "fta", "cbam", "ad/cvd", "anti-dumping",
        "countervailing", "hs code", "classification", "importer", "export control",
        "forced labor", "uflpa", "section 301", "section 232", "usmca", "cepa",
        "관세", "통관", "원산지", "수출통제", "강제노동", "반덤핑", "상계관세",
        "품목분류", "수입자", "보세", "수입신고", "수출신고",
    ],
}

LOW_RELEVANCE_TERMS = [
    "wheat", "rice", "corn", "soybean", "sugar", "agriculture", "agricultural",
    "livestock", "pork", "beef", "fishery", "food", "grain", "flour",
    "밀", "쌀", "옥수수", "농산물", "농업", "축산", "식품", "곡물", "밀가루",
    "염소산업", "혈통관리", "소비자물가", "주식", "배당", "concert", "sports",
]

GENERAL_NEWS_TERMS = [
    "수출 85.9", "수출입 현황", "증시", "주가", "배당", "실적", "gdp", "환율",
    "market outlook", "stock", "shares", "dividend", "budget deficit",
]

CRITICAL_CUSTOMS_TERMS = [
    "anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "상계관세",
    "cbam", "carbon border", "탄소국경",
    "customs enforcement", "cbp", "ior", "importer of record", "bond",
    "관세청", "세관", "수입신고", "보세", "통관",
    "export control", "entity list", "strategic goods", "dual-use", "수출통제", "전략물자",
    "forced labor", "uflpa", "강제노동",
    "hs code", "classification", "품목분류", "hscode",
    "section 301", "section 232", "tariff", "관세율", "쿼터", "환급",
    "fta", "origin", "rules of origin", "certificate of origin", "원산지", "co 발급",
]

def _row_text(row: pd.Series, cols: list[str] | None = None) -> str:
    if cols is None:
        cols = [
            "Headline", "Summary", "AI Analysis", "Action Plan", "Major Changes", "Issue",
            "Country", "Agency", "KeywordMatches", "affected_products", "hs_hint",
            "tariff_rate_hint", "fta_impact", "export_control_impact", "hs_impact", "tariff_impact",
            "Impact Reason", "Affected Subsidiary",
        ]
    return " ".join(clean(row.get(c)) for c in cols).lower()

def _has_any(text: str, terms: list[str]) -> bool:
    return any(t.lower() in text for t in terms)

def samsung_relevance_score(row: pd.Series) -> int:
    """Samsung product / subsidiary relevance score.

    This prevents generic trade news from being promoted to Top3 only because it
    contains words like export, tariff, or FTA.
    """
    text = _row_text(row)
    score = 0

    if clean(row.get("Samsung Impact")) == "Direct":
        score += 2200
    elif clean(row.get("Samsung Impact")) == "Indirect":
        score += 1000
    elif clean(row.get("Samsung Impact")) == "Watch":
        score += 200

    if _has_any(text, SAMSUNG_RELEVANCE_TERMS["high"]):
        score += 1400
    if _has_any(text, SAMSUNG_RELEVANCE_TERMS["medium"]):
        score += 500

    subs = clean(row.get("Affected Subsidiary"))
    if subs and subs not in {"관련 법인 검토", "SEC/HQ", "HQ", "본사"}:
        score += 400
    products = clean(row.get("affected_products"))
    if products:
        score += 500

    if _has_any(text, LOW_RELEVANCE_TERMS):
        score -= 1600
    if _has_any(text, GENERAL_NEWS_TERMS):
        score -= 800

    return score

def customs_actionability_score(row: pd.Series) -> int:
    """Score whether the item requires customs/trade compliance action."""
    text = _row_text(row)
    issue = clean(row.get("Issue"))
    score = 0

    if issue in {"관세정책", "AD/CVD", "반덤핑/상계관세", "CBAM", "수출통제", "FTA/원산지", "통관", "통관/세관", "HS/품목분류"}:
        score += 700
    if clean(row.get("Content Type")) == "Regulation":
        score += 500
    if normalize_risk(row.get("Risk")) == "상":
        score += 500
    elif normalize_risk(row.get("Risk")) == "중":
        score += 200

    if _has_any(text, CRITICAL_CUSTOMS_TERMS):
        score += 900
    for col in ["hs_hint", "tariff_rate_hint", "effective_date_hint", "change_detail_hint"]:
        if non_empty_hint(row.get(col)):
            score += 250

    # Actionable only when there is a plausible internal task.
    if any(k in text for k in ["시행", "적용", "신고", "세율", "hs", "관세율", "원산지", "co ", "환급", "증빙", "허가", "license"]):
        score += 300

    if _has_any(text, LOW_RELEVANCE_TERMS):
        score -= 1200

    return score

def reference_reason(row: pd.Series) -> str:
    text = _row_text(row)
    if _has_any(text, LOW_RELEVANCE_TERMS):
        return "삼성전자 주요 제품·부품·원재료와 직접 관련성이 낮은 일반 품목/농산물성 규제입니다."
    if samsung_relevance_score(row) < 200 and customs_actionability_score(row) < 700:
        return "관세업무 실행 조치가 필요한 수준의 HS·세율·원산지·신고절차 변경이 확인되지 않았습니다."
    if _has_any(text, GENERAL_NEWS_TERMS):
        return "일반 경제/시장 동향 성격이 강해 임원 보고 Top3보다는 참고 모니터링에 적합합니다."
    return ""

def executive_priority(row: pd.Series) -> str:
    """CORE / WATCH / REFERENCE override for STEP5 mail quality."""
    ref = reference_reason(row)
    if ref:
        return "REFERENCE"
    srel = samsung_relevance_score(row)
    act = customs_actionability_score(row)
    if srel >= 1800 and act >= 1400:
        return "CORE"
    if act >= 1400:
        return "POLICY_WATCH"
    if srel >= 1000 and act >= 900:
        return "USABLE"
    return clean(row.get("Priority Group")) or "WATCH"

def report_score(row: pd.Series) -> float:
    """Override: report score based on relevance/actionability, not raw keyword score."""
    base = safe_num(row.get("Importance Score"))
    impact_weight = {"Direct": 2200, "Indirect": 900, "Watch": 100, "Reference": -800}.get(clean(row.get("Samsung Impact")), 0)
    risk = risk_weight(row.get("Risk"))
    type_weight = 350 if clean(row.get("Content Type")) == "Regulation" else 0
    priority = executive_priority(row)
    pri_weight = {"CORE": 1200, "POLICY_WATCH": 800, "USABLE": 450, "WATCH": 200, "REFERENCE": -1200}.get(priority, 0)
    return base + impact_weight + risk + type_weight + pri_weight + samsung_relevance_score(row) + customs_actionability_score(row)

def top3_deep_score(row: pd.Series) -> float:
    """Override: choose Top3 only when Samsung relevance + customs actionability are high."""
    score = report_score(row)
    text = _row_text(row)
    priority = executive_priority(row)

    if priority == "REFERENCE":
        score -= 10000
    if clean(row.get("Content Type")) == "Regulation":
        score += 300
    if normalize_risk(row.get("Risk")) == "상":
        score += 300

    # Strong boost for issues that usually require HQ action.
    if clean(row.get("Issue")) in {"AD/CVD", "반덤핑/상계관세", "수출통제", "CBAM", "관세정책", "HS/품목분류", "통관"}:
        score += 700

    # Explicitly avoid "export" only items without Samsung/customs action.
    if "export" in text and not _has_any(text, CRITICAL_CUSTOMS_TERMS) and samsung_relevance_score(row) < 800:
        score -= 2500

    return score

def choose_top3(rows: pd.DataFrame) -> pd.DataFrame:
    """Override: exclude REFERENCE/low-actionability from Top3 and keep issue diversity."""
    pool = rows.copy()
    pool["Executive Priority"] = pool.apply(executive_priority, axis=1)
    pool["_top3_score"] = pool.apply(top3_deep_score, axis=1)

    # Top3 후보 필터: Reference 제외 + 업무 실행성 최소값
    candidate = pool[
        (pool["Executive Priority"].ne("REFERENCE")) &
        (pool.apply(customs_actionability_score, axis=1) >= 700)
    ].copy()

    if candidate.empty:
        candidate = pool[pool["Executive Priority"].ne("REFERENCE")].copy()
    if candidate.empty:
        candidate = pool.copy()

    candidate = candidate.sort_values(["_top3_score", "_sort_date"], ascending=[False, False])
    selected = []
    used_issues = set()
    for _, row in candidate.iterrows():
        issue = clean(row.get("Issue"))
        if issue in used_issues and len(selected) < 3:
            continue
        selected.append(row)
        used_issues.add(issue)
        if len(selected) == 3:
            break

    if len(selected) < 3:
        for _, row in candidate.iterrows():
            if any(clean(row.get("Headline")) == clean(x.get("Headline")) for x in selected):
                continue
            selected.append(row)
            if len(selected) == 3:
                break

    out = pd.DataFrame(selected).reset_index(drop=True)
    if not out.empty:
        out["No"] = range(1, len(out) + 1)
    return out

def prepare_rows(rows: pd.DataFrame) -> pd.DataFrame:
    """Override: rewrite report fields and demote low relevance rows at STEP5."""
    rows = rows.copy()
    rows["Issue"] = rows.apply(issue_for, axis=1)
    rows = dedup_report_rows(rows)
    rows["Mail Group"] = rows["Content Type"].map({"Regulation": GROUP_REGULATION}).fillna(GROUP_NEWS)
    rows["Executive Priority"] = rows.apply(executive_priority, axis=1)

    # Demote low relevance items so they remain in report table but do not become Top3.
    rows.loc[rows["Executive Priority"].eq("REFERENCE"), "Priority Group"] = "REFERENCE"
    rows.loc[rows["Executive Priority"].eq("REFERENCE"), "Samsung Impact"] = "Reference"

    rows["Major Changes"] = rows.apply(major_changes, axis=1)
    rows["Summary"] = rows.apply(report_summary, axis=1)
    rows["AI Analysis"] = rows.apply(report_impact, axis=1)
    rows["Action Plan"] = rows.apply(report_action, axis=1)
    rows["_report_score"] = rows.apply(report_score, axis=1)

    # Final mail ordering should reflect executive relevance.
    rows = rows.sort_values(["_report_score", "_sort_date"], ascending=[False, False]).reset_index(drop=True)
    rows["No"] = range(1, len(rows) + 1)
    return rows

def _issue_summary_detail(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    country = clean(row.get("Country")) or "확인 필요"
    date = clean(row.get("Date")) or "확인 필요"
    hs = non_empty_hint(row.get("hs_hint")) or "원문/법인 품목 기준 확인 필요"
    rate = non_empty_hint(row.get("tariff_rate_hint")) or "해당 시 별도 산출 필요"
    change = short_text(row.get("Major Changes"), "주요 변경내역 확인 필요", 420)
    return (
        f"• 이슈구분: {issue}\n"
        f"• 대상국가: {country}\n"
        f"• 게시/시행일: {date}\n"
        f"• 대상 HS/품목: {hs}\n"
        f"• 세율/쿼터/허가 변화: {rate}\n"
        f"• 핵심내용: {change}"
    )

def _samsung_impact_detail(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    subs = clean(row.get("Affected Subsidiary")) or "SEC/HQ 및 관련 해외법인"
    products = non_empty_hint(row.get("affected_products")) or "법인별 수입·수출 실적 기준 매핑 필요"
    impact = clean(row.get("Samsung Impact")) or "Watch"

    if impact == "Reference":
        ref = reference_reason(row)
        return (
            f"• 영향등급: Reference\n"
            f"• 판단사유: {ref or '삼성전자 직접 영향은 낮고 정책 방향성 모니터링 가치 중심입니다.'}\n"
            f"• 영향법인: 즉시 특정 불필요\n"
            f"• 영향품목: {products}"
        )

    base = [
        f"• 영향등급: {impact}",
        f"• 영향법인 후보: {subs}",
        f"• 영향제품 후보: {products}",
    ]

    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        base.append("• 영향업무: 공급국 변경, 원산지 증빙, 가격자료 방어, AD/CVD 추가관세 비용 산출")
        base.append("• 리스크: 대상 HS를 사용하는 원재료·부품 수입 시 추가 관세 및 사후심사 대응 부담 발생 가능")
    elif issue == "CBAM":
        base.append("• 영향업무: EU향 품목의 배출량 자료 확보, CBAM 신고, 인증서 비용 반영")
        base.append("• 리스크: 공급사 배출량 자료 미확보 시 신고 오류 또는 비용 추정 누락 가능")
    elif issue == "수출통제":
        base.append("• 영향업무: ECCN/전략물자 분류, 최종사용자 확인, 우회수출 스크리닝")
        base.append("• 리스크: 허가 필요 품목을 무허가 수출하거나 제재 거래처와 거래할 가능성")
    elif issue == "FTA/원산지":
        base.append("• 영향업무: BOM 원산지 판정, CO 발급요건, 직접운송, 누적기준, FTA Master 정합성")
        base.append("• 리스크: 원산지 기준 미충족 상태에서 특혜세율 적용 또는 CO 발급 오류 가능")
    elif issue in {"통관", "통관/세관"}:
        base.append("• 영향업무: 수입신고, 보세운송/보세공장, 관세사 제출자료, 통관 SOP")
        base.append("• 리스크: 신고 지연, 자동수리 조건 오류, 세관 제출자료 누락 가능")
    elif issue == "HS/품목분류":
        base.append("• 영향업무: HS Master, 품목 설명, 관세율, FTA 판정 기준")
        base.append("• 리스크: 법인별 HS 불일치 및 관세율 오적용 가능")
    else:
        base.append("• 영향업무: 관련 국가·품목 기준 통상 리스크 모니터링")
        base.append("• 리스크: 현재 직접 영향은 제한적이나 정책 확산 여부 확인 필요")

    return "\n".join(base)

def _customs_impact_detail(row: pd.Series) -> str:
    issue = clean(row.get("Issue"))
    lines = []
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        lines = [
            "• 관세비용: 대상 HS 수입금액 × 추가관세율로 잠재 비용 산출 필요",
            "• 신고리스크: 공급국·원산지·제조자 기준이 불명확하면 AD/CVD 회피 의심 가능",
            "• 증빙자료: 구매계약서, 원산지증명, 가격결정자료, 공급자 진술서 보관 필요",
        ]
    elif issue == "CBAM":
        lines = [
            "• 관세/준조세 비용: CBAM 인증서 가격 및 내재배출량 기준 비용 반영 필요",
            "• 신고리스크: EU 수입자 신고자료와 공급사 배출량 자료 불일치 가능",
            "• 증빙자료: 배출량 산정서, 공급사 확인서, 품목별 CN/HS 매핑 필요",
        ]
    elif issue == "수출통제":
        lines = [
            "• 수출허가: ECCN/전략물자 해당 여부 및 최종사용자 확인 필요",
            "• 거래심사: 제재국·제재자·우회수출 경로 스크리닝 필요",
            "• 시스템: Item Master에 전략물자/ECCN/허가필요 여부 필드 반영 검토",
        ]
    elif issue == "FTA/원산지":
        lines = [
            "• FTA 비용: 특혜세율 적용 가능 여부 및 미적용 시 관세비용 차이 산출 필요",
            "• 원산지 리스크: BOM, Vendor 원산지확인서, HS 기준 불일치 가능",
            "• 시스템: FTA Master·HS Master·Item Master 정합성 점검 필요",
        ]
    elif issue in {"통관", "통관/세관"}:
        lines = [
            "• 신고절차: 관세사 신고 양식, 제출자료, 자동수리 조건 변경 여부 확인 필요",
            "• 운영리스크: 보세·수입신고 오류 또는 지연 시 비용/가산세 발생 가능",
            "• 시스템: 통관 체크리스트와 법인 SOP 업데이트 필요",
        ]
    elif issue == "HS/품목분류":
        lines = [
            "• HS 리스크: 동일 품목에 대한 법인·관세사별 HS 불일치 가능",
            "• 비용영향: HS 변경 시 기본세율, FTA 세율, AD/CVD 적용 여부 재산정 필요",
            "• 시스템: HS Master 변경 승인 Workflow 필요",
        ]
    else:
        lines = [
            "• 직접 관세비용 영향은 현재 낮음",
            "• 정책 방향성 모니터링 후 유사 규제가 전자부품·전략물자로 확대되는지 확인 필요",
        ]
    return "\n".join(lines)

def _action_detail(row: pd.Series) -> str:
    impact = clean(row.get("Samsung Impact"))
    issue = clean(row.get("Issue"))
    subs = clean(row.get("Affected Subsidiary")) or "SEC/HQ"

    if impact == "Reference":
        return (
            "• 즉시조치: 불필요\n"
            "• 모니터링: 동일 국가에서 전자부품·전략물자·관세율 관련 후속 공지가 나오는지 확인\n"
            "• GTI 처리: 본문 Top3 제외, Reference 뉴스로 보관\n"
            "• Owner: GTI 운영자"
        )

    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return (
            f"• 즉시조치: {subs} 대상 HS·공급국·벤더 매핑\n"
            "• 1주 내: 최근 12개월 수입실적 기준 잠재 AD/CVD 비용 산출\n"
            "• 1개월 내: 원산지/가격자료 방어 파일 구축 및 관세사 신고 기준 공유\n"
            "• Owner: HQ Customs + 구매 + 해당 법인 통관담당"
        )
    if issue == "CBAM":
        return (
            f"• 즉시조치: {subs} EU향 대상품목 및 공급사 배출량 자료 보유 여부 확인\n"
            "• 1주 내: CBAM 신고 대상 CN/HS와 공급사별 배출량 Gap List 작성\n"
            "• 1개월 내: 인증서 비용 반영 로직 및 ESG/통관 공동관리 체계 수립\n"
            "• Owner: HQ Customs + ESG + EU 판매법인"
        )
    if issue == "수출통제":
        return (
            f"• 즉시조치: {subs} 대상 제품의 ECCN/전략물자 분류 확인\n"
            "• 1주 내: 거래처·최종사용자·목적지 스크리닝 결과 재점검\n"
            "• 1개월 내: Item Master에 Export Control Flag 반영\n"
            "• Owner: HQ Export Control + 사업부 + 해외법인"
        )
    if issue == "FTA/원산지":
        return (
            f"• 즉시조치: {subs} 대상 품목의 FTA 적용 여부와 CO 발급/수취 현황 확인\n"
            "• 1주 내: BOM 원산지, Vendor 원산지확인서, HS 기준 일치 여부 점검\n"
            "• 1개월 내: FTA Master·HS Master·Item Master 업데이트\n"
            "• Owner: HQ Customs/FTA + 법인 구매/물류"
        )
    if issue in {"통관", "통관/세관"}:
        return (
            f"• 즉시조치: {subs} 관세사에 신고절차 변경 여부 확인\n"
            "• 1주 내: 통관 SOP, 보세운송/보세공장 체크리스트, 제출자료 양식 개정\n"
            "• 1개월 내: ERP/ONE-Origin 반영 필요 필드 정의\n"
            "• Owner: HQ Customs + 법인 통관담당 + 관세사"
        )
    if issue == "HS/품목분류":
        return (
            f"• 즉시조치: {subs} 품목별 HS Master와 신고 HS 비교\n"
            "• 1주 내: 불일치 품목 Root Cause 분석 및 변경 승인자료 확보\n"
            "• 1개월 내: HS 변경 Workflow 및 관세율 영향표 반영\n"
            "• Owner: HQ Customs + 법인 Master Data 담당"
        )
    return (
        f"• 즉시조치: {subs} 적용 가능성 확인\n"
        "• 1주 내: 대상 국가·품목·HS·법인 매핑\n"
        "• 1개월 내: 후속 공지 모니터링 및 필요 시 Master 반영\n"
        "• Owner: HQ Customs"
    )

def report_summary(row: pd.Series) -> str:
    return _issue_summary_detail(row).replace("\n", " ")

def report_impact(row: pd.Series) -> str:
    return _samsung_impact_detail(row).replace("\n", " ")

def report_action(row: pd.Series) -> str:
    return _action_detail(row).replace("\n", " ")

def top3_summary_sentence(row: pd.Series) -> str:
    title = clean(row.get("Headline"))
    issue = clean(row.get("Issue"))
    impact = clean(row.get("Samsung Impact"))
    if impact == "Reference":
        return f"{title}: 삼성전자 직접 영향 낮음. Reference로 관리하고 Top3에서는 제외하는 것이 적절합니다."
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        return f"{title}: 대상 HS·공급국·벤더 기준 AD/CVD 비용과 원산지 방어자료 점검이 필요합니다."
    if issue == "CBAM":
        return f"{title}: EU향 품목의 배출량 자료, CBAM 신고 및 인증서 비용 반영 여부를 확인해야 합니다."
    if issue == "수출통제":
        return f"{title}: ECCN·전략물자 분류와 최종사용자 스크리닝을 우선 점검해야 합니다."
    if issue == "FTA/원산지":
        return f"{title}: CO 발급요건, BOM 원산지, FTA Master 정합성 재검토가 필요합니다."
    if issue in {"관세정책", "통관", "통관/세관", "HS/품목분류"}:
        return f"{title}: 대상 HS, 신고절차, 관세율 및 법인 SOP 반영 여부를 확인해야 합니다."
    return f"{title}: 삼성 관련성과 관세업무 실행 필요성을 확인해야 합니다."

def top3_html(top3: pd.DataFrame) -> str:
    blocks = []
    for idx, row in top3.iterrows():
        blocks.append(f"""
        <div style="margin:14px 0 18px 0;padding:15px;border-left:5px solid #C00000;background:#FFF7F7;">
          <div style="font-size:15px;font-weight:bold;margin-bottom:6px;">Top {idx + 1}. {html_link(row.get('Headline'), row.get('URL'))}</div>
          <div style="font-size:12px;color:#555;margin-bottom:9px;">
            Type: {html.escape(clean(row.get('Content Type')))} | Topic: {html.escape(clean(row.get('Issue')))} |
            Samsung Impact: <b>{html.escape(clean(row.get('Samsung Impact')))}</b> |
            Executive Priority: <b>{html.escape(executive_priority(row))}</b> |
            Subsidiary: {html.escape(clean(row.get('Affected Subsidiary')) or 'SEC/HQ')} |
            Agency: {html.escape(clean(row.get('Agency')))} | Publish Date: {html.escape(clean(row.get('Date')))} |
            Country: {html.escape(clean(row.get('Country')))} |
            Risk: <span style="color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</span> |
            Score: {report_score(row):.0f}
          </div>
          <div style="margin-top:8px;"><b>1) 이슈 요약</b><br><pre style="white-space:pre-wrap;font-family:Arial,'Malgun Gothic',sans-serif;margin:4px 0 0 0;">{html.escape(_issue_summary_detail(row))}</pre></div>
          <div style="margin-top:8px;"><b>2) 삼성전자 영향</b><br><pre style="white-space:pre-wrap;font-family:Arial,'Malgun Gothic',sans-serif;margin:4px 0 0 0;">{html.escape(_samsung_impact_detail(row))}</pre></div>
          <div style="margin-top:8px;"><b>3) 관세업무 영향 / 리스크</b><br><pre style="white-space:pre-wrap;font-family:Arial,'Malgun Gothic',sans-serif;margin:4px 0 0 0;">{html.escape(_customs_impact_detail(row))}</pre></div>
          <div style="margin-top:8px;"><b>4) 대응방안</b><br><pre style="white-space:pre-wrap;font-family:Arial,'Malgun Gothic',sans-serif;margin:4px 0 0 0;">{html.escape(_action_detail(row))}</pre></div>
        </div>
        """)
    return "".join(blocks)

def table_html(title: str, rows: pd.DataFrame, color: str) -> str:
    """Override: add Executive Priority column to make REFERENCE demotion visible."""
    if rows.empty:
        return f"<h3 style='color:{color};'>{html.escape(title)} (0건)</h3>"
    trs = []
    for _, row in rows.iterrows():
        trs.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(str(row.get('No')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Issue')))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html_link(row.get('Headline'), row.get('URL'))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('Major Changes'), '주요 변경내역 확인 필요', 260))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('AI Analysis'), '영향 검토 필요', 300))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html.escape(short_text(row.get('Action Plan'), '담당 부서 확인 필요', 300))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Country')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Samsung Impact')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(executive_priority(row))}</td>
        </tr>
        """)
    return f"""
    <h3 style="margin-top:24px;color:{color};">{html.escape(title)} ({len(rows)}건)</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;table-layout:fixed;">
      <colgroup>
        <col style="width:3%;"><col style="width:7%;"><col style="width:21%;">
        <col style="width:20%;"><col style="width:18%;"><col style="width:18%;">
        <col style="width:5%;"><col style="width:4%;"><col style="width:5%;"><col style="width:7%;">
      </colgroup>
      <thead>
        <tr style="background:{color};color:white;">
          <th style="padding:7px;border:1px solid #ddd;">No</th>
          <th style="padding:7px;border:1px solid #ddd;">Issue</th>
          <th style="padding:7px;border:1px solid #ddd;">Headline</th>
          <th style="padding:7px;border:1px solid #ddd;">주요 변경내역</th>
          <th style="padding:7px;border:1px solid #ddd;">삼성 영향</th>
          <th style="padding:7px;border:1px solid #ddd;">Action</th>
          <th style="padding:7px;border:1px solid #ddd;">Country</th>
          <th style="padding:7px;border:1px solid #ddd;">Risk</th>
          <th style="padding:7px;border:1px solid #ddd;">Impact</th>
          <th style="padding:7px;border:1px solid #ddd;">Priority</th>
        </tr>
      </thead>
      <tbody>{''.join(trs)}</tbody>
    </table>
    """

def overall_html(rows: pd.DataFrame, top3: pd.DataFrame) -> str:
    reg = rows[rows["Content Type"].eq("Regulation")]
    news = rows[rows["Content Type"].eq("News")]
    direct = rows[rows["Samsung Impact"].eq("Direct")]
    indirect = rows[rows["Samsung Impact"].eq("Indirect")]
    watch = rows[rows["Samsung Impact"].eq("Watch")]
    ref = rows[rows["Samsung Impact"].eq("Reference")]
    issues = rows["Issue"].value_counts().head(6)
    issue_txt = ", ".join(f"{k} {v}건" for k, v in issues.items())
    top_lines = "".join(f"<li>{html.escape(top3_summary_sentence(r))}</li>" for _, r in top3.iterrows())

    ref_note = ""
    if len(ref):
        ref_note = f"<div style='margin-top:8px;color:#666;'>Reference {len(ref)}건은 삼성 직접 영향 또는 관세업무 실행성이 낮아 Top3 후보에서 제외했습니다.</div>"

    return f"""
    <div style="padding:15px;background:#F4F6F8;border-left:6px solid #1F4E78;margin-bottom:18px;">
      <div style="font-size:14px;color:#555;margin-bottom:8px;">
        금일 선별 결과: 법규 {len(reg)}건, 주요뉴스 {len(news)}건 | Direct {len(direct)}건, Indirect {len(indirect)}건, Watch {len(watch)}건, Reference {len(ref)}건
      </div>
      <div style="font-size:15px;font-weight:bold;line-height:1.8;margin-bottom:8px;">
        금일 GTI Radar는 {html.escape(issue_txt)} 중심으로 관세·통상 변화가 포착되었습니다.
        Top3는 단순 키워드가 아니라 삼성 관련성, 관세업무 실행성, 시행 긴급성, 비용/리스크 규모를 기준으로 재선정했습니다.
      </div>
      <div style="margin-top:8px;"><b>Top3 요약</b><ol style="margin-top:6px;">{top_lines}</ol></div>
      {ref_note}
    </div>
    """

# ======================================================================
# End of GTI STEP5 Executive Quality Patch v3
# ======================================================================


# ======================================================================
# GTI STEP5 Executive Quality Patch v4
# ----------------------------------------------------------------------
# v3 보완사항
# 1) REFERENCE 판정이 과도하게 적용되어 AD/CVD, CBAM, 통관, 수출통제까지
#    Top3 후보에서 제외되는 문제 수정
# 2) Top3가 1건만 나오는 문제 수정: CORE/POLICY_WATCH/USABLE/WATCH 순서로
#    반드시 3건까지 보충
# 3) 공식 법규, 고위험 관세/통관 이슈는 삼성 직접법인이 특정되지 않아도
#    Top3 후보로 유지
# ======================================================================

ACTIONABLE_ISSUES = {
    "관세정책", "AD/CVD", "반덤핑/상계관세", "CBAM",
    "수출통제", "FTA/원산지", "통관", "통관/세관", "HS/품목분류"
}

REFERENCE_ONLY_ISSUES = {"무역일반", "일반경제", "시장동향", "기타"}

def reference_reason(row: pd.Series) -> str:
    """v4 override: only demote clearly non-actionable items.

    Do not demote AD/CVD, CBAM, customs, export control, HS, or official high-risk
    regulations simply because affected subsidiary is not specified yet.
    """
    text = _row_text(row)
    issue = clean(row.get("Issue"))
    title = clean(row.get("Headline")).lower()
    content_type = clean(row.get("Content Type"))
    risk = normalize_risk(row.get("Risk"))

    # 1) Explicit agriculture/food/general item with no Samsung/customs action.
    if _has_any(text, LOW_RELEVANCE_TERMS):
        if issue not in ACTIONABLE_ISSUES or issue in REFERENCE_ONLY_ISSUES or "wheat" in text or "염소산업" in text:
            return "삼성전자 주요 제품·부품·원재료와 직접 관련성이 낮은 일반 품목/농산물성 규제입니다."

    # 2) General stock/economy news, even if semiconductor appears as macro statistic.
    if _has_any(text, GENERAL_NEWS_TERMS):
        if not any(k in text for k in ["tariff rate", "관세율", "수출통제", "cbam", "ad/cvd", "반덤핑", "상계관세", "hs code"]):
            return "일반 경제/시장 동향 성격이 강해 임원 보고 Top3보다는 참고 모니터링에 적합합니다."

    # 3) Generic trade/export notice without HS/tariff/origin/customs/export-control action.
    if issue in REFERENCE_ONLY_ISSUES or issue == "무역일반":
        if not _has_any(text, CRITICAL_CUSTOMS_TERMS):
            return "관세업무 실행 조치가 필요한 수준의 HS·세율·원산지·신고절차 변경이 확인되지 않았습니다."

    # 4) Weak non-regulation item with no actionable issue.
    if content_type != "Regulation" and issue not in ACTIONABLE_ISSUES:
        if samsung_relevance_score(row) < 300 and customs_actionability_score(row) < 600:
            return "삼성 관련성과 관세업무 실행성이 낮아 Reference 관리가 적절합니다."

    # Keep official regulations and actionable issues as candidates.
    if content_type == "Regulation" and risk in {"상", "중"}:
        return ""
    if issue in ACTIONABLE_ISSUES:
        return ""

    return ""

def executive_priority(row: pd.Series) -> str:
    """v4 override: preserve actionable customs issues as candidates."""
    ref = reference_reason(row)
    if ref:
        return "REFERENCE"

    srel = samsung_relevance_score(row)
    act = customs_actionability_score(row)
    issue = clean(row.get("Issue"))
    content_type = clean(row.get("Content Type"))
    risk = normalize_risk(row.get("Risk"))

    if clean(row.get("Samsung Impact")) == "Direct":
        return "CORE"
    if clean(row.get("Samsung Impact")) == "Indirect" and issue in ACTIONABLE_ISSUES:
        return "CORE"
    if issue in {"AD/CVD", "반덤핑/상계관세", "수출통제", "CBAM", "HS/품목분류"}:
        return "POLICY_WATCH" if clean(row.get("Samsung Impact")) != "Indirect" else "CORE"
    if issue in {"관세정책", "통관", "통관/세관"}:
        return "POLICY_WATCH"
    if content_type == "Regulation" and risk == "상":
        return "POLICY_WATCH"
    if issue == "FTA/원산지" and (srel >= 900 or "battery" in _row_text(row) or risk == "상"):
        return "USABLE"
    if act >= 1100:
        return "USABLE"
    return clean(row.get("Priority Group")) or "WATCH"

def report_score(row: pd.Series) -> float:
    """v4 override: balanced scoring."""
    base = safe_num(row.get("Importance Score"))
    issue = clean(row.get("Issue"))
    priority = executive_priority(row)
    content_type = clean(row.get("Content Type"))

    issue_weight = {
        "AD/CVD": 1500,
        "반덤핑/상계관세": 1500,
        "수출통제": 1450,
        "CBAM": 1300,
        "관세정책": 1150,
        "통관": 1100,
        "통관/세관": 1100,
        "HS/품목분류": 1100,
        "FTA/원산지": 800,
        "무역일반": 100,
    }.get(issue, 250)

    pri_weight = {
        "CORE": 1800,
        "POLICY_WATCH": 1300,
        "USABLE": 800,
        "WATCH": 300,
        "REFERENCE": -2500,
    }.get(priority, 0)

    impact_weight = {
        "Direct": 2400,
        "Indirect": 1200,
        "Watch": 300,
        "Reference": -1000,
    }.get(clean(row.get("Samsung Impact")), 0)

    type_weight = 550 if content_type == "Regulation" else 0
    risk_w = risk_weight(row.get("Risk"))

    return (
        base + issue_weight + pri_weight + impact_weight + type_weight + risk_w
        + max(samsung_relevance_score(row), -1000)
        + max(customs_actionability_score(row), -500)
    )

def top3_deep_score(row: pd.Series) -> float:
    """v4 override: Top3 ranking with explicit non-reference and issue priority."""
    priority = executive_priority(row)
    if priority == "REFERENCE":
        return -999999

    issue = clean(row.get("Issue"))
    score = report_score(row)

    # Favor concrete cost/compliance topics.
    if issue in {"AD/CVD", "반덤핑/상계관세"}:
        score += 1800
    elif issue == "수출통제":
        score += 1600
    elif issue == "CBAM":
        score += 1400
    elif issue in {"관세정책", "통관", "통관/세관"}:
        score += 1200
    elif issue == "HS/품목분류":
        score += 1000
    elif issue == "FTA/원산지":
        score += 600

    if clean(row.get("Content Type")) == "Regulation":
        score += 500
    if normalize_risk(row.get("Risk")) == "상":
        score += 500

    return score

def choose_top3(rows: pd.DataFrame) -> pd.DataFrame:
    """v4 override: always select up to 3 non-reference rows if available."""
    pool = rows.copy()
    pool["Executive Priority"] = pool.apply(executive_priority, axis=1)
    pool["_top3_score"] = pool.apply(top3_deep_score, axis=1)

    non_ref = pool[pool["Executive Priority"].ne("REFERENCE")].copy()
    if non_ref.empty:
        non_ref = pool.copy()

    non_ref = non_ref.sort_values(["_top3_score", "_sort_date"], ascending=[False, False])

    selected = []
    used_issues = set()

    # 1st pass: issue diversity
    for _, row in non_ref.iterrows():
        issue = clean(row.get("Issue"))
        if issue in used_issues:
            continue
        selected.append(row)
        used_issues.add(issue)
        if len(selected) == 3:
            break

    # 2nd pass: fill remaining even if same issue
    if len(selected) < 3:
        for _, row in non_ref.iterrows():
            title = clean(row.get("Headline"))
            if any(clean(x.get("Headline")) == title for x in selected):
                continue
            selected.append(row)
            if len(selected) == 3:
                break

    out = pd.DataFrame(selected).reset_index(drop=True)
    if not out.empty:
        out["No"] = range(1, len(out) + 1)
    return out

def prepare_rows(rows: pd.DataFrame) -> pd.DataFrame:
    """v4 override: demote only true references, keep actionable issues."""
    rows = rows.copy()
    rows["Issue"] = rows.apply(issue_for, axis=1)
    rows = dedup_report_rows(rows)
    rows["Mail Group"] = rows["Content Type"].map({"Regulation": GROUP_REGULATION}).fillna(GROUP_NEWS)
    rows["Executive Priority"] = rows.apply(executive_priority, axis=1)

    ref_mask = rows["Executive Priority"].eq("REFERENCE")
    rows.loc[ref_mask, "Priority Group"] = "REFERENCE"
    rows.loc[ref_mask, "Samsung Impact"] = "Reference"

    # Preserve Priority Group for actionable issues.
    rows.loc[~ref_mask, "Priority Group"] = rows.loc[~ref_mask, "Executive Priority"]

    rows["Major Changes"] = rows.apply(major_changes, axis=1)
    rows["Summary"] = rows.apply(report_summary, axis=1)
    rows["AI Analysis"] = rows.apply(report_impact, axis=1)
    rows["Action Plan"] = rows.apply(report_action, axis=1)
    rows["_report_score"] = rows.apply(report_score, axis=1)

    rows = rows.sort_values(["_report_score", "_sort_date"], ascending=[False, False]).reset_index(drop=True)
    rows["No"] = range(1, len(rows) + 1)
    return rows

# ======================================================================
# End of GTI STEP5 Executive Quality Patch v4
# ======================================================================


# ======================================================================
# GTI STEP5 Executive Quality Patch v5
# ----------------------------------------------------------------------
# v4 보완사항
# 1) 2.Top3 / 3.Regulation / 4.주요뉴스 모두에 게시물 요약 반영
# 2) "주요 변경내역" = 기존 변경내역 + 게시물 전체요약 2~3줄
# 3) STEP4 원본 Summary/AI Analysis를 보존하여 STEP5에서 덮어쓰기 전에 활용
# ======================================================================

def _source_summary_text(row: pd.Series, limit: int = 520) -> str:
    """Return original post/article summary from STEP4 before STEP5 rewrite."""
    candidates = [
        clean(row.get("Original Summary")),
        clean(row.get("Original AI Analysis")),
        clean(row.get("Impact Reason")),
        clean(row.get("Summary")),
    ]
    for text in candidates:
        if text and text not in {"본문에서 확인 불가", "nan", "None"}:
            # Remove repetitive generic STEP5-style sentences if already rewritten.
            bad = [
                "공식 법규/공지입니다. 핵심은",
                "에서 포착된",
                "삼성전자 관련 법인",
            ]
            if any(b in text for b in bad) and len(text) < 180:
                continue
            text = re.sub(r"\s+", " ", text).strip()
            return text[:limit] + ("..." if len(text) > limit else "")
    return "원문 요약 정보가 부족합니다. 원문 링크 기준으로 세부 내용 확인이 필요합니다."

def _two_three_line_summary(row: pd.Series) -> str:
    """Create 2~3 line bulletin summary for mail display."""
    src = _source_summary_text(row, 680)
    issue = clean(row.get("Issue"))
    country = clean(row.get("Country")) or "관련국"
    date = clean(row.get("Date")) or "게시일 확인 필요"
    headline = clean(row.get("Headline"))

    # If source summary is weak, create a structured fallback from available fields.
    if src.startswith("원문 요약 정보가 부족"):
        src = (
            f"{headline} 관련 {issue} 이슈입니다. "
            f"대상 국가는 {country}이며 게시/확인일은 {date}입니다. "
            "세부 대상 품목, HS, 세율, 시행일은 원문 및 법인 실적 기준으로 추가 확인이 필요합니다."
        )

    sentences = re.split(r"(?<=[.!?。？！])\s+|(?<=다\.)\s+|(?<=니다\.)\s+", src)
    sentences = [s.strip(" -•\n\t") for s in sentences if s.strip()]
    if len(sentences) >= 2:
        return "\n".join(f"• {s}" for s in sentences[:3])
    # Short text fallback: split by length.
    if len(src) > 180:
        return f"• {src[:180].strip()}\n• {src[180:360].strip()}"
    return f"• {src}"

def major_changes(row: pd.Series) -> str:
    """v5 override: current change detail + article/post full summary 2~3 lines."""
    issue = clean(row.get("Issue"))
    headline = clean(row.get("Headline"))
    title_l = headline.lower()

    if "보세창고" in headline and "특허" in headline:
        current = (
            "개정 사유: 자가용보세창고 특허요건 완화 및 불명확한 규정 보완 필요. "
            "주요 개정 내용: 자가용보세창고 반입 대상에 국제무역선·기 적재 자가화물 외 수리용 예비부분품 및 부속품 장치를 허용하고, "
            "관세법 제178조상 물품반입 정지기간을 오해 없이 적용할 수 있도록 규정을 명확화하는 내용입니다."
        )
    elif "환전영업자" in headline and "관리" in headline:
        current = (
            "주요 내용: 환전영업자의 등록·관리, 보고·자료제출, 영업장 운영 및 관세청 관리 기준과 관련된 고시입니다. "
            "해외출장·주재원·외환거래 지원 프로세스와 연결될 수 있어 실제 법인 업무 해당 여부 확인이 필요합니다."
        )
    elif "cbam" in title_l and "certificate price" in title_l:
        current = (
            "주요 내용: EU CBAM 인증서 가격이 공표되었거나 공표 일정이 확정된 사안입니다. "
            "EU 수입품의 내재배출량 신고, 인증서 구매 비용, 공급사 배출량 자료 확보 체계에 영향을 줄 수 있습니다."
        )
    elif "customs enforcement" in title_l and "executive order" in title_l:
        current = (
            "주요 내용: 미국 세관 집행 강화 행정명령 관련 사안입니다. "
            "수입신고 정확성, 저가신고·우회수입·전자상거래 물품 관리 및 CBP 심사 강화 가능성을 확인해야 합니다."
        )
    elif "수입신고" in headline and "가산세" in headline:
        current = (
            "주요 내용: 수입신고 지연 가산세 부과 대상이 되는 매점매석 금지 품목의 적용기간 연장 공고입니다. "
            "해당 품목 수입 시 신고 지연, 재고 운영, 통관 일정 관리 기준을 확인해야 합니다."
        )
    else:
        parts = [
            hint_line("시행/적용일", row.get("effective_date_hint")),
            hint_line("변경 내용", row.get("change_detail_hint")),
            hint_line("대상 HS", row.get("hs_hint")),
            hint_line("관세율/쿼터", row.get("tariff_rate_hint")),
            hint_line("키워드", row.get("KeywordMatches")),
        ]
        if any(parts):
            current = compact_parts(parts, "")
        elif issue == "관세정책":
            current = "관세율, 쿼터, 면세/환급 또는 Section 301/232 등 관세 비용에 영향을 줄 수 있는 정책 변화입니다."
        elif issue in {"AD/CVD", "반덤핑/상계관세"}:
            current = "반덤핑 또는 상계관세 조사·판정·연장 가능성이 있는 사안입니다. 공급국, 대상 품목, 조사 기간과 관세율 확인이 필요합니다."
        elif issue == "CBAM":
            current = "CBAM 신고, 인증서 가격, 배출량 자료 또는 EU 수입통관 절차와 연결되는 탄소국경조정 변화입니다."
        elif issue == "FTA/원산지":
            current = "FTA/CEPA 협정, 원산지 기준, CO 발급 또는 특혜관세 적용 가능성에 영향을 주는 변화입니다."
        elif issue == "수출통제":
            current = "Entity List, ECCN, UFLPA, forced labor 또는 전략물자·제재 스크리닝 관련 변화입니다."
        elif issue == "통관":
            current = "보세, 통관, 신고, 세관 심사 또는 행정절차 기준에 영향을 줄 수 있는 공식 공지입니다."
        elif issue == "HS/품목분류":
            current = "HS 분류 기준 또는 품목 해석이 달라질 수 있어 품목 마스터와 신고 기준 점검이 필요한 사안입니다."
        else:
            current = f"{headline} 관련 관세·통상 모니터링 사안입니다."

    post_summary = _two_three_line_summary(row)
    return f"{current}\n\n[게시물 요약]\n{post_summary}"

def _issue_summary_detail(row: pd.Series) -> str:
    """v5 override: Top3 issue summary includes post summary."""
    issue = clean(row.get("Issue"))
    country = clean(row.get("Country")) or "확인 필요"
    date = clean(row.get("Date")) or "확인 필요"
    hs = non_empty_hint(row.get("hs_hint")) or "원문/법인 품목 기준 확인 필요"
    rate = non_empty_hint(row.get("tariff_rate_hint")) or "해당 시 별도 산출 필요"
    change = clean(row.get("Major Changes")) or major_changes(row)
    return (
        f"• 이슈구분: {issue}\n"
        f"• 대상국가: {country}\n"
        f"• 게시/시행일: {date}\n"
        f"• 대상 HS/품목: {hs}\n"
        f"• 세율/쿼터/허가 변화: {rate}\n"
        f"• 주요 변경내역 및 게시물 요약:\n{change}"
    )

def prepare_rows(rows: pd.DataFrame) -> pd.DataFrame:
    """v5 override: preserve original summaries, then generate Major Changes with post summary."""
    rows = rows.copy()
    rows["Issue"] = rows.apply(issue_for, axis=1)
    rows = dedup_report_rows(rows)
    rows["Mail Group"] = rows["Content Type"].map({"Regulation": GROUP_REGULATION}).fillna(GROUP_NEWS)

    # Preserve STEP4 source texts before STEP5 overwrites them.
    if "Original Summary" not in rows.columns:
        rows["Original Summary"] = rows.get("Summary", "")
    if "Original AI Analysis" not in rows.columns:
        rows["Original AI Analysis"] = rows.get("AI Analysis", "")
    if "Original Action Plan" not in rows.columns:
        rows["Original Action Plan"] = rows.get("Action Plan", "")

    rows["Executive Priority"] = rows.apply(executive_priority, axis=1)

    ref_mask = rows["Executive Priority"].eq("REFERENCE")
    rows.loc[ref_mask, "Priority Group"] = "REFERENCE"
    rows.loc[ref_mask, "Samsung Impact"] = "Reference"
    rows.loc[~ref_mask, "Priority Group"] = rows.loc[~ref_mask, "Executive Priority"]

    rows["Major Changes"] = rows.apply(major_changes, axis=1)
    rows["Summary"] = rows.apply(report_summary, axis=1)
    rows["AI Analysis"] = rows.apply(report_impact, axis=1)
    rows["Action Plan"] = rows.apply(report_action, axis=1)
    rows["_report_score"] = rows.apply(report_score, axis=1)

    rows = rows.sort_values(["_report_score", "_sort_date"], ascending=[False, False]).reset_index(drop=True)
    rows["No"] = range(1, len(rows) + 1)
    return rows

def table_html(title: str, rows: pd.DataFrame, color: str) -> str:
    """v5 override: widen Major Changes and show line breaks for post summary."""
    if rows.empty:
        return f"<h3 style='color:{color};'>{html.escape(title)} (0건)</h3>"
    trs = []
    for _, row in rows.iterrows():
        major = html.escape(short_text(row.get('Major Changes'), '주요 변경내역 및 게시물 요약 확인 필요', 620)).replace("\n", "<br>")
        impact = html.escape(short_text(row.get('AI Analysis'), '영향 검토 필요', 360)).replace("\n", "<br>")
        action = html.escape(short_text(row.get('Action Plan'), '담당 부서 확인 필요', 360)).replace("\n", "<br>")
        trs.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(str(row.get('No')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Issue')))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html_link(row.get('Headline'), row.get('URL'))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{major}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{impact}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{action}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Country')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Samsung Impact')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(executive_priority(row))}</td>
        </tr>
        """)
    return f"""
    <h3 style="margin-top:24px;color:{color};">{html.escape(title)} ({len(rows)}건)</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;table-layout:fixed;">
      <colgroup>
        <col style="width:3%;"><col style="width:7%;"><col style="width:19%;">
        <col style="width:25%;"><col style="width:16%;"><col style="width:16%;">
        <col style="width:5%;"><col style="width:4%;"><col style="width:5%;"><col style="width:7%;">
      </colgroup>
      <thead>
        <tr style="background:{color};color:white;">
          <th style="padding:7px;border:1px solid #ddd;">No</th>
          <th style="padding:7px;border:1px solid #ddd;">Issue</th>
          <th style="padding:7px;border:1px solid #ddd;">Headline</th>
          <th style="padding:7px;border:1px solid #ddd;">주요 변경내역 + 게시물 요약</th>
          <th style="padding:7px;border:1px solid #ddd;">삼성 영향</th>
          <th style="padding:7px;border:1px solid #ddd;">Action</th>
          <th style="padding:7px;border:1px solid #ddd;">Country</th>
          <th style="padding:7px;border:1px solid #ddd;">Risk</th>
          <th style="padding:7px;border:1px solid #ddd;">Impact</th>
          <th style="padding:7px;border:1px solid #ddd;">Priority</th>
        </tr>
      </thead>
      <tbody>{''.join(trs)}</tbody>
    </table>
    """

# ======================================================================
# End of GTI STEP5 Executive Quality Patch v5
# ======================================================================


# ======================================================================
# GTI STEP5 Final Mail Patch - 2026-06-14
# ----------------------------------------------------------------------
# 1) Top3는 같은 이슈를 중복 선정하지 않음
# 2) 뉴스는 30~50건 범위 유지: Step4 summary가 적으면 audit candidates에서 보강
# 3) 삼성전자 본사 관세담당자 관점으로 영향도 재분류
# 4) Mail 표기: Priority -> Publish Date, 주요 변경내역+게시물 요약 -> Summary
# 5) 총평은 짧고 실행형 문장으로 구성
# ======================================================================

NEWS_MIN_REPORT_ROWS = int(os.getenv("GTI_NEWS_MIN_REPORT_ROWS", "30"))
NEWS_MAX_REPORT_ROWS = int(os.getenv("GTI_NEWS_MAX_REPORT_ROWS", "50"))
NEWS_AUDIT_INPUT_FILE = Path(os.getenv("GTI_NEWS_AUDIT_INPUT_FILE", r"C:\Temp\4-2.news_ai_audit_candidates.xlsx"))


def _text_blob(row: pd.Series) -> str:
    return " ".join(
        clean(row.get(c))
        for c in [
            "Headline", "Major Changes", "Summary", "Original Summary", "AI Analysis",
            "Original AI Analysis", "Action Plan", "Impact Reason", "Issue",
            "Country", "Agency", "KeywordMatches"
        ]
    )


def _is_bad_report_url(url: str) -> bool:
    low = clean(url).lower()
    if not low:
        return True
    if "news.google.com" in low:
        return True
    return low in {
        "https://news.google.com", "https://news.google.com/",
        "https://google.com", "https://www.google.com", "https://www.google.com/",
    }


def _canonical_issue_key(row: pd.Series) -> str:
    text = _text_blob(row).lower()
    issue = clean(row.get("Issue")) or issue_for(row)

    if any(k in text for k in ["anti-dumping", "antidumping", "countervailing", "ad/cvd", "덤핑", "반덤핑", "상계관세"]):
        if any(k in text for k in ["zinc", "아연", "galvanized", "도금", "cold-rolled", "냉간압연", "steel", "철강"]):
            return "AD_CVD_STEEL_ZINC"
        return "AD_CVD_GENERAL"
    if "cbam" in text:
        return "CBAM_CERTIFICATE_PRICE" if ("certificate price" in text or "certificate" in text or "인증서" in text) else "CBAM_GENERAL"
    if ("morocco" in text or "모로코" in text) and ("cepa" in text or "fta" in text):
        return "FTA_CEPA_MOROCCO"
    if "rare earth" in text or "희토류" in text:
        return "EXPORT_CONTROL_RARE_EARTH"
    if "forced labor" in text or "uflpa" in text or "강제노동" in text:
        return "FORCED_LABOR_CUSTOMS"
    if "section 301" in text or "section 232" in text or "301조" in text or "232조" in text:
        return "US_SECTION_TARIFF"
    if "보세창고" in text:
        return "CUSTOMS_BONDED_WAREHOUSE"
    if "보세공장" in text:
        return "CUSTOMS_BONDED_FACTORY"
    if "과세가격" in text or "customs valuation" in text:
        return "CUSTOMS_VALUATION"

    title = clean(row.get("Headline")).lower()
    title = re.sub(r"\([^)]*\)|\[[^]]*\]", " ", title)
    title = re.sub(r"[^0-9a-z가-힣]+", " ", title)
    tokens = [t for t in title.split() if len(t) >= 2]
    return f"{issue}:{' '.join(tokens[:7])}"


def _normalized_issue_type(row: pd.Series) -> str:
    issue = clean(row.get("Issue")) or issue_for(row)
    text = _text_blob(row).lower()
    if issue in {"AD/CVD", "반덤핑/상계관세"} or any(k in text for k in ["anti-dumping", "antidumping", "countervailing", "ad/cvd", "반덤핑", "상계관세", "덤핑방지"]):
        return "AD_CVD"
    if issue == "FTA/원산지" or any(k in text for k in ["fta", "cepa", "origin", "원산지"]):
        return "FTA_ORIGIN"
    if issue == "CBAM" or "cbam" in text:
        return "CBAM"
    if issue == "수출통제" or any(k in text for k in ["export control", "entity list", "uflpa", "forced labor", "수출통제", "강제노동"]):
        return "EXPORT_CONTROL"
    if issue in {"통관", "통관/세관"} or any(k in text for k in ["customs", "통관", "보세", "과세가격"]):
        return "CUSTOMS"
    if issue == "관세정책" or any(k in text for k in ["tariff", "관세", "quota", "쿼터", "section 301", "section 232"]):
        return "TARIFF_POLICY"
    if issue == "HS/품목분류" or any(k in text for k in ["hs code", "품목분류"]):
        return "HS_CLASSIFICATION"
    return issue


def _display_issue(row: pd.Series) -> str:
    issue = clean(row.get("Issue")) or issue_for(row)
    text = _text_blob(row).lower()
    if issue in {"반덤핑/상계관세", "AD/CVD"} or any(k in text for k in ["anti-dumping", "antidumping", "countervailing", "덤핑방지", "상계관세"]):
        return "AD/CVD"
    if any(k in text for k in ["보세공장", "보세창고", "보세운송", "반출입신고", "통관", "customs clearance", "과세가격"]):
        return "통관/세관"
    if issue in {"FTA/원산지"} or any(k in text for k in ["fta", "cepa", "origin", "원산지"]):
        return "FTA/원산지"
    return issue


def _mail_news_score(row: pd.Series) -> float:
    text = _text_blob(row).lower()
    score = safe_num(row.get("Importance Score")) + risk_weight(row.get("Risk")) + priority_weight(row.get("Priority Group"))

    high_terms = [
        "anti-dumping", "antidumping", "countervailing", "ad/cvd", "덤핑", "반덤핑", "상계관세",
        "cbam", "section 301", "section 232", "301조", "232조", "tariff", "관세", "quota", "쿼터",
        "export control", "entity list", "forced labor", "uflpa", "수출통제", "강제노동",
        "fta", "cepa", "origin", "원산지", "customs", "통관", "보세", "hs code", "품목분류",
        "rare earth", "희토류", "battery", "배터리", "semiconductor", "반도체", "steel", "철강",
    ]
    for term in high_terms:
        if term in text:
            score += 220

    weak_terms = [
        "주가", "증시", "코스피", "코스닥", "환율", "부동산", "금리", "은행", "채권",
        "혈통관리", "연예", "스포츠", "신간", "서평", "bookreview",
        "세관인", "주무관 선정", "공무원 선정", "표창", "수상",
    ]
    for term in weak_terms:
        if term in text:
            score -= 500
    if "수출 85.9" in text and not any(k in text for k in ["관세", "통관", "fta", "원산지", "수출통제"]):
        score -= 700
    if _is_bad_report_url(row.get("URL")):
        score -= 1500
    if _hard_reference_news(row):
        score -= 3000
    if clean(row.get("Content Type")) == "News":
        score += 50
    return score


def _hard_reference_news(row: pd.Series) -> bool:
    if clean(row.get("Content Type")) != "News":
        return False
    text = _text_blob(row).lower()
    url = clean(row.get("URL")).lower()
    hard_weak = [
        "신간", "서평", "bookreview", "/culture/", "문화", "혈통관리",
        "세관인", "주무관 선정", "공무원 선정", "표창", "수상",
    ]
    explicit_trade = [
        "tariff", "관세", "customs", "통관", "fta", "cepa", "origin", "원산지",
        "cbam", "anti-dumping", "antidumping", "countervailing", "덤핑", "상계관세",
        "export control", "수출통제", "entity list", "forced labor", "uflpa",
    ]
    return any(k in text or k in url for k in hard_weak) and not any(k in text for k in explicit_trade)


def _dedup_by_issue_or_url(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return rows
    rows = rows.copy()
    rows["_mail_url_key"] = rows["URL"].apply(lambda v: clean(v).lower())
    rows["_mail_issue_key"] = rows.apply(_canonical_issue_key, axis=1)
    rows["_mail_score"] = rows.apply(_mail_news_score, axis=1)
    rows = rows.sort_values(["_mail_score", "_sort_date"], ascending=[False, False])
    rows = rows.drop_duplicates(subset=["_mail_url_key"], keep="first")
    rows = rows.drop_duplicates(subset=["_mail_issue_key"], keep="first")
    return rows.drop(columns=["_mail_url_key", "_mail_issue_key", "_mail_score"], errors="ignore").reset_index(drop=True)


def read_step4_results() -> pd.DataFrame:
    frames = []
    if REGULATION_INPUT_FILE.exists():
        frames.append(normalize_input(pd.read_excel(REGULATION_INPUT_FILE), "Regulation", REGULATION_INPUT_FILE))

    news_frames = []
    if NEWS_INPUT_FILE.exists():
        news_frames.append(normalize_input(pd.read_excel(NEWS_INPUT_FILE), "News", NEWS_INPUT_FILE))
    if NEWS_AUDIT_INPUT_FILE.exists():
        try:
            news_frames.append(normalize_input(pd.read_excel(NEWS_AUDIT_INPUT_FILE), "News", NEWS_AUDIT_INPUT_FILE))
        except Exception as exc:
            print(f"[WARN] news audit top-up skipped: {NEWS_AUDIT_INPUT_FILE} / {exc}")

    if news_frames:
        news = pd.concat(news_frames, ignore_index=True)
        news = news[~news["URL"].apply(_is_bad_report_url)].copy()
        news = _dedup_by_issue_or_url(news)
        max_rows = NEWS_MAX_ROWS if NEWS_MAX_ROWS > 0 else NEWS_MAX_REPORT_ROWS
        max_rows = max(NEWS_MIN_REPORT_ROWS, min(NEWS_MAX_REPORT_ROWS, max_rows))
        frames.append(news.head(max_rows))

    if not frames:
        raise FileNotFoundError(f"STEP4 outputs not found: {REGULATION_INPUT_FILE}, {NEWS_INPUT_FILE}")

    rows = pd.concat(frames, ignore_index=True)
    rows["_dedup_key"] = rows.apply(
        lambda r: clean(r.get("URL")).lower() or (
            clean(r.get("Headline"))[:160] + "|" + clean(r.get("Agency")) + "|" + clean(r.get("Date"))
        ),
        axis=1,
    )
    rows = rows.drop_duplicates(subset=["_dedup_key"], keep="first").drop(columns=["_dedup_key"], errors="ignore")
    rows["_integrated_score"] = rows.apply(
        lambda r: priority_weight(r["Priority Group"])
        + risk_weight(r["Risk"])
        + (180 if r["Content Type"] == "Regulation" else 0)
        + safe_num(r["Importance Score"]),
        axis=1,
    )
    return rows.reset_index(drop=True)


def final_samsung_impact(row: pd.Series) -> str:
    text = _text_blob(row).lower()
    issue = clean(row.get("Issue")) or issue_for(row)

    if _hard_reference_news(row):
        return "Reference"

    direct_terms = [
        "semiconductor", "반도체", "memory", "hbm", "display", "oled", "smartphone",
        "battery", "배터리", "lithium", "nickel", "cobalt", "rare earth", "희토류",
        "steel", "철강", "aluminum", "알루미늄", "copper", "zinc", "아연",
    ]
    actionable_terms = [
        "anti-dumping", "antidumping", "countervailing", "ad/cvd", "덤핑", "반덤핑", "상계관세",
        "cbam", "tariff", "관세", "quota", "쿼터", "section 301", "section 232",
        "export control", "수출통제", "entity list", "forced labor", "uflpa", "강제노동",
        "fta", "cepa", "origin", "원산지", "customs", "통관", "보세", "과세가격", "hs code", "품목분류",
    ]
    weak_terms = ["주가", "증시", "환율", "금리", "은행", "부동산", "스포츠", "연예", "혈통관리"]

    if any(t in text for t in weak_terms) and not any(t in text for t in actionable_terms):
        return "Reference"
    if any(t in text for t in direct_terms) and any(t in text for t in actionable_terms):
        return "Indirect"
    if issue in {"AD/CVD", "CBAM", "FTA/원산지", "수출통제", "관세정책", "통관", "통관/세관", "HS/품목분류"}:
        return "Indirect"
    if any(t in text for t in actionable_terms):
        return "Watch"
    return clean(row.get("Samsung Impact")) or "Watch"


def executive_priority(row: pd.Series) -> str:
    impact = final_samsung_impact(row)
    issue = clean(row.get("Issue")) or issue_for(row)
    text = _text_blob(row).lower()
    if impact == "Reference":
        return "REFERENCE"
    if issue in {"AD/CVD", "CBAM", "수출통제", "관세정책", "통관", "통관/세관"}:
        return "CORE"
    if issue in {"FTA/원산지", "HS/품목분류"}:
        return "POLICY_WATCH"
    if any(k in text for k in ["tariff", "관세", "quota", "쿼터", "customs", "통관", "fta", "cepa"]):
        return "POLICY_WATCH"
    return "WATCH"


def report_score(row: pd.Series) -> float:
    impact_weight = {"Direct": 2500, "Indirect": 1200, "Watch": 200, "Reference": -1500}.get(final_samsung_impact(row), 0)
    priority = executive_priority(row)
    pri_weight = {"CORE": 1400, "POLICY_WATCH": 900, "USABLE": 500, "WATCH": 200, "REFERENCE": -1600}.get(priority, 0)
    issue = clean(row.get("Issue")) or issue_for(row)
    issue_weight = {
        "AD/CVD": 1800, "수출통제": 1600, "CBAM": 1500, "관세정책": 1400,
        "통관": 1200, "통관/세관": 1200, "FTA/원산지": 1000, "HS/품목분류": 900,
    }.get(issue, 250)
    type_weight = 1100 if clean(row.get("Content Type")) == "Regulation" else 0
    text = _text_blob(row).lower()
    if clean(row.get("Content Type")) == "Regulation" and any(k in text for k in ["관세", "통관", "보세", "덤핑", "상계관세", "원산지", "fta", "cepa", "cbam"]):
        type_weight += 900
    return safe_num(row.get("Importance Score")) + risk_weight(row.get("Risk")) + impact_weight + pri_weight + issue_weight + type_weight + _mail_news_score(row) / 5


def choose_top3(rows: pd.DataFrame) -> pd.DataFrame:
    pool = rows.copy()
    pool["Samsung Impact"] = pool.apply(final_samsung_impact, axis=1)
    pool["Executive Priority"] = pool.apply(executive_priority, axis=1)
    pool["_top3_score"] = pool.apply(report_score, axis=1)
    pool["_issue_key"] = pool.apply(_canonical_issue_key, axis=1)
    pool = pool[pool["Executive Priority"].ne("REFERENCE")].copy()
    if pool.empty:
        pool = rows.copy()
        pool["_top3_score"] = pool.apply(report_score, axis=1)
        pool["_issue_key"] = pool.apply(_canonical_issue_key, axis=1)

    pool = pool.sort_values(["_top3_score", "_sort_date"], ascending=[False, False])
    selected, used_keys, used_issue_types = [], set(), set()
    for _, row in pool.iterrows():
        key = clean(row.get("_issue_key"))
        issue = _normalized_issue_type(row)
        if key in used_keys:
            continue
        if issue in used_issue_types and len(selected) < 3:
            continue
        selected.append(row)
        used_keys.add(key)
        used_issue_types.add(issue)
        if len(selected) == 3:
            break
    if len(selected) < 3:
        for _, row in pool.iterrows():
            key = clean(row.get("_issue_key"))
            title = clean(row.get("Headline"))
            if key in used_keys or any(clean(x.get("Headline")) == title for x in selected):
                continue
            selected.append(row)
            used_keys.add(key)
            if len(selected) == 3:
                break
    out = pd.DataFrame(selected).drop(columns=["_top3_score", "_issue_key"], errors="ignore").reset_index(drop=True)
    if not out.empty:
        out["No"] = range(1, len(out) + 1)
    return out


def prepare_rows(rows: pd.DataFrame) -> pd.DataFrame:
    rows = rows.copy()
    rows["Issue"] = rows.apply(issue_for, axis=1)
    rows["Issue"] = rows.apply(_display_issue, axis=1)
    if "Original Summary" not in rows.columns:
        rows["Original Summary"] = rows.get("Summary", "")
    if "Original AI Analysis" not in rows.columns:
        rows["Original AI Analysis"] = rows.get("AI Analysis", "")
    if "Original Action Plan" not in rows.columns:
        rows["Original Action Plan"] = rows.get("Action Plan", "")

    # Regulation은 신규/변경 법규를 모두 보여야 하므로 이슈 중복 압축을 적용하지 않는다.
    reg = rows[rows["Content Type"].eq("Regulation")].copy()
    news = rows[rows["Content Type"].eq("News")].copy()
    news = news[~news["URL"].apply(_is_bad_report_url)].copy()
    news = _dedup_by_issue_or_url(news)

    rows = pd.concat([reg, news], ignore_index=True)
    rows["Mail Group"] = rows["Content Type"].map({"Regulation": GROUP_REGULATION}).fillna(GROUP_NEWS)
    rows["Samsung Impact"] = rows.apply(final_samsung_impact, axis=1)
    rows["Executive Priority"] = rows.apply(executive_priority, axis=1)
    rows.loc[rows["Executive Priority"].eq("REFERENCE"), "Priority Group"] = "REFERENCE"
    rows.loc[~rows["Executive Priority"].eq("REFERENCE"), "Priority Group"] = rows.loc[~rows["Executive Priority"].eq("REFERENCE"), "Executive Priority"]
    rows["Major Changes"] = rows.apply(major_changes, axis=1)
    rows["Summary"] = rows.apply(report_summary, axis=1)
    rows["AI Analysis"] = rows.apply(report_impact, axis=1)
    rows["Action Plan"] = rows.apply(report_action, axis=1)
    rows["_report_score"] = rows.apply(report_score, axis=1)

    reg = rows[rows["Content Type"].eq("Regulation")].copy()
    news = rows[rows["Content Type"].eq("News")].copy()
    news = news.sort_values(["_report_score", "_sort_date"], ascending=[False, False]).head(NEWS_MAX_REPORT_ROWS)
    rows = pd.concat([reg, news], ignore_index=True)
    rows = rows.sort_values(["_report_score", "_sort_date"], ascending=[False, False]).drop(columns=["_issue_key"], errors="ignore").reset_index(drop=True)
    rows["No"] = range(1, len(rows) + 1)
    return rows


def top3_summary_sentence(row: pd.Series) -> str:
    title = clean(row.get("Headline"))
    issue = clean(row.get("Issue")) or issue_for(row)
    text = _text_blob(row).lower()
    if issue == "AD/CVD":
        if any(k in text for k in ["steel", "철강", "zinc", "아연", "도금", "냉간압연"]):
            return f"{title} → 중국산 철강재 조달비용 상승 리스크 확대"
        return f"{title} → 대상 HS·공급국·벤더 기준 AD/CVD 비용 및 원산지 방어자료 점검 필요"
    if issue == "FTA/원산지":
        if "morocco" in text or "모로코" in text:
            return f"{title} → 배터리 공급망 안정화 및 FTA 활용 기회 확대"
        return f"{title} → CO 발급요건, BOM 원산지, 특혜세율 적용 가능성 재검토 필요"
    if issue == "CBAM":
        return f"{title} → EU 수출품 탄소비용 및 CBAM 인증서 구매비용 관리 필요"
    if issue == "수출통제":
        return f"{title} → 핵심소재·부품 공급망과 수출통제 스크리닝 강화 필요"
    if issue in {"통관", "통관/세관"}:
        return f"{title} → 통관신고, 보세운영, 과세가격 자료관리 절차 점검 필요"
    if issue == "관세정책":
        return f"{title} → 관세율·쿼터·공급국 선택에 따른 원가 영향 재산정 필요"
    return f"{title} → 관련 품목의 HS, 원산지, 관세율 영향 확인 필요"


def overall_html(rows: pd.DataFrame, top3: pd.DataFrame) -> str:
    reg = rows[rows["Content Type"].eq("Regulation")]
    news = rows[rows["Content Type"].eq("News")]
    direct = rows[rows["Samsung Impact"].eq("Direct")]
    indirect = rows[rows["Samsung Impact"].eq("Indirect")]
    watch = rows[rows["Samsung Impact"].eq("Watch")]
    ref = rows[rows["Samsung Impact"].eq("Reference")]
    top_lines = "".join(f"<li>{html.escape(top3_summary_sentence(r))}</li>" for _, r in top3.iterrows())
    if not top_lines:
        top_lines = "<li>금일 Top3 후보는 추가 검토가 필요합니다.</li>"
    return f"""
    <div style="background:#F7F9FC;border-left:5px solid #1F4E79;padding:13px 15px;line-height:1.65;">
      <div>
        금일 GTI Radar는 글로벌 통상환경에서 <b>'규제 강화'</b>와 <b>'FTA 확대'</b>가 동시에 진행되고 있으며,
        삼성전자는 원가 리스크 관리와 공급망 다변화 전략을 병행할 필요가 있습니다.
      </div>
      <ul style="margin-top:8px;margin-bottom:8px;padding-left:20px;">{top_lines}</ul>
      <div style="margin-top:8px;color:#777;font-size:12px;">
        * 금일 선별 결과: 법규 {len(reg)}건, 주요뉴스 {len(news)}건 |
        Direct {len(direct)}건, Indirect {len(indirect)}건, Watch {len(watch)}건, Reference {len(ref)}건
      </div>
    </div>
    """


def table_html(title: str, rows: pd.DataFrame, color: str) -> str:
    if rows.empty:
        return f"<h3 style='color:{color};'>{html.escape(title)} (0건)</h3>"
    trs = []
    for _, row in rows.iterrows():
        summary = html.escape(short_text(row.get("Major Changes"), "Summary 확인 필요", 650)).replace("\n", "<br>")
        impact = html.escape(short_text(row.get("AI Analysis"), "영향 검토 필요", 360)).replace("\n", "<br>")
        action = html.escape(short_text(row.get("Action Plan"), "담당 부서 확인 필요", 360)).replace("\n", "<br>")
        trs.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(str(row.get('No')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Issue')))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{html_link(row.get('Headline'), row.get('URL'))}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{summary}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{impact}</td>
          <td style="padding:7px;border:1px solid #ddd;vertical-align:top;">{action}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Country')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;color:{risk_color(row.get('Risk'))};font-weight:bold;">{html.escape(clean(row.get('Risk')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Samsung Impact')))}</td>
          <td style="padding:7px;border:1px solid #ddd;text-align:center;vertical-align:top;">{html.escape(clean(row.get('Date')))}</td>
        </tr>
        """)
    return f"""
    <h3 style="margin-top:24px;color:{color};">{html.escape(title)} ({len(rows)}건)</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;table-layout:fixed;">
      <colgroup>
        <col style="width:3%;"><col style="width:7%;"><col style="width:19%;">
        <col style="width:25%;"><col style="width:16%;"><col style="width:16%;">
        <col style="width:5%;"><col style="width:4%;"><col style="width:5%;"><col style="width:7%;">
      </colgroup>
      <thead>
        <tr style="background:{color};color:white;">
          <th style="padding:7px;border:1px solid #ddd;">No</th>
          <th style="padding:7px;border:1px solid #ddd;">Issue</th>
          <th style="padding:7px;border:1px solid #ddd;">Headline</th>
          <th style="padding:7px;border:1px solid #ddd;">Summary</th>
          <th style="padding:7px;border:1px solid #ddd;">삼성 영향</th>
          <th style="padding:7px;border:1px solid #ddd;">Action</th>
          <th style="padding:7px;border:1px solid #ddd;">Country</th>
          <th style="padding:7px;border:1px solid #ddd;">Risk</th>
          <th style="padding:7px;border:1px solid #ddd;">Impact</th>
          <th style="padding:7px;border:1px solid #ddd;">Publish Date</th>
        </tr>
      </thead>
      <tbody>{''.join(trs)}</tbody>
    </table>
    """


def save_excel(rows: pd.DataFrame, top3: pd.DataFrame, paths: dict[str, Path]) -> None:
    wb = Workbook()
    sheets = [
        ("GTI Radar", rows),
        ("Top3 Deep Analysis", top3),
        ("Regulation", rows[rows["Content Type"].eq("Regulation")]),
        ("주요뉴스", rows[rows["Content Type"].eq("News")]),
    ]
    first = True
    for name, frame in sheets:
        ws = wb.active if first else wb.create_sheet(name[:31])
        first = False
        ws.title = name[:31]
        ws.append(OUTPUT_COLUMNS)
        for _, row in frame.iterrows():
            append_output_row(ws, row)
        style_sheet(ws)

    runlog = wb.create_sheet("Run Log")
    runlog.append(["item", "value"])
    runlog.append(["regulation_input", str(REGULATION_INPUT_FILE)])
    runlog.append(["news_input", str(NEWS_INPUT_FILE)])
    runlog.append(["news_audit_input", str(NEWS_AUDIT_INPUT_FILE)])
    runlog.append(["run_date", RUN_DATE])
    runlog.append(["total_rows", len(rows)])
    runlog.append(["regulation_rows", int(rows["Content Type"].eq("Regulation").sum())])
    runlog.append(["news_rows", int(rows["Content Type"].eq("News").sum())])
    runlog.append(["direct_rows", int(rows["Samsung Impact"].eq("Direct").sum())])
    runlog.append(["indirect_rows", int(rows["Samsung Impact"].eq("Indirect").sum())])
    runlog.append(["watch_rows", int(rows["Samsung Impact"].eq("Watch").sum())])
    runlog.append(["reference_rows", int(rows["Samsung Impact"].eq("Reference").sum())])
    runlog.append(["news_min_report_rows", NEWS_MIN_REPORT_ROWS])
    runlog.append(["news_max_report_rows", NEWS_MAX_REPORT_ROWS])
    style_sheet(runlog)
    wb.save(paths["mail_xlsx"])
    wb.save(paths["analysis"])
    rows[OUTPUT_COLUMNS].to_excel(paths["cumulative"], index=False)


# ======================================================================
# End of GTI STEP5 Final Mail Patch - 2026-06-14
# ======================================================================


def main() -> None:
    paths = output_paths()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    rows = prepare_rows(read_step4_results())
    top3 = choose_top3(rows)
    html_body = build_html(rows, top3)
    save_excel(rows, top3, paths)
    paths["mail_html"].write_text(html_body, encoding="utf-8")
    send_email(html_body, paths["mail_xlsx"])

    reg_n = int(rows["Content Type"].eq("Regulation").sum())
    news_n = int(rows["Content Type"].eq("News").sum())
    print(f"[DONE] HTML: {paths['mail_html']}")
    print(f"[DONE] XLSX: {paths['mail_xlsx']}")
    print(
        f"[ROWS] total={len(rows)}, regulation={reg_n}, news={news_n}, "
        f"direct={(rows['Samsung Impact'] == 'Direct').sum()}, "
        f"indirect={(rows['Samsung Impact'] == 'Indirect').sum()}, "
        f"watch={(rows['Samsung Impact'] == 'Watch').sum()}"
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=None)
    parser.add_argument("--regulation-input", default=None)
    parser.add_argument("--news-input", default=None)
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--no-email", action="store_true")
    args = parser.parse_args()
    if args.date:
        RUN_DATE = args.date
    if args.regulation_input:
        REGULATION_INPUT_FILE = Path(args.regulation_input)
    if args.news_input:
        NEWS_INPUT_FILE = Path(args.news_input)
    if args.output_dir:
        OUTPUT_DIR = Path(args.output_dir)
    if args.no_email:
        SEND_EMAIL = False
    main()
