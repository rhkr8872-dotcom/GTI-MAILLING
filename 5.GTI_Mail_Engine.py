# -*- coding: utf-8 -*-
"""GTI STEP5 v45 evidence-gated executive report engine.

One preparation path, one quality contract, one send path.  --preview and
--no-email never mutate cumulative history.
"""
from __future__ import annotations

import argparse
import html
import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd
from gti_quality_contract import apply_quality_contract, VERSION as CONTRACT_VERSION
from gti_report_window import previous_kst_day_mask


BASE = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
OUT_DIR = Path(os.getenv("GTI_OUTPUT_DIR", str(BASE / "12345" / "c_type_outputs")))
NEWS_FILE = BASE / "4-2.news_ai_summary.xlsx"
REG_FILE = BASE / "4-1.regulation_ai_summary.xlsx"
CUM_FILE = BASE / "gti_news_cumulative.xlsx"
RECIPIENT_FILE = BASE / "00.xlsx"
SMTP_HOST = os.getenv("GTI_SMTP_HOST", "smtp.naver.com")
SMTP_PORT = int(os.getenv("GTI_SMTP_PORT", "465"))
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or "").strip()


def s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()

def publication_date(v) -> str:
    """Return the source publication date used for sensing, never the run date."""
    dt = pd.to_datetime(v, errors="coerce")
    if pd.isna(dt):
        return "원문 게시일 확인 필요"
    if dt.hour or dt.minute or dt.second:
        return dt.strftime("%Y-%m-%d %H:%M")
    return dt.strftime("%Y-%m-%d")


def row_publication_date(row: pd.Series) -> str:
    for col in ("Publish Date", "Date", "Original Publish Date"):
        value = row.get(col, "")
        if value is not None and not pd.isna(value) and s(value):
            return publication_date(value)
    return "원문 게시일 확인 필요"


def regulation_event_key(row: pd.Series) -> str:
    """Preserve bill/document identity so same-title bills never collapse."""
    existing = s(row.get("EventKey"))
    identity = s(row.get("DocumentIdentity"))
    bill_no = re.sub(r"\.0$", "", s(row.get("BillNo")))
    url = s(row.get("URL"))
    bill_match = re.search(r"/out/(\d+)/", url)
    if identity:
        return identity.lower()
    if bill_no:
        return f"bill:{bill_no.lower()}"
    if bill_match:
        return f"bill:{bill_match.group(1)}"
    if existing:
        return existing
    headline = re.sub(r"\W+", "_", s(row.get("Headline")).lower())[:130]
    return f"REG_{headline}_{publication_date(row.get('Date'))}"


def read_excel_safe(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(path)
        df.columns = [s(c) for c in df.columns]
        return df
    except Exception as exc:
        raise RuntimeError(f"INPUT UNREADABLE: {path.name}: {type(exc).__name__}: {exc}") from exc


def within_24h(df: pd.DataFrame, now: datetime) -> tuple[pd.DataFrame, pd.DataFrame]:
    if df.empty:
        return df.copy(), df.copy()
    out = df.copy()
    col = "Publish Date" if "Publish Date" in out else "Date"
    keep, out["_published"], _, _ = previous_kst_day_mask(out.get(col), now=now)
    return out[keep].drop(columns="_published"), out[~keep].drop(columns="_published")


def has_substantive_text(value, headline="", min_chars: int = 45) -> bool:
    text = re.sub(r"\s+", " ", s(value)).strip()
    title = re.sub(r"\s+", " ", s(headline)).strip()
    if not text or len(text) < min_chars:
        return False
    if title and re.sub(r"\W+", "", text).lower() == re.sub(r"\W+", "", title).lower():
        return False
    return True


def normalize_regulation(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    out = df.copy()
    if "Headline" not in out:
        for c in ("Title", "Regulation", "제목"):
            if c in out:
                out["Headline"] = out[c]
                break
    for col in ("Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "URL", "Date"):
        if col not in out: out[col] = ""
    out["Content Type"] = "Regulation"
    verified = out.get("Body Verified", pd.Series("N", index=out.index)).astype(str).str.upper().eq("Y")
    score = pd.to_numeric(out.get("Importance Score", 70), errors="coerce").fillna(70)
    out["ExecutiveScore"] = score.where(verified, score.clip(upper=59))
    out["ExecutiveTier"] = "WATCH"
    out.loc[verified & score.ge(85), "ExecutiveTier"] = "PRIORITY_WATCH"
    out["Original Publish Date"] = out["Date"].map(publication_date)
    out["EventKey"] = out.apply(regulation_event_key, axis=1)
    out["ContractReason"] = "공식 원문 검증 완료: 적용범위·시행일 및 삼성 거래 매핑 확인"
    out.loc[~verified, "ContractReason"] = "공식 게시물이나 원문 본문 미확인: 확인 완료 전 경영진 우선정책 승격 금지"
    out["SamsungDirectFlag"] = "N"
    out["DirectConfirmedFlag"] = "N"
    out["OfficialSourceFlag"] = "Y"
    out["VerificationStatus"] = verified.map({True: "VERIFIED", False: "PENDING"})
    out["DecisionStatus"] = verified.map({True: "Urgent Verification", False: "Verification Pending"})
    out["PriorityEligible"] = (verified & score.ge(85)).map({True: "Y", False: "N"})
    return out


def direct_confirmed_mask(rows: pd.DataFrame) -> pd.Series:
    if rows.empty:
        return pd.Series(False, index=rows.index, dtype=bool)
    return rows.get("DirectConfirmedFlag", pd.Series("N", index=rows.index)).astype(str).str.upper().eq("Y")


def enrich_decision_status(rows: pd.DataFrame) -> pd.DataFrame:
    if rows.empty:
        return rows.copy()
    out = rows.copy()
    derived_verification = out.get("Body Verified", pd.Series("N", index=out.index)).astype(str).str.upper().eq("Y").map({True: "VERIFIED", False: "PENDING"})
    if "DecisionStatus" not in out: out["DecisionStatus"] = "Monitoring"
    else: out["DecisionStatus"] = out["DecisionStatus"].where(out["DecisionStatus"].fillna("").astype(str).str.strip().ne(""), "Monitoring")
    if "VerificationStatus" not in out: out["VerificationStatus"] = derived_verification
    else: out["VerificationStatus"] = out["VerificationStatus"].where(out["VerificationStatus"].fillna("").astype(str).str.strip().ne(""), derived_verification)
    if "PriorityEligible" not in out: out["PriorityEligible"] = "N"
    else: out["PriorityEligible"] = out["PriorityEligible"].fillna("N")
    news = out.get("Content Type", pd.Series("", index=out.index)).astype(str).str.lower().eq("news")
    direct = direct_confirmed_mask(out)
    proposed = out.get("Policy Stage", pd.Series("", index=out.index)).astype(str).str.upper().eq("PROPOSED_OR_MONITORING")
    verified = out["VerificationStatus"].eq("VERIFIED")
    score = pd.to_numeric(out.get("ExecutiveScore", 0), errors="coerce").fillna(0)
    priority_signal = out.get("ExecutiveTier", pd.Series("", index=out.index)).isin(["EXECUTIVE", "PRIORITY_WATCH"])
    out.loc[news & verified, "DecisionStatus"] = "Monitoring"
    out.loc[news & verified & proposed & priority_signal, "DecisionStatus"] = "Scenario Analysis"
    out.loc[news & verified & ~proposed & priority_signal & ~direct, "DecisionStatus"] = "Urgent Verification"
    out.loc[direct, "DecisionStatus"] = "Action Required"
    out.loc[news & verified & (direct | priority_signal) & score.ge(70), "PriorityEligible"] = "Y"
    return out


def select_priority(rows: pd.DataFrame, limit: int = 3) -> pd.DataFrame:
    if rows.empty:
        return rows.copy()
    eligible = rows[rows.get("PriorityEligible", pd.Series("N", index=rows.index)).astype(str).str.upper().eq("Y")].copy()
    rank = {"Action Required": 0, "Urgent Verification": 1, "Scenario Analysis": 2, "Monitoring": 3}
    eligible["_decision_rank"] = eligible.get("DecisionStatus", "Monitoring").map(rank).fillna(9)
    eligible = eligible.sort_values(["_decision_rank", "ExecutiveScore"], ascending=[True, False], kind="stable")
    return eligible.head(limit).drop(columns="_decision_rank")


def historical_keys(df: pd.DataFrame) -> set[str]:
    keys: set[str] = set()
    for _, r in df.iterrows():
        url = s(r.get("URL")).lower()
        headline = re.sub(r"\W+", "", s(r.get("Headline")).lower())
        if url: keys.add("U:" + url)
        if headline: keys.add("H:" + headline)
    return keys


def remove_history(rows: pd.DataFrame, old: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    if rows.empty or old.empty:
        return rows.copy(), 0
    known = historical_keys(old)
    mask = []
    for _, r in rows.iterrows():
        url = s(r.get("URL")).lower()
        headline = re.sub(r"\W+", "", s(r.get("Headline")).lower())
        # A valid official URL is the document identity.  Headline fallback is
        # used only when URL is absent, because separate bills can share the
        # exact same title (for example multiple 관세법 일부개정법률안).
        mask.append(("U:" + url in known) if url else (bool(headline) and "H:" + headline in known))
    dup = pd.Series(mask, index=rows.index)
    return rows[~dup].copy(), int(dup.sum())


def executive_sentence(rows: pd.DataFrame) -> str:
    if rows.empty:
        return "최근 24시간 내 새로 확인된 삼성전자 관세·통상 핵심 조치는 없습니다. 기존 고위험 정책은 변동 여부를 계속 모니터링합니다."
    axes = []
    priority = select_priority(rows, 3)
    source = priority if not priority.empty else rows.head(3)
    for _, r in source.iterrows():
        fam = s(r.get("PolicyFamily"))
        if fam == "SEMICONDUCTOR_TARIFF": axes.append("미국 반도체 관세와 현지생산 조건")
        elif fam == "CUSTOMS_PROCEDURE": axes.append("베트남 통관절차 변경")
        elif fam == "ORIGIN_TRANSshipment": axes.append("원산지·우회수출 검증")
        elif fam == "TARIFF": axes.append("주요국 관세협상 후속조건")
    axes = list(dict.fromkeys(axes))
    joined = "·".join(axes) if axes else "관세·통상 정책 변화"
    return f"금일 핵심 센싱은 {joined}입니다. HQ Customs는 시행조건과 대상 품목·법인·거래경로를 확인하고, 확정 전 사안은 비용 시나리오와 증빙 준비 수준으로 관리해야 합니다."


def action_text(row: pd.Series) -> str:
    fam = s(row.get("PolicyFamily"))
    if fam == "SEMICONDUCTOR_TARIFF":
        return "대미 반도체 품목별 관세 Cost와 미국 생산 전환 손익을 시나리오별 산출하고 발표·시행 Trigger를 지정"
    if fam == "CUSTOMS_PROCEDURE":
        return "베트남 법인과 Circular 원문·적용 국경·시행일을 확인하고 통관 SOP 및 시스템 변경사항을 Gap 분석"
    if fam == "ORIGIN_TRANSshipment":