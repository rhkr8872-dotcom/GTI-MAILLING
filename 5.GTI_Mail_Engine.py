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
    return dt.strftime("%Y-%m-%d")


def row_publication_date(row: pd.Series) -> str:
    for col in ("Original Publish Date", "Publish Date", "Date"):
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
    out["_published"] = pd.to_datetime(out.get(col), errors="coerce")
    cutoff = pd.Timestamp(now - timedelta(hours=24))
    future_limit = pd.Timestamp(now + timedelta(hours=2))
    keep = out["_published"].between(cutoff, future_limit, inclusive="both")
    return out[keep].drop(columns="_published"), out[~keep].drop(columns="_published")


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
        return "베트남 생산품의 BOM·원산지·제조공정·선적경로를 연결한 Origin Traceability 증빙 점검"
    return s(row.get("Action Plan")) or "원문·적용범위·시행일을 확인하고 관련 법인 영향도를 재판정"


def esc(v) -> str:
    return html.escape(s(v))


def build_html(rows: pd.DataFrame, run_date: str) -> str:
    priority = select_priority(rows, 3)
    direct_n = int(direct_confirmed_mask(rows).sum())
    urgent_n = int(rows.get("DecisionStatus", pd.Series(dtype=str)).eq("Urgent Verification").sum())
    scenario_n = int(rows.get("DecisionStatus", pd.Series(dtype=str)).eq("Scenario Analysis").sum())
    monitoring_n = int(rows.get("DecisionStatus", pd.Series(dtype=str)).eq("Monitoring").sum())
    pending_n = int(rows.get("DecisionStatus", pd.Series(dtype=str)).eq("Verification Pending").sum())
    cards = []
    for i, (_, r) in enumerate(priority.iterrows(), 1):
        badge = esc(r.get("DecisionStatus"))
        cards.append(f"""
        <div class='card'><div class='badge'>{badge}</div>
        <h3>{i}. {esc(r.get('Headline'))}</h3>
        <p><b>임원 판단</b> {esc(r.get('ContractReason'))}</p>
        <p><b>삼성전자 관세업무</b> {esc(r.get('AI Analysis')) or '직접 비용은 미확정이며 적용범위 검증이 필요합니다.'}</p>
        <p><b>지시사항</b> {esc(action_text(r))}</p>
        <p class='meta'><b>원본 게시일자</b> {esc(row_publication_date(r))} · {esc(r.get('Country'))} · {esc(r.get('Agency') or r.get('Source'))} · <a href='{esc(r.get('URL'))}'>원문</a></p></div>""")
    if not cards:
        cards.append("<div class='empty'><b>신규 Action Queue 없음</b><br>Direct 확정 또는 원문 검증을 통과한 우선 검토대상이 없습니다.</div>")

    def table_rows(frame: pd.DataFrame) -> str:
        rendered = []
        for _, r in frame.iterrows():
            original_date = row_publication_date(r)
            rendered.append(f"<tr><td>{esc(r.get('DecisionStatus'))}</td><td>{esc(original_date)}</td><td><a href='{esc(r.get('URL'))}'>{esc(r.get('Headline'))}</a></td><td>{esc(r.get('Country'))}</td><td>{esc(r.get('ContractReason'))}</td></tr>")
        return "".join(rendered)

    verified = rows[rows.get("VerificationStatus", pd.Series("PENDING", index=rows.index)).eq("VERIFIED")]
    pending = rows[rows.get("VerificationStatus", pd.Series("PENDING", index=rows.index)).ne("VERIFIED")]
    return f"""<!doctype html><html><head><meta charset='utf-8'><style>
    body{{font-family:Arial,'Malgun Gothic',sans-serif;color:#172033;max-width:980px;margin:auto;padding:24px;background:#f5f7fb}}
    header,.section{{background:white;border-radius:12px;padding:22px;margin-bottom:14px}} h1{{margin:0;color:#123b70}} h2{{color:#123b70}}
    .lead{{font-size:17px;line-height:1.65;border-left:5px solid #1d63b7;padding:12px 16px;background:#eef5ff}}
    .metric{{display:inline-block;margin-right:18px;font-weight:bold}} .card{{border:1px solid #dbe4ef;border-radius:10px;padding:16px;margin:12px 0}}
    .badge{{display:inline-block;background:#e8f1ff;color:#174f91;border-radius:10px;padding:4px 9px;font-size:12px}} .meta{{font-size:12px;color:#687386}}
    .pending{{border-left:5px solid #d79b00}} .verified{{border-left:5px solid #2b8a3e}}
    table{{width:100%;border-collapse:collapse}} th,td{{padding:9px;border-bottom:1px solid #e4e8ef;text-align:left;font-size:13px}} .empty{{padding:18px;background:#f7f8fa}}
    </style></head><body><header><h1>[GTI Radar] Global Trade Intelligence</h1><p>{run_date} | Samsung Electronics Customs Executive Brief</p></header>
    <section class='section'><h2>1. 오늘의 관세정책 센싱</h2><p class='lead'><b>{esc(executive_sentence(rows))}</b></p>
    <span class='metric'>보고 {len(rows)}건</span><span class='metric'>Action Required {direct_n}건</span><span class='metric'>Urgent Verification {urgent_n}건</span><span class='metric'>Scenario {scenario_n}건</span><span class='metric'>Monitoring {monitoring_n}건</span><span class='metric'>Verification Pending {pending_n}건</span></section>
    <section class='section'><h2>2. Samsung Customs Action Queue</h2>{''.join(cards)}</section>
    <section class='section verified'><h2>3. 원문 검증 완료 ({len(verified)}건)</h2><table><tr><th>상태</th><th>원본 게시일자</th><th>정책 신호</th><th>국가</th><th>선정 근거</th></tr>{table_rows(verified)}</table></section>
    <section class='section pending'><h2>4. 원문 확인 필요 ({len(pending)}건)</h2><table><tr><th>상태</th><th>원본 게시일자</th><th>정책 신호</th><th>국가</th><th>확인 사유</th></tr>{table_rows(pending)}</table></section>
    <section class='section'><small>정책 존재는 기사 원문·공식출처로만 판정하며 AI 분석문은 증거로 사용하지 않습니다. Contract {CONTRACT_VERSION}</small></section></body></html>"""


def write_xlsx(path: Path, rows: pd.DataFrame) -> None:
    top3 = select_priority(rows, 3)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        rows.to_excel(writer, sheet_name="Executive Radar", index=False)
        top3.to_excel(writer, sheet_name="Priority Watch Top3", index=False)
        if rows.empty:
            pd.DataFrame({"Message": ["금일 신규 핵심정책 없음"]}).to_excel(writer, sheet_name="Run Summary", index=False)
        else:
            pd.DataFrame({"Metric": ["Selected", "Action Required", "Urgent Verification", "Scenario Analysis", "Monitoring", "Verification Pending", "Contract"], "Value": [len(rows), int(direct_confirmed_mask(rows).sum()), int(rows["DecisionStatus"].eq("Urgent Verification").sum()), int(rows["DecisionStatus"].eq("Scenario Analysis").sum()), int(rows["DecisionStatus"].eq("Monitoring").sum()), int(rows["DecisionStatus"].eq("Verification Pending").sum()), CONTRACT_VERSION]}).to_excel(writer, sheet_name="Run Summary", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.row_dimensions[1].height = 24
            for cell in ws[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="1F4E78")
            for col in ws.columns:
                letter = col[0].column_letter
                ws.column_dimensions[letter].width = min(55, max(12, max(len(s(c.value)) for c in col) + 2))


# ---------------------------------------------------------------------------
# v50 fixed-form executive report contract
# ---------------------------------------------------------------------------
ENGINE_VERSION = "v50.3 WATCH-REASON-REPAIR"
HEALTH_FILE = BASE / "1.site_crawl_health.xlsx"


def yn(df: pd.DataFrame, col: str) -> pd.Series:
    return df.get(col, pd.Series("N", index=df.index)).fillna("N").astype(str).str.upper().eq("Y")


def text_blob(df: pd.DataFrame) -> pd.Series:
    cols = [c for c in ["Headline", "Summary", "AI Analysis", "Issue", "Country", "Agency", "Source"] if c in df]
    if not cols:
        return pd.Series("", index=df.index)
    out = pd.Series("", index=df.index)
    for col in cols:
        out = out + " " + df[col].fillna("").astype(str)
    return out.str.lower()


def classify_report_layers(rows: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    """Apply the immutable report hierarchy without forced filling.

    Action Queue requires all three gates. Core policy requires current delta
    plus official evidence. Samsung Watch holds business-relevant candidates
    with one or more gates pending. Remaining valid policy context is compact.
    """
    if rows.empty:
        empty = rows.copy()
        return empty, {k: empty.copy() for k in ["action", "core", "watch", "context", "excluded"]}
    d = rows.copy()
    for col, default in [
        ("Content Type", "News"), ("ReportLayer", "REFERENCE"),
        ("PolicyDeltaFlag", "N"), ("EvidenceGateFlag", "N"),
        ("SamsungTradeGateFlag", "N"), ("DirectConfirmedFlag", "N"),
        ("Body Verified", "N"), ("ContractReason", ""), ("DecisionStatus", "Monitoring"),
    ]:
        if col not in d: d[col] = default
    d["_score"] = pd.to_numeric(d.get("ExecutiveScore", d.get("Importance Score", 0)), errors="coerce").fillna(0)
    blob = text_blob(d)
    native = pd.Series("", index=d.index)
    # Samsung scope must come from article-native evidence. AI summaries often
    # mention Samsung hypothetically and must never create relevance by itself.
    for col in ["Headline", "Direct Evidence", "Article Body Evidence", "Country", "Agency"]:
        if col in d: native = native + " " + d[col].fillna("").astype(str)
    native = native.str.lower()
    reg = d["Content Type"].astype(str).str.lower().eq("regulation")
    delta = yn(d, "PolicyDeltaFlag")
    evidence = yn(d, "EvidenceGateFlag")
    body = yn(d, "Body Verified")
    trade = yn(d, "SamsungTradeGateFlag")
    direct = yn(d, "DirectConfirmedFlag") | trade

    # A daily STEP4-1 regulation row is already a newly detected official
    # document. Its body status controls evidence quality, not policy novelty.
    d.loc[reg, "PolicyDeltaFlag"] = "Y"
    delta = delta | reg
    # Official regulations are evidence-backed only when the body itself was
    # obtained. A listing title alone is not enough.
    evidence = evidence | (reg & body)
    d.loc[reg & body, "EvidenceGateFlag"] = "Y"

    customs_operation = blob.str.contains(
        r"외국환거래규정|수출입대금|납세신고\s*정정|수입신고\s*정정|전자문서\s*변경|5fe|5fk|품목번호\s*연계표",
        regex=True, na=False,
    )
    d.loc[customs_operation, "PolicyDeltaFlag"] = "Y"
    delta = delta | customs_operation

    # Defensive event-family merge. The same KCS table or FTA event is often
    # reported by several outlets. Propagate the strongest gate evidence to
    # the family and keep one representative article.
    d["_family"] = d.get("EventKey", pd.Series("", index=d.index)).fillna("").astype(str)
    kcs_family = native.str.contains(r"품목번호\s*연계표|hs code cross-reference|한.?미\s*품목번호", regex=True, na=False) & native.str.contains(r"관세청|korea customs service", regex=True, na=False)
    fta_family = native.str.contains("메르코수르", regex=False, na=False) & native.str.contains("싱가포르", regex=False, na=False) & native.str.contains("fta", regex=False, na=False)
    d.loc[kcs_family, "_family"] = "KR_US_ADDITIONAL_TARIFF_HS_CROSS_REFERENCE"
    d.loc[fta_family, "_family"] = "MERCOSUR_SINGAPORE_FTA_EFFECTIVE"
    for family in ["KR_US_ADDITIONAL_TARIFF_HS_CROSS_REFERENCE", "MERCOSUR_SINGAPORE_FTA_EFFECTIVE"]:
        members = d["_family"].eq(family)
        if not members.any(): continue
        if yn(d.loc[members], "PolicyDeltaFlag").any() or family.startswith("KR_US_"):
            d.loc[members, "PolicyDeltaFlag"] = "Y"
        if yn(d.loc[members], "EvidenceGateFlag").any():
            d.loc[members, "EvidenceGateFlag"] = "Y"
    d["_evidence_rank"] = yn(d, "EvidenceGateFlag").astype(int)
    d["_delta_rank"] = yn(d, "PolicyDeltaFlag").astype(int)
    d = d.sort_values(["_evidence_rank", "_delta_rank", "_score"], ascending=[False, False, False], kind="stable")
    has_family = d["_family"].astype(str).str.strip().ne("")
    family_rows = d[has_family].drop_duplicates("_family", keep="first")
    d = pd.concat([family_rows, d[~has_family]], axis=0).sort_values("_score", ascending=False, kind="stable")
    blob = text_blob(d)
    native = pd.Series("", index=d.index)
    for col in ["Headline", "Direct Evidence", "Article Body Evidence", "Country", "Agency"]:
        if col in d: native = native + " " + d[col].fillna("").astype(str)
    native = native.str.lower()
    reg = d["Content Type"].astype(str).str.lower().eq("regulation")
    delta = yn(d, "PolicyDeltaFlag")
    evidence = yn(d, "EvidenceGateFlag") | (reg & yn(d, "Body Verified"))
    trade = yn(d, "SamsungTradeGateFlag")
    direct = yn(d, "DirectConfirmedFlag") | trade

    samsung_scope = native.str.contains(
        r"삼성전자|samsung electronics|반도체|semiconductor|디스플레이|스마트폰|휴대폰|가전|네트워크장비|"
        r"품목번호\s*연계표|한.?미\s*품목번호|메르코수르.{0,20}싱가포르|싱가포르.{0,20}메르코수르|"
        r"외국환거래규정|수출입대금|납세신고\s*정정|수입신고\s*정정|전자문서\s*변경|5fe|5fk|"
        r"통관지원|세관상호지원|관세협력|합동단속|위조상품|국경단계|지식재산권\s*보호",
        regex=True, na=False,
    )
    declared_excluded = d["ReportLayer"].astype(str).str.upper().eq("EXCLUDED") | d.get("DecisionStatus", "").astype(str).str.lower().eq("excluded")
    substantive_customs = native.str.contains(
        r"통관지원|통관\s*지원|세관상호지원|세관\s*상호지원|관세협력|관세\s*협력|"
        r"합동단속|합동\s*단속|위조상품|위조\s*상품|국경단계|지식재산권\s*보호|"
        r"customs\s*(?:assistance|cooperation|enforcement)|counterfeit",
        regex=True, na=False,
    )
    administrative_only = native.str.contains(
        r"채용|입찰|교육생\s*모집|설명회\s*개최|행사\s*안내|공고\s*마감|"
        r"recruitment|tender\s+notice|training\s+registration",
        regex=True, na=False,
    ) & ~substantive_customs
    rescued_customs = declared_excluded & substantive_customs
    d.loc[rescued_customs, "ReportLayer"] = "SAMSUNG_WATCH"
    d.loc[rescued_customs, "ContractReason"] = (
        "관세당국 간 통관지원·단속 협력 내용 확인: 삼성전자 관련 품목·법인·적용범위 추가 검증 필요"
    )
    d.loc[rescued_customs, "Missing Facts"] = (
        "공식 MOU·협정 원문 | 대상 품목·법인 | 적용 시점 | 통관지원 연락창구·단속 절차"
    )
    excluded_mask = (declared_excluded & ~substantive_customs) | administrative_only
    action_mask = ~excluded_mask & delta & evidence & direct
    core_mask = ~excluded_mask & ~action_mask & delta & evidence
    watch_mask = ~excluded_mask & ~action_mask & ~core_mask & samsung_scope

    action = d[action_mask].sort_values("_score", ascending=False, kind="stable").head(3).copy()
    core = d[core_mask].sort_values("_score", ascending=False, kind="stable").head(5).copy()
    watch = d[watch_mask].sort_values("_score", ascending=False, kind="stable").head(10).copy()
    watch = sort_by_country_and_publish_date(watch)
    used = set(action.index) | set(core.index) | set(watch.index)
    context_candidates = d[~excluded_mask & ~d.index.isin(used)].copy()
    context = context_candidates.sort_values("_score", ascending=False, kind="stable").head(20).copy()
    used |= set(context.index)
    excluded = d[excluded_mask | ~d.index.isin(used)].copy()

    for frame, layer, status in [
        (action, "ACTION_QUEUE", "Action Required"),
        (core, "CORE_POLICY", "Core Policy"),
        (watch, "SAMSUNG_WATCH", "Verification Pending"),
        (context, "GLOBAL_CONTEXT", "Monitoring"),
        (excluded, "EXCLUDED", "Excluded"),
    ]:
        frame["ReportSection"] = layer
        frame["DecisionStatus"] = status
    visible = pd.concat([action, core, watch, context], ignore_index=True, sort=False)
    helper_cols = ["_score", "_family", "_evidence_rank", "_delta_rank"]
    visible = visible.drop(columns=helper_cols, errors="ignore")
    layers = {"action": action.drop(columns=helper_cols, errors="ignore"),
              "core": core.drop(columns=helper_cols, errors="ignore"),
              "watch": watch.drop(columns=helper_cols, errors="ignore"),
              "context": context.drop(columns=helper_cols, errors="ignore"),
              "excluded": excluded.drop(columns=helper_cols, errors="ignore")}
    return visible, layers


def concise(v, limit=180) -> str:
    value = s(v)
    return value if len(value) <= limit else value[:limit - 1].rstrip() + "…"


def health_summary() -> dict[str, int]:
    result = {"OK_NEW": 0, "NO_NEW": 0, "PARSE_ZERO": 0, "FAIL": 0}
    health = read_excel_safe(HEALTH_FILE)
    if health.empty:
        return result
    status_col = next((c for c in ["HealthStatus", "health_status", "final_status", "status", "zero_yield_status"] if c in health), None)
    if not status_col:
        return result
    site_col = next((c for c in [
        "SiteKey", "site_key", "Site ID", "site_id", "Site", "site",
        "SiteName", "site_name", "Agency", "Source", "URL", "url",
    ] if c in health), None)
    time_col = next((c for c in [
        "DiagnosedAt", "diagnosed_at", "CheckedAt", "checked_at",
        "RunDateTime", "run_datetime", "Timestamp", "timestamp",
        "CrawlTime", "crawl_time", "DateTime", "datetime", "Date", "date",
    ] if c in health), None)
    latest = health.copy()
    if site_col:
        latest["_site_key"] = latest[site_col].fillna("").astype(str).str.strip()
        blank = latest["_site_key"].eq("")
        latest.loc[blank, "_site_key"] = [f"__ROW_{i}" for i in latest.index[blank]]
        latest["_row_order"] = range(len(latest))
        if time_col:
            latest["_diagnosed_at"] = pd.to_datetime(latest[time_col], errors="coerce")
            latest = latest.sort_values(
                ["_site_key", "_diagnosed_at", "_row_order"],
                ascending=[True, True, True], na_position="first", kind="stable",
            )
        latest = latest.drop_duplicates("_site_key", keep="last")
    values = latest[status_col].fillna("").astype(str).str.upper().str.strip()
    result["OK_NEW"] = int(values.str.contains(r"^OK|VALID_REGULATION", regex=True).sum())
    result["NO_NEW"] = int(values.str.contains("NO_NEW", regex=False).sum())
    result["PARSE_ZERO"] = int(values.str.contains("PARSE_ZERO|PARTIAL_COVERAGE", regex=True).sum())
    result["FAIL"] = int(values.str.contains("FAIL|BLOCKED|ERROR", regex=True).sum())
    return result


def conclusion_lines(layers: dict[str, pd.DataFrame], health: dict[str, int]) -> list[str]:
    action, core, watch = layers["action"], layers["core"], layers["watch"]
    context = layers["context"]
    line1 = (f"금일 삼성전자 본·지사 거래에 직접 영향이 확인된 신규 관세정책은 {len(action)}건입니다."
             if len(action) else "금일 삼성전자 본·지사 거래에 직접 영향이 확인된 신규 관세정책은 없습니다.")
    if len(core) or len(watch):
        names = [concise(x, 42) for x in pd.concat([core, watch]).get("Headline", pd.Series(dtype=str)).head(2)]
        line2 = f"공식근거가 확인된 핵심 정책 {len(core)}건과 삼성 관련 확인 후보 {len(watch)}건을 센싱했습니다"
        if names: line2 += ": " + " / ".join(names)
        line2 += "."
    elif len(context):
        names = [concise(x, 42) for x in context.get("Headline", pd.Series(dtype=str)).head(2)]
        line2 = f"확정 정책은 아니지만 글로벌 관세·통상 동향 {len(context)}건을 모니터링 대상으로 확인했습니다"
        if names: line2 += ": " + " / ".join(names)
        line2 += "."
    else:
        line2 = "금일 보고기준을 충족한 신규 관세·통상 동향은 없습니다."
    if len(action):
        line3 = "HQ Customs는 직접 영향 항목의 대상법인·품목·거래경로와 실행기한을 즉시 확정해야 합니다."
    elif len(core) or len(watch):
        line3 = "HQ Customs는 공식 원문과 적용 시점, 대상 품목·법인 및 통관지원·단속 절차를 확인해야 합니다."
    elif len(context):
        line3 = "HQ Customs는 전망성 보도를 확정 정책과 구분하고 공식 발표 여부만 후속 확인해야 합니다."
    else:
        line3 = "기존 고위험 정책의 변동 여부와 수집 실패 사이트를 계속 확인해야 합니다."
    return [line1, line2, line3]


def summary_counter(layers: dict[str, pd.DataFrame], health: dict[str, int]) -> str:
    return (f"※ 삼성전자 본·지사 직접 영향 {len(layers['action'])}건, 핵심 정책 {len(layers['core'])}건, "
            f"삼성 관련 확인 후보 {len(layers['watch'])}건, 수집상태 FAIL {health['FAIL']}건, "
            f"PARSE_ZERO {health['PARSE_ZERO']}건 재확인 필요.")


def sort_by_country_and_publish_date(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    out = frame.copy()
    out["_country_sort"] = out.get("Country", pd.Series("", index=out.index)).fillna("").astype(str).str.casefold()
    date_values = pd.Series(pd.NaT, index=out.index, dtype="datetime64[ns]")
    # Prefer timestamp-bearing source fields for newest-first ordering. The
    # normalized Original Publish Date may intentionally contain date only.
    for col in ("Publish Date", "Date", "Original Publish Date"):
        if col in out:
            date_values = date_values.fillna(pd.to_datetime(out[col], errors="coerce"))
    out["_date_sort"] = date_values
    return out.sort_values(
        ["_country_sort", "_date_sort"], ascending=[True, False],
        na_position="last", kind="stable",
    ).drop(columns=["_country_sort", "_date_sort"])


def empty_message(label: str) -> str:
    return {
        "action": "금일 삼성전자 거래에 직접 영향과 실행조치가 모두 확인된 정책은 없습니다.",
        "core": "공식근거가 확인된 신규 중요 정책은 없습니다.",
        "watch": "추가 확인이 필요한 삼성전자 관련 관세정책 후보는 없습니다.",
        "context": "보고기준을 충족한 글로벌 관세정책 변화는 없습니다.",
    }[label]


def build_html_v50(layers: dict[str, pd.DataFrame], run_date: str, health: dict[str, int]) -> str:
    action, core, watch, context = (layers[k] for k in ["action", "core", "watch", "context"])
    watch = sort_by_country_and_publish_date(watch)
    lines = conclusion_lines(layers, health)
    counter = summary_counter(layers, health)
    def empty(k): return f"<div class='empty'>{esc(empty_message(k))}</div>"
    action_html = []
    for _, r in action.iterrows():
        action_html.append(f"<div class='card'><h3><a href='{esc(r.get('URL'))}'>{esc(r.get('Headline'))}</a></h3><p><b>직접영향</b> {esc(concise(r.get('AI Analysis'),260))}</p><p><b>조치</b> {esc(concise(r.get('Action Plan'),220))}</p><p class='meta'>원본 게시일 {esc(row_publication_date(r))} · {esc(r.get('Country'))} · {esc(r.get('Agency'))}</p></div>")
    core_html = []
    for _, r in core.iterrows():
        core_html.append(f"<div class='card'><h3><a href='{esc(r.get('URL'))}'>{esc(r.get('Headline'))}</a></h3><p><b>정책 변화</b> {esc(concise(r.get('Summary'),240))}</p><p><b>삼성 관세 시사점</b> {esc(concise(r.get('AI Analysis'),260))}</p><p class='meta'>원본 게시일 {esc(row_publication_date(r))} · {esc(r.get('Country'))} · {esc(r.get('Agency'))}</p></div>")
    def table(frame, cols):
        heads = "".join(f"<th>{esc(h)}</th>" for h, _ in cols)
        body = []
        for _, r in frame.iterrows():
            cells=[]
            for h,c in cols:
                val = concise(r.get(c), 115)
                if c == "Headline": val=f"<a href='{esc(r.get('URL'))}'>{esc(val)}</a>"
                else: val=esc(val)
                cells.append(f"<td>{val}</td>")
            body.append("<tr>"+"".join(cells)+"</tr>")
        return f"<table><tr>{heads}</tr>{''.join(body)}</table>"
    for frame in (watch, context):
        if not frame.empty:
            frame["_Display Publish Date"] = frame.apply(row_publication_date, axis=1)
    watch_table = table(watch, [("국가/권역","Country"),("원문 게시일","_Display Publish Date"),("상태","DecisionStatus"),("정책 신호","Headline"),("삼성 관련성","ContractReason"),("추가 확인","Missing Facts")]) if len(watch) else empty("watch")
    context_table = table(context, [("국가/권역","Country"),("원문 게시일","_Display Publish Date"),("정책유형","Issue"),("정책 동향","Headline"),("근거상태","OfficialSourceStatus")]) if len(context) else empty("context")
    return f"""<!doctype html><html><head><meta charset='utf-8'><style>
body{{font-family:Arial,'Malgun Gothic',sans-serif;color:#172033;max-width:1040px;margin:auto;padding:24px;background:#f5f7fb}}header,.section{{background:#fff;border-radius:12px;padding:22px;margin-bottom:14px}}h1{{margin:0;color:#123b70}}h2{{color:#123b70}}.lead{{font-size:16px;line-height:1.7;border-left:5px solid #1d63b7;padding:12px 16px;background:#eef5ff}}.summary-counter{{font-size:11px;color:#7a828c;line-height:1.55;margin:9px 0 0 16px}}.card{{border:1px solid #dbe4ef;border-radius:9px;padding:14px;margin:10px 0}}.meta{{font-size:12px;color:#687386}}table{{width:100%;border-collapse:collapse}}th,td{{padding:8px;border-bottom:1px solid #e4e8ef;text-align:left;font-size:12px;vertical-align:top}}.empty{{padding:16px;background:#f7f8fa;color:#53606f}}
</style></head><body><header><h1>[GTI Radar] Global Trade Intelligence</h1><p>{run_date} | Samsung Electronics Customs Executive Brief</p></header>
<section class='section'><h2>1. 요약</h2><div class='lead'>{'<br>'.join(esc(x) for x in lines)}</div><p class='summary-counter'>{esc(counter)}</p></section>
<section class='section'><h2>2. Samsung Action Queue ({len(action)}건)</h2>{''.join(action_html) if action_html else empty('action')}</section>
<section class='section'><h2>3. 핵심 정책 분석 ({len(core)}건)</h2>{''.join(core_html) if core_html else empty('core')}</section>
<section class='section'><h2>4. Samsung Customs Watch ({len(watch)}건)</h2>{watch_table}</section>
<section class='section'><h2>5. Global Context Radar ({len(context)}건)</h2>{context_table}</section>
<section class='section'><h2>6. 수집·검증 상태</h2><table><tr><th>OK_NEW</th><th>NO_NEW</th><th>PARSE_ZERO</th><th>FAIL</th></tr><tr><td>{health['OK_NEW']}건</td><td>{health['NO_NEW']}건</td><td>{health['PARSE_ZERO']}건</td><td>{health['FAIL']}건</td></tr></table><p class='meta'>0건과 수집 실패를 구분합니다. 상세 분석·미확인 사항은 첨부 Excel에 보관합니다. Engine {ENGINE_VERSION} · Contract {CONTRACT_VERSION}</p></section></body></html>"""


def write_xlsx_v50(path: Path, visible: pd.DataFrame, layers: dict[str, pd.DataFrame], health: dict[str, int]) -> None:
    summary = pd.DataFrame({"Metric": ["Action Queue","Core Policy","Samsung Watch","Global Context","Total Visible","OK_NEW","NO_NEW","PARSE_ZERO","FAIL","Engine","Contract"],
                            "Value": [len(layers['action']),len(layers['core']),len(layers['watch']),len(layers['context']),len(visible),health['OK_NEW'],health['NO_NEW'],health['PARSE_ZERO'],health['FAIL'],ENGINE_VERSION,CONTRACT_VERSION]})
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Executive Summary", index=False)
        layers["action"].to_excel(writer, sheet_name="Action Queue", index=False)
        layers["core"].to_excel(writer, sheet_name="Core Policy", index=False)
        layers["watch"].to_excel(writer, sheet_name="Samsung Watch", index=False)
        layers["context"].to_excel(writer, sheet_name="Global Context", index=False)
        visible.to_excel(writer, sheet_name="Details", index=False)
        layers["excluded"].to_excel(writer, sheet_name="Excluded", index=False)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = __import__("openpyxl").styles.Font(bold=True, color="FFFFFF")
                cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="1F4E78")
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(55, max(12, max(len(s(c.value)) for c in col) + 2))


def recipients() -> list[str]:
    found = [x.strip() for x in re.split(r"[;,]", os.getenv("GTI_MAIL_TO", "")) if x.strip()]
    if RECIPIENT_FILE.exists():
        try:
            for value in pd.read_excel(RECIPIENT_FILE).astype(str).to_numpy().ravel():
                if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", value.strip()): found.append(value.strip())
        except Exception: pass
    return list(dict.fromkeys(x.lower() for x in found))


def send_mail(body: str, xlsx: Path, run_date: str) -> None:
    to = recipients()
    if not to: raise RuntimeError("MAIL RECIPIENT MISSING: GTI_MAIL_TO or 00.xlsx")
    if not SMTP_USER or not SMTP_PASS: raise RuntimeError("SMTP CREDENTIAL MISSING")
    msg = EmailMessage()
    msg["Subject"] = f"[GTI Radar] Global Trade Intelligence({run_date})"
    msg["From"] = formataddr(("GTI Radar", SMTP_USER)); msg["To"] = ", ".join(to)
    msg.set_content("GTI Radar HTML report and XLSX are attached."); msg.add_alternative(body, subtype="html")
    msg.add_attachment(xlsx.read_bytes(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=xlsx.name)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context(), timeout=30) as server:
        server.login(SMTP_USER, SMTP_PASS); server.send_message(msg)
    print(f"[MAIL SENT] {len(to)} recipients")


def main() -> int:
    ap = argparse.ArgumentParser(); ap.add_argument("--preview", action="store_true"); ap.add_argument("--no-email", action="store_true"); ap.add_argument("--date")
    ap.add_argument("--regulation-input", default=str(REG_FILE)); ap.add_argument("--news-input", default=str(NEWS_FILE)); ap.add_argument("--output-dir", default=str(OUT_DIR))
    args = ap.parse_args()
    now_text = os.getenv("GTI_NOW", "").strip()
    now = datetime.fromisoformat(now_text) if now_text else datetime.now()
    run_date = args.date or now.strftime("%Y-%m-%d")
    print(f"[INFO] GTI STEP5 {ENGINE_VERSION} START")
    news = read_excel_safe(Path(args.news_input)); reg = read_excel_safe(Path(args.regulation_input))
    news, stale = within_24h(news, now)
    # Keep STEP4 exclusions until STEP5 classification. This permits the
    # narrow substantive-customs rescue rule while preserving genuine
    # administrative exclusions in the Excluded audit sheet.
    input_news_count = len(news)
    reg = normalize_regulation(reg)
    frames = [x for x in (reg, news) if not x.empty]
    rows = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    if not rows.empty:
        rows["Original Publish Date"] = rows.apply(row_publication_date, axis=1)
        if "ExecutiveScore" not in rows:
            rows["ExecutiveScore"] = pd.to_numeric(rows.get("Importance Score", 0), errors="coerce").fillna(0)
        rows = rows.sort_values("ExecutiveScore", ascending=False, kind="stable").drop_duplicates("EventKey", keep="first")
    old = pd.DataFrame() if args.preview else read_excel_safe(CUM_FILE)
    rows, historical_removed = remove_history(rows, old)
    rows = rows.reset_index(drop=True)
    visible, layers = classify_report_layers(rows)
    health = health_summary()
    print(f"[STEP5 CONTRACT] news_input={input_news_count} / classified={len(news)} / excluded={len(layers['excluded'])} / stale={len(stale)}")
    print(f"[STEP5 LIVE NOVELTY] removed={historical_removed} / report={len(visible)} / forced_fill=0")
    output_dir = Path(args.output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"[GTI Radar] Global Trade Intelligence({run_date})"; html_path = output_dir / f"{stem}.html"; xlsx_path = output_dir / f"{stem}.xlsx"
    body = build_html_v50(layers, run_date, health); html_path.write_text(body, encoding="utf-8"); write_xlsx_v50(xlsx_path, visible, layers, health)
    if not args.preview:
        cumulative = pd.concat([old, visible], ignore_index=True, sort=False)
        if not cumulative.empty: cumulative = cumulative.drop_duplicates(["EventKey"], keep="last")
        cumulative.to_excel(CUM_FILE, index=False)
    if not args.preview and not args.no_email: send_mail(body, xlsx_path, run_date)
    else: print("[MAIL SKIP] preview/no-email")
    print(f"[DONE] HTML: {html_path}"); print(f"[DONE] XLSX: {xlsx_path}")
    print(f"[ROWS] total={len(visible)}, action_queue={len(layers['action'])}, core_policy={len(layers['core'])}, samsung_watch={len(layers['watch'])}, global_context={len(layers['context'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
