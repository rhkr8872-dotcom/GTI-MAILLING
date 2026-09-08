# -*- coding: utf-8 -*-
"""GTI STEP5 v44 evidence-gated executive report engine.

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
    out["EventKey"] = "REG_" + out["Headline"].map(lambda x: re.sub(r"\W+", "_", s(x).lower())[:150])
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
        mask.append((bool(url) and "U:" + url in known) or (bool(headline) and "H:" + headline in known))
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
        <p class='meta'>{esc(r.get('Country'))} · {esc(r.get('Agency') or r.get('Source'))} · <a href='{esc(r.get('URL'))}'>원문</a></p></div>""")
    if not cards:
        cards.append("<div class='empty'><b>신규 Action Queue 없음</b><br>Direct 확정 또는 원문 검증을 통과한 우선 검토대상이 없습니다.</div>")

    def table_rows(frame: pd.DataFrame) -> str:
        rendered = []
        for _, r in frame.iterrows():
            rendered.append(f"<tr><td>{esc(r.get('DecisionStatus'))}</td><td><a href='{esc(r.get('URL'))}'>{esc(r.get('Headline'))}</a></td><td>{esc(r.get('Country'))}</td><td>{esc(r.get('ContractReason'))}</td></tr>")
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
    <section class='section verified'><h2>3. 원문 검증 완료 ({len(verified)}건)</h2><table><tr><th>상태</th><th>정책 신호</th><th>국가</th><th>선정 근거</th></tr>{table_rows(verified)}</table></section>
    <section class='section pending'><h2>4. 원문 확인 필요 ({len(pending)}건)</h2><table><tr><th>상태</th><th>정책 신호</th><th>국가</th><th>확인 사유</th></tr>{table_rows(pending)}</table></section>
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
    args = ap.parse_args()
    now_text = os.getenv("GTI_NOW", "").strip()
    now = datetime.fromisoformat(now_text) if now_text else datetime.now()
    run_date = args.date or now.strftime("%Y-%m-%d")
    print("[INFO] GTI STEP5 v44 EVIDENCE-GATED EXECUTIVE ENGINE START")
    news = read_excel_safe(NEWS_FILE); reg = read_excel_safe(REG_FILE)
    news, stale = within_24h(news, now)
    news, rejected = apply_quality_contract(news, include_reference=False)
    reg = normalize_regulation(reg)
    frames = [x for x in (reg, news) if not x.empty]
    rows = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    if not rows.empty:
        rows = rows.sort_values("ExecutiveScore", ascending=False, kind="stable").drop_duplicates("EventKey", keep="first")
        rows = enrich_decision_status(rows)
    old = pd.DataFrame() if args.preview else read_excel_safe(CUM_FILE)
    rows, historical_removed = remove_history(rows, old)
    rows = rows.reset_index(drop=True)
    print(f"[STEP5 CONTRACT] news_input={len(news)+len(rejected)} / selected={len(news)} / rejected={len(rejected)} / stale={len(stale)}")
    print(f"[STEP5 LIVE NOVELTY] removed={historical_removed} / report={len(rows)} / forced_fill=0")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stem = f"[GTI Radar] Global Trade Intelligence({run_date})"; html_path = OUT_DIR / f"{stem}.html"; xlsx_path = OUT_DIR / f"{stem}.xlsx"
    body = build_html(rows, run_date); html_path.write_text(body, encoding="utf-8"); write_xlsx(xlsx_path, rows)
    if not args.preview:
        cumulative = pd.concat([old, rows], ignore_index=True, sort=False)
        if not cumulative.empty: cumulative = cumulative.drop_duplicates(["EventKey"], keep="last")
        cumulative.to_excel(CUM_FILE, index=False)
    if not args.preview and not args.no_email: send_mail(body, xlsx_path, run_date)
    else: print("[MAIL SKIP] preview/no-email")
    print(f"[DONE] HTML: {html_path}"); print(f"[DONE] XLSX: {xlsx_path}")
    print(f"[ROWS] total={len(rows)}, action_required={int(direct_confirmed_mask(rows).sum())}, urgent={int(rows.get('DecisionStatus', pd.Series(dtype=str)).eq('Urgent Verification').sum())}, scenario={int(rows.get('DecisionStatus', pd.Series(dtype=str)).eq('Scenario Analysis').sum())}, monitoring={int(rows.get('DecisionStatus', pd.Series(dtype=str)).eq('Monitoring').sum())}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
