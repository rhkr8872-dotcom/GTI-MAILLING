# -*- coding: utf-8 -*-
"""
GTI STEP5 Mail Engine

Input:
  - 4.news_ai_analysis.xlsx
Output:
  - GTI_Radar_YYYY-MM-DD_Top30.xlsx
  - GTI_Radar_YYYY-MM-DD_Top30_Email.html
  - mail_cumulative.xlsx

Set GTI_SEND_EMAIL=Y to send mail. SMTP credentials must be provided through
environment variables or GitHub Actions secrets.
"""

from __future__ import annotations

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
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(os.getenv("GTI_OUTPUT_DIR", ".")).resolve()
INPUT_ANALYSIS = Path(os.getenv("GTI_ANALYSIS_FILE", BASE_DIR / "4.news_ai_analysis.xlsx"))
TODAY = os.getenv("GTI_RUN_DATE", datetime.now().strftime("%Y-%m-%d"))
SUBJECT = f"[GTI Radar] Global Trade Intelligence | {TODAY}"

OUTPUT_XLSX = BASE_DIR / f"GTI_Radar_{TODAY}_Top30.xlsx"
OUTPUT_HTML = BASE_DIR / f"GTI_Radar_{TODAY}_Top30_Email.html"
MAIL_CUMULATIVE = BASE_DIR / "mail_cumulative.xlsx"

SMTP_HOST = os.getenv("GTI_SMTP_HOST", "smtp.naver.com")
SMTP_PORT = int(os.getenv("GTI_SMTP_PORT", "465"))
SMTP_USER = (os.getenv("GTI_SMTP_USER") or os.getenv("GTI_MAIL_ID") or "").strip()
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or "").strip()
MAIL_FROM_NAME = os.getenv("GTI_MAIL_FROM_NAME", "GTI Radar")
SEND_EMAIL = os.getenv("GTI_SEND_EMAIL", "N").strip().upper() == "Y"
RECIPIENTS = os.getenv("GTI_MAIL_TO", "").strip()

RISK_ORDER = {"상": 1, "중": 2, "하": 3}


def clean(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def html_link(headline: str, url: str) -> str:
    h = html.escape(clean(headline))
    u = html.escape(clean(url))
    if u:
        return f"<a href=\"{u}\" style=\"color:#0563C1;font-weight:700;text-decoration:underline;\">{h}</a>"
    return h


def load_recipients() -> list[str]:
    emails = []
    if RECIPIENTS:
        emails.extend(x.strip() for x in RECIPIENTS.split(",") if x.strip())
    for candidate in [BASE_DIR / "00.xlsx", BASE_DIR / "mail.xlsx", Path(r"C:\temp\00.xlsx"), Path(r"C:\temp\mail.xlsx")]:
        if candidate.exists():
            try:
                df = pd.read_excel(candidate)
                text = "\n".join(df.astype(str).fillna("").values.ravel().tolist())
                emails.extend(re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text))
            except Exception as exc:
                print(f"[WARN] recipient file skipped: {candidate} / {exc}")
    return list(dict.fromkeys(emails))


def section_for(row: pd.Series) -> str:
    text = f"{row.get('Headline','')} {row.get('Summary','')}".lower()
    if any(x in text for x in ["fta", "origin", "원산지", "customs union"]):
        return "2. FTA·원산지"
    if any(x in text for x in ["통관", "clearance", "valuation", "운임", "신고"]):
        return "3. 통관·심사"
    if any(x in text for x in ["세율", "tariff", "관세", "refund", "duty"]):
        return "1. 관세·세율 영향"
    return "4. 정책 모니터링"


def prepare_mail_table() -> pd.DataFrame:
    if not INPUT_ANALYSIS.exists():
        raise FileNotFoundError(f"Analysis file not found: {INPUT_ANALYSIS}")
    df = pd.read_excel(INPUT_ANALYSIS)
    df = df.rename(columns={"출처URL": "URL", "importance": "Risk"})
    for col in ["Date", "Headline", "Summary", "AI Analysis", "Action Plan", "Country", "agency", "Risk", "URL", "source"]:
        if col not in df.columns:
            df[col] = ""
    df["Section"] = df.apply(section_for, axis=1)
    df["_risk_order"] = df["Risk"].map(RISK_ORDER).fillna(9)
    df = df.sort_values(["Section", "_risk_order", "Country", "Date"], ascending=[True, True, True, False])
    df.insert(0, "No", range(1, len(df) + 1))
    return df[["No", "Section", "Date", "Headline", "Summary", "AI Analysis", "Action Plan", "Country", "agency", "Risk", "URL", "source"]]


def build_html(df: pd.DataFrame) -> str:
    top3 = df.sort_values(["Risk"], key=lambda s: s.map(RISK_ORDER).fillna(9)).head(3)
    top_blocks = []
    for _, row in top3.iterrows():
        top_blocks.append(
            f"""
            <div style="border-left:5px solid #B00020;background:#FFF6F6;padding:14px 16px;margin:12px 0 16px;">
              <div style="font-size:15px;margin-bottom:6px;">{html_link(row['Headline'], row['URL'])}</div>
              <div style="font-size:12px;color:#555;margin-bottom:10px;">발표일: {html.escape(clean(row['Date']))} | 국가: {html.escape(clean(row['Country']))} | 기관: {html.escape(clean(row['agency']))} | 중요도: {html.escape(clean(row['Risk']))}</div>
              <div><b>Summary</b><br>{html.escape(clean(row['Summary']))}</div>
              <div style="margin-top:8px;"><b>AI Analysis</b><br>{html.escape(clean(row['AI Analysis']))}</div>
              <div style="margin-top:8px;"><b>Action</b><br>{html.escape(clean(row['Action Plan']))}</div>
            </div>
            """
        )

    rows = []
    for _, row in df.iterrows():
        rows.append(
            f"""
            <tr>
              <td>{int(row['No'])}</td>
              <td>{html.escape(clean(row['Section']))}</td>
              <td>{html_link(row['Headline'], row['URL'])}</td>
              <td>{html.escape(clean(row['Summary']))}</td>
              <td>{html.escape(clean(row['Date']))}</td>
              <td>{html.escape(clean(row['Country']))}</td>
              <td>{html.escape(clean(row['agency']))}</td>
              <td style="text-align:center;font-weight:700;">{html.escape(clean(row['Risk']))}</td>
              <td>{html.escape(clean(row['Action Plan']))}</td>
            </tr>
            """
        )

    return f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>{html.escape(SUBJECT)}</title>
</head>
<body style="font-family:Arial,'Malgun Gothic',sans-serif;color:#222;font-size:13px;line-height:1.48;">
  <div style="max-width:1280px;margin:0 auto;">
    <h2 style="margin-bottom:4px;">[GTI Radar] Global Trade Intelligence</h2>
    <div style="margin-bottom:14px;"><b>Date:</b> {html.escape(TODAY)}</div>
    <h3 style="margin-top:18px;color:#B00020;">주요 뉴스 Top3</h3>
    {''.join(top_blocks)}
    <h3 style="margin-top:24px;color:#1F4E78;">정책 이벤트 표</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;">
      <thead>
        <tr style="background:#1F4E78;color:#fff;">
          <th>No</th><th>Section</th><th>Headline</th><th>주요내용</th><th>발표일</th>
          <th>대상 국가</th><th>관련 기관</th><th>중요도</th><th>Action</th>
        </tr>
      </thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
  </div>
  <style>
    th, td {{ border:1px solid #D9D9D9; padding:7px; vertical-align:top; }}
    th {{ text-align:center; }}
  </style>
</body>
</html>"""


def save_excel(df: pd.DataFrame) -> None:
    df.to_excel(OUTPUT_XLSX, index=False)
    wb = load_workbook(OUTPUT_XLSX)
    ws = wb.active
    ws.title = "GTI Radar Top30"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    widths = [6, 18, 15, 48, 55, 60, 58, 18, 32, 10, 42, 22]
    for idx, cell in enumerate(ws[1], start=1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(idx)].width = widths[idx - 1]
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
        ws.row_dimensions[row[0].row].height = 96
        headline_cell = row[3]
        url = row[10].value
        if url:
            headline_cell.hyperlink = url
            headline_cell.font = Font(color="0563C1", bold=True, underline="single")
        risk = row[9].value
        row[9].fill = PatternFill("solid", fgColor={"상": "F4CCCC", "중": "FFF2CC", "하": "D9EAD3"}.get(risk, "FFFFFF"))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(OUTPUT_XLSX)


def update_cumulative(df: pd.DataFrame) -> None:
    out = df.copy()
    out.insert(0, "mail_date", TODAY)
    out.insert(1, "subject", SUBJECT)
    if MAIL_CUMULATIVE.exists() and os.getenv("GTI_RESET_CUMULATIVE", "N").upper() != "Y":
        old = pd.read_excel(MAIL_CUMULATIVE)
        out = pd.concat([old, out], ignore_index=True)
    out = out.drop_duplicates(subset=["mail_date", "Headline", "URL"], keep="last")
    out.to_excel(MAIL_CUMULATIVE, index=False)


def send_email(body: str) -> None:
    recipients = load_recipients()
    if not recipients:
        print("[MAIL SKIP] No recipients. Set GTI_MAIL_TO or provide 00.xlsx/mail.xlsx.")
        return
    if not SMTP_USER or not SMTP_PASS:
        print("[MAIL SKIP] Missing SMTP credentials.")
        return
    msg = EmailMessage()
    msg["Subject"] = SUBJECT
    msg["From"] = formataddr((MAIL_FROM_NAME, SMTP_USER))
    msg["To"] = ", ".join(recipients)
    msg.set_content("GTI Radar HTML email. Please view with an HTML-capable mail client.")
    msg.add_alternative(body, subtype="html")
    with open(OUTPUT_XLSX, "rb") as fh:
        msg.add_attachment(
            fh.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=OUTPUT_XLSX.name,
        )
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    print(f"[MAIL SENT] recipients={len(recipients)}")


def main() -> None:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    df = prepare_mail_table()
    save_excel(df)
    body = build_html(df)
    OUTPUT_HTML.write_text(body, encoding="utf-8")
    update_cumulative(df)
    print(f"[SAVE] {OUTPUT_XLSX}")
    print(f"[SAVE] {OUTPUT_HTML}")
    print(f"[SAVE] {MAIL_CUMULATIVE}")
    if SEND_EMAIL:
        send_email(body)
    else:
        print("[MAIL SKIP] GTI_SEND_EMAIL is not Y.")


if __name__ == "__main__":
    main()
