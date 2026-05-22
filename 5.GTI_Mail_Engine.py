# -*- coding: utf-8 -*-
"""
5.GTI Mail Engine.py
GTI STEP5 FINAL - Executive Mail Delivery Engine

개선사항:
- Top3/Top25 선정 기준을 단순 Risk/키워드 방식에서 Samsung Impact Score 방식으로 변경
- 삼성전자 생산거점/판매법인/제품군 직접 영향 중심으로 Top3 선정
- TED signed out, 로그인, 일반 경제, 식품, 시장동향 등 노이즈 강력 제거
- 총평을 고정문구+기사 잘라내기 방식에서 임원용 Impact Summary 방식으로 변경
- 총평 구성: Top3 전체 1줄 총평 + Top3 각 뉴스별 1줄 영향 요약
- EVENT LIST 제목에서 건수 표기 제거

역할:
- STEP4 분석 결과(news_raw.xlsx)를 우선 사용
- Top25 선정
- Executive HTML 메일 생성
- Excel 첨부 생성
- C:\\temp\\00.xlsx 에서 수신자 자동 추출
- GTI_SEND_EMAIL 환경변수 없이 항상 메일 발송 시도
- mail_cumulative.xlsx 저장 실패해도 메일 발송은 계속 진행
"""

from __future__ import annotations

import os
import re
import ssl
import html
import smtplib
import traceback
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr

import pandas as pd

BASE_DIR = Path(r"C:\temp")
INPUT_CANDIDATES = [
    BASE_DIR / "news_raw.xlsx",
]
RECIPIENT_FILE = BASE_DIR / "00.xlsx"
TODAY = datetime.now().strftime("%Y-%m-%d")
SUBJECT = f"[GTI Radar] Global Trade Intelligence | {TODAY}"
OUTPUT_XLSX = BASE_DIR / f"GTI_Radar_{TODAY}_Top25.xlsx"
OUTPUT_HTML = BASE_DIR / f"GTI_Radar_{TODAY}_Top25_Email.html"
MAIL_CUMULATIVE = BASE_DIR / "mail_cumulative.xlsx"
TOP_N = 25
SEND_EMAIL = True

SMTP_HOST = "smtp.naver.com"
SMTP_PORT = 465
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS_DEFAULT = "3GKBVKMZMEKK"
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or SMTP_PASS_DEFAULT).strip()
MAIL_FROM_NAME = "GTI Radar"


# =============================================================================
# BASIC UTILITIES
# =============================================================================

def log(msg: str) -> None:
    print(msg, flush=True)


def clean_text(v) -> str:
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    s = html.unescape(str(v))
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def clean_for_match(v) -> str:
    return clean_text(v).lower()


def safe_date(v) -> str:
    s = clean_text(v)
    if not s:
        return ""
    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return s[:16]
    return dt.strftime("%Y-%m-%d %H:%M")


def normalize_risk(v) -> str:
    s = clean_text(v).upper()
    if "상" in s or "HIGH" in s or "직접" in s:
        return "상"
    if "하" in s or "LOW" in s or "기타" in s:
        return "하"
    if "중" in s or "MED" in s or "간접" in s:
        return "중"
    return "중"


def safe_url(v) -> str:
    s = clean_text(v)
    return s if s.startswith(("http://", "https://")) else ""


def contains_any(text: str, terms: list[str]) -> bool:
    t = text.lower()
    return any(term.lower() in t for term in terms)


def count_hits(text: str, terms: list[str]) -> int:
    t = text.lower()
    return sum(1 for term in terms if term.lower() in t)


def cut_sentence(s: str, max_len: int = 90) -> str:
    s = clean_text(s)
    if not s:
        return ""
    s = re.sub(r"\s+", " ", s).strip()
    if len(s) <= max_len:
        return s
    return s[:max_len].rstrip() + "..."


# =============================================================================
# INPUT / COLUMN NORMALIZATION
# =============================================================================

def find_input_file() -> Path:
    for fp in INPUT_CANDIDATES:
        if fp.exists():
            return fp
    raise FileNotFoundError("입력 파일 없음: news_raw.xlsx")


def make_unique_columns(cols) -> list[str]:
    used, out = {}, []
    for c in cols:
        base = clean_text(c) or "col"
        if base not in used:
            used[base] = 0
            out.append(base)
        else:
            used[base] += 1
            out.append(f"{base}_{used[base]}")
    return out


def read_input(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df.columns = make_unique_columns(df.columns)
    return df.fillna("")


def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for key in candidates:
        if key.lower() in lower_map:
            return lower_map[key.lower()]
    for c in df.columns:
        lc = str(c).strip().lower()
        for key in candidates:
            if key.lower() in lc:
                return c
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_date = pick_col(df, ["Date", "date", "publish date", "published", "뉴스게시일", "뉴스 원본 게시일시"])
    col_headline = pick_col(df, ["Headline", "headline", "Title", "title", "뉴스 제목", "Head"])
    col_summary = pick_col(df, ["Summary", "summary", "요약", "주요내용"])
    col_analysis = pick_col(df, ["AI Analysis", "AI_Analysis", "analysis", "Impact", "전문관세사 분석"])
    col_action = pick_col(df, ["Action Plan", "Action", "Action_Plan", "대응방안"])
    col_country = pick_col(df, ["Country", "country", "국가", "대상 국가"])
    col_agency = pick_col(df, ["agency", "Agency", "관련 기관", "관련기관", "정부기관"])
    col_risk = pick_col(df, ["Risk", "importance", "중요도", "risk"])
    col_url = pick_col(df, ["URL", "url", "Link", "link", "출처URL"])
    col_source = pick_col(df, ["source", "Source", "출처", "date source"])

    out = pd.DataFrame()
    out["Date"] = df[col_date].apply(safe_date) if col_date else ""
    out["Headline"] = df[col_headline].apply(clean_text) if col_headline else ""
    out["Summary"] = df[col_summary].apply(clean_text) if col_summary else ""
    out["AI Analysis"] = df[col_analysis].apply(clean_text) if col_analysis else ""
    out["Action Plan"] = df[col_action].apply(clean_text) if col_action else ""
    out["Country"] = df[col_country].apply(clean_text) if col_country else ""
    out["Agency"] = df[col_agency].apply(clean_text) if col_agency else ""
    out["Risk"] = df[col_risk].apply(normalize_risk) if col_risk else "중"
    out["URL"] = df[col_url].apply(safe_url) if col_url else ""
    out["Source"] = df[col_source].apply(clean_text) if col_source else ""

    for i, r in out.iterrows():
        if not clean_text(r["Headline"]):
            s = clean_text(r["Summary"])
            out.at[i, "Headline"] = (re.split(r"[.!?。]\s+|\n", s)[0][:90] if s else "제목 확인 필요")
        if not clean_text(r["AI Analysis"]):
            out.at[i, "AI Analysis"] = "STEP4 분석값이 비어 있습니다. 원문 및 관련 법령 확인 후 삼성전자 관세·통상 영향을 재검토해야 합니다."
        if not clean_text(r["Action Plan"]):
            out.at[i, "Action Plan"] = "관련 국가·HS·제품군을 확인하고, 법인별 통관·원산지·FTA 적용 영향 여부를 점검합니다."
        if not clean_text(r["Country"]):
            out.at[i, "Country"] = "Global"
        if not clean_text(r["Agency"]):
            out.at[i, "Agency"] = "Relevant customs/trade authority"
    return out


# =============================================================================
# SAMSUNG IMPACT SCORING
# =============================================================================

SAMSUNG_PRODUCTION_TERMS = [
    "samsung", "삼성", "삼성전자",
    "sev", "sevt", "siel", "samex", "sem", "setk", "sehc", "sece",
    "베트남", "vietnam", "mexico", "멕시코", "india", "인도", "china", "중국",
    "poland", "폴란드", "hungary", "헝가리", "brazil", "브라질", "thailand", "태국",
]

SAMSUNG_PRODUCT_TERMS = [
    "semiconductor", "반도체", "hbm", "chip", "chips", "memory", "메모리",
    "mobile", "smartphone", "스마트폰", "phone", "휴대폰",
    "consumer electronics", "가전", "tv", "display", "디스플레이",
    "component", "부품", "battery", "배터리", "pcb", "ccl", "mlcc",
    "network equipment", "네트워크", "telecom", "통신장비",
]

CUSTOMS_TRADE_CORE_TERMS = [
    "tariff", "관세", "customs", "세관", "통관", "duty", "수입세", "수출세",
    "fta", "free trade", "원산지", "origin", "certificate of origin", "coo",
    "hs code", "hscode", "hs ", "classification", "품목분류",
    "anti-dumping", "antidumping", "반덤핑", "countervailing", "상계관세",
    "safeguard", "세이프가드", "section 301", "301조", "section 232", "232조",
    "cbp", "ustr", "mofcom", "관세청", "taxud", "wto", "wco",
]

REGULATION_EFFECTIVE_TERMS = [
    "regulation", "directive", "rule", "decree", "notice", "law", "act", "시행", "규정", "법령", "시행령", "고시",
    "mandatory", "의무", "effective", "발효", "enforcement", "집행", "compliance", "준수",
    "ppwr", "cbam", "uflpa", "eudr", "forced labor", "수출통제", "export control", "ear", "dual-use", "전략물자",
]

COST_IMPACT_TERMS = [
    "cost", "원가", "pricing", "가격", "refund", "환급", "drawback", "duty saving", "관세절감",
    "vat", "부가세", "pva", "deferment", "납부유예", "surcharge", "additional tariff", "추가관세",
    "penalty", "벌금", "fine", "처벌", "audit", "심사", "사후검증", "investigation", "조사",
]

PRIORITY_COUNTRY_TERMS = [
    "united states", "u.s.", "usa", "미국", "cbp", "ustr", "cit",
    "eu", "european union", "유럽", "taxud", "집행위원회",
    "vietnam", "베트남", "india", "인도", "mexico", "멕시코", "china", "중국",
]

HIGH_VALUE_ISSUE_TERMS = [
    "export control", "수출통제", "전략물자", "ear", "dual-use", "희토류", "rare earth",
    "section 301", "301조", "section 232", "232조", "anti-dumping", "반덤핑", "countervailing", "상계관세",
    "uflpa", "forced labor", "강제노동", "cbam", "ppwr", "usmca", "fta", "원산지", "origin",
    "tariff ruling", "관세 판결", "cit", "refund", "환급",
]

STRONG_NOISE_TERMS = [
    "ted successfully signed out", "successfully signed out", "you are signed out", "eu login",
    "call for tenders", "prior information notice", "procurement of data", "planning - ted",
    "로그인", "signed out", "captcha", "cookie", "subscribe", "advertisement",
    "연예", "가수", "콘서트", "instagram", "youtube.com", "twitter", "x.com/search",
    "스포츠", "야구", "축구", "부동산", "주가", "증시", "코스피", "bitcoin", "crypto",
    "소고기", "beef", "seafood", "수산물", "gold", "silver", "금값", "은값", "담배", "마약", "fentanyl",
    "식품", "k-food", "푸드", "식약처", "가성소다", "naoh", "caustic soda",
]

GENERAL_LOW_VALUE_TERMS = [
    "market growth", "시장 성장", "경제 성장", "경상수지", "gdp", "전망", "기회 확대",
    "logistics business", "reshaping rules", "general trade", "일반 무역 동향",
]


def row_text(row: pd.Series) -> str:
    return " ".join(
        clean_text(row.get(c, ""))
        for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Source"]
    )


def is_noise(row: pd.Series) -> bool:
    text = row_text(row).lower()
    headline = clean_text(row.get("Headline", "")).lower()
    summary = clean_text(row.get("Summary", "")).lower()

    if contains_any(text, STRONG_NOISE_TERMS):
        return True

    if ("ted" in headline or "ted" in summary) and contains_any(text, ["signed out", "call for tenders", "procurement", "prior information notice"]):
        return True

    has_customs_core = contains_any(text, CUSTOMS_TRADE_CORE_TERMS + REGULATION_EFFECTIVE_TERMS + HIGH_VALUE_ISSUE_TERMS)
    has_samsung_context = contains_any(text, SAMSUNG_PRODUCTION_TERMS + SAMSUNG_PRODUCT_TERMS)

    if contains_any(text, GENERAL_LOW_VALUE_TERMS) and not (has_customs_core and has_samsung_context):
        return True

    return False


def infer_risk_by_score(score: int, current_risk: str) -> str:
    current_risk = normalize_risk(current_risk)
    if score >= 220:
        return "상"
    if score >= 120:
        return "중"
    if current_risk == "상" and score >= 90:
        return "중"
    return "하"


def samsung_impact_score(row: pd.Series) -> int:
    text = row_text(row).lower()
    score = 0

    if is_noise(row):
        return -999

    risk = normalize_risk(row.get("Risk", "중"))
    score += {"상": 40, "중": 20, "하": 0}.get(risk, 10)

    # 1. 삼성 생산/판매거점 및 제품 직접 영향
    production_hits = count_hits(text, SAMSUNG_PRODUCTION_TERMS)
    product_hits = count_hits(text, SAMSUNG_PRODUCT_TERMS)
    score += min(production_hits, 5) * 25
    score += min(product_hits, 5) * 22

    if production_hits >= 1 and product_hits >= 1:
        score += 60

    # 2. 관세/통관/FTA/원산지 핵심성
    customs_hits = count_hits(text, CUSTOMS_TRADE_CORE_TERMS)
    score += min(customs_hits, 8) * 18

    # 3. 실제 시행 규정/강제성
    regulation_hits = count_hits(text, REGULATION_EFFECTIVE_TERMS)
    score += min(regulation_hits, 6) * 16

    # 4. 원가/환급/추징/처벌/감사 영향
    cost_hits = count_hits(text, COST_IMPACT_TERMS)
    score += min(cost_hits, 6) * 14

    # 5. 미국/EU/베트남/인도/멕시코/중국 우선
    country_hits = count_hits(text, PRIORITY_COUNTRY_TERMS)
    score += min(country_hits, 6) * 10

    # 6. 임원 관심 고위험 이슈
    high_issue_hits = count_hits(text, HIGH_VALUE_ISSUE_TERMS)
    score += min(high_issue_hits, 6) * 22

    # 7. 직접영향 문구 가산
    if contains_any(text, ["직접 영향", "직접적인 영향", "direct impact", "directly affect", "directly impacts"]):
        score += 50
    if contains_any(text, ["북미 판매법인", "eu 판매법인", "생산거점", "생산 법인", "production site", "manufacturing site"]):
        score += 45
    if contains_any(text, ["hs mapping", "hs code mapping", "원산지 판정", "fta 적용", "관세율 재산출"]):
        score += 40

    # 8. 단순 일반론 감점
    if contains_any(text, GENERAL_LOW_VALUE_TERMS):
        score -= 80
    if not contains_any(text, CUSTOMS_TRADE_CORE_TERMS + REGULATION_EFFECTIVE_TERMS + HIGH_VALUE_ISSUE_TERMS):
        score -= 120
    if not contains_any(text, SAMSUNG_PRODUCTION_TERMS + SAMSUNG_PRODUCT_TERMS):
        score -= 60

    return score


def prepare_top25(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["_score"] = df.apply(samsung_impact_score, axis=1)
    df["_noise"] = df.apply(is_noise, axis=1)
    df["_url_key"] = df["URL"].astype(str).str.lower().str.strip()
    df["_headline_key"] = (
        df["Headline"].astype(str).str.lower().str.strip()
        .str.replace(r"[^a-z0-9가-힣]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
    )

    before = len(df)
    df = df[df["_score"] > -300].copy()
    df = df[~df["_noise"]].copy()
    log(f"[FILTER] noise removed={before - len(df)}")

    if df.empty:
        raise ValueError("Top25 후보가 없습니다. STEP4 결과 또는 노이즈 필터 기준을 확인하세요.")

    if df["_url_key"].str.len().gt(0).any():
        df = df.drop_duplicates(subset=["_url_key"], keep="first")
    df = df.drop_duplicates(subset=["_headline_key"], keep="first")

    df["Risk"] = df.apply(lambda r: infer_risk_by_score(int(r["_score"]), r.get("Risk", "중")), axis=1)
    df = df.sort_values(["_score", "Date"], ascending=[False, False])

    top = df.head(TOP_N).copy()
    top.insert(0, "No", range(1, len(top) + 1))
    top["Samsung Impact Score"] = top["_score"].astype(int)

    return top[[
        "No", "Date", "Headline", "Summary", "AI Analysis", "Action Plan",
        "Country", "Agency", "Risk", "Samsung Impact Score", "URL", "Source"
    ]]


# =============================================================================
# EXECUTIVE SUMMARY GENERATION
# =============================================================================

def detect_main_theme(row: pd.Series) -> str:
    text = row_text(row).lower()
    if contains_any(text, ["ppwr", "packaging", "포장", "포장재"]):
        return "EU PPWR"
    if contains_any(text, ["cbam", "탄소국경"]):
        return "EU CBAM"
    if contains_any(text, ["uflpa", "forced labor", "강제노동"]):
        return "UFLPA"
    if contains_any(text, ["export control", "수출통제", "전략물자", "ear", "dual-use"]):
        return "수출통제"
    if contains_any(text, ["section 301", "301조"]):
        return "美 301조 관세"
    if contains_any(text, ["section 232", "232조"]):
        return "美 232조 관세"
    if contains_any(text, ["cit", "tariff ruling", "관세 판결", "refund", "환급"]):
        return "美 관세소송/환급"
    if contains_any(text, ["anti-dumping", "antidumping", "반덤핑"]):
        return "반덤핑"
    if contains_any(text, ["countervailing", "상계관세"]):
        return "상계관세"
    if contains_any(text, ["fta", "free trade", "원산지", "origin", "coo"]):
        return "FTA/원산지"
    if contains_any(text, ["customs", "세관", "통관", "audit", "심사", "penalty", "처벌"]):
        return "통관/세관심사"
    if contains_any(text, ["tariff", "관세", "duty"]):
        return "관세율 변동"
    return "통상 리스크"


def detect_site_product(row: pd.Series) -> str:
    text = row_text(row).lower()
    sites = []
    if contains_any(text, ["vietnam", "베트남", "sev", "sevt"]):
        sites.append("SEV/SEVT")
    if contains_any(text, ["india", "인도", "siel"]):
        sites.append("SIEL")
    if contains_any(text, ["mexico", "멕시코", "samex"]):
        sites.append("SAMEX")
    if contains_any(text, ["poland", "폴란드"]):
        sites.append("폴란드")
    if contains_any(text, ["china", "중국"]):
        sites.append("중국")
    if not sites:
        sites.append("주요 생산거점")

    products = []
    if contains_any(text, ["semiconductor", "반도체", "hbm", "memory", "chip"]):
        products.append("반도체")
    if contains_any(text, ["mobile", "smartphone", "phone", "스마트폰", "휴대폰"]):
        products.append("Mobile")
    if contains_any(text, ["consumer electronics", "가전", "tv"]):
        products.append("CE")
    if contains_any(text, ["display", "디스플레이"]):
        products.append("Display")
    if contains_any(text, ["component", "부품", "pcb", "ccl", "mlcc"]):
        products.append("Component")
    if contains_any(text, ["network", "telecom", "통신장비"]):
        products.append("Network")
    if not products:
        products.append("주요 제품")

    return f"{','.join(sites[:3])} {','.join(products[:3])}"


def executive_impact_summary(row: pd.Series) -> str:
    theme = detect_main_theme(row)
    target = detect_site_product(row)
    text = row_text(row).lower()

    if theme == "EU PPWR":
        return f"EU PPWR 시행 → {target}의 EU향 포장재 규제·BOM·통관증빙 재점검 필요"
    if theme == "EU CBAM":
        return f"EU CBAM 확대 → {target}의 탄소자료·원산지·수입신고 증빙 정합성 확보 필요"
    if theme == "UFLPA":
        return f"UFLPA/강제노동 규제 → {target}의 공급망 원산지 증빙 및 미국 수입통관 리스크 점검 필요"
    if theme == "수출통제":
        return f"수출통제 강화 → {target}의 HS·ECCN·최종사용자 심사 및 출하통제 기준 재확인 필요"
    if theme == "美 301조 관세":
        return f"美 301조 관세 변동 → {target}의 대미 수출 관세부담·원산지·가격전가 영향 재산출 필요"
    if theme == "美 232조 관세":
        return f"美 232조 관세 확대 → {target}의 원재료·부품 관세부담 및 북미 판매법인 원가 영향 점검 필요"
    if theme == "美 관세소송/환급":
        return f"美 관세소송/환급 이슈 → {target}의 기존 납부관세 환급 가능액 및 북미 손익 영향 검토 필요"
    if theme == "반덤핑":
        return f"반덤핑 조사·관세 → {target}의 공급국 전환, HS 분류, 과세가격 및 원산지 방어자료 확보 필요"
    if theme == "상계관세":
        return f"상계관세 리스크 → {target}의 보조금 관련 원산지·가격자료 및 미국/EU 수입규제 대응 필요"
    if theme == "FTA/원산지":
        return f"FTA/원산지 변화 → {target}의 원산지 판정기준, CO 발급, 특혜세율 적용 가능성 재검토 필요"
    if theme == "통관/세관심사":
        return f"통관·세관심사 강화 → {target}의 HS·과세가격·원산지 신고 정확성 및 관세사 업무지침 개정 필요"
    if theme == "관세율 변동":
        return f"관세율 변동 → {target}의 수입원가, 판매가격, FTA 활용 및 생산거점별 관세 시나리오 재산출 필요"

    # fallback: AI Analysis를 임원용 한 줄로 축약
    ai = clean_text(row.get("AI Analysis", ""))
    if ai:
        return cut_sentence(ai, 95)
    return f"{theme} → {target}의 관세·원산지·통관 내부통제 영향 점검 필요"


def build_executive_total_line(top3: pd.DataFrame) -> str:
    themes = [detect_main_theme(r) for _, r in top3.iterrows()]
    theme_txt = "·".join(list(dict.fromkeys(themes))[:3])

    country_text = " ".join(clean_text(r.get("Country", "")) for _, r in top3.iterrows()).lower()
    areas = []
    if contains_any(country_text, ["미국", "us", "u.s.", "usa", "united states"]):
        areas.append("미국")
    if contains_any(country_text, ["eu", "유럽", "europe"]):
        areas.append("EU")
    if contains_any(country_text, ["vietnam", "베트남"]):
        areas.append("베트남")
    if contains_any(country_text, ["mexico", "멕시코"]):
        areas.append("멕시코")
    if contains_any(country_text, ["india", "인도"]):
        areas.append("인도")
    if contains_any(country_text, ["china", "중국"]):
        areas.append("중국")
    if not areas:
        areas = ["미국", "EU", "아시아"]

    return (
        f"{','.join(areas[:3])} 중심의 {theme_txt} 이슈가 삼성전자 주요 생산거점 및 판매법인의 "
        f"관세원가·원산지·FTA·통관 내부통제 재점검 필요성을 높이고 있습니다."
    )


def build_overall_summary(top3: pd.DataFrame) -> str:
    total_line = build_executive_total_line(top3)
    bullets = [f"• {html.escape(executive_impact_summary(r))}" for _, r in top3.iterrows()]
    return f"""
    <div style="margin-top:10px;margin-bottom:20px;padding:16px;background:#F4F6F8;border-left:6px solid #1F4E78;color:#222;">
      <div style="font-size:12pt;font-weight:bold;line-height:1.7;margin-bottom:10px;">
        {html.escape(total_line)}
      </div>
      <div style="margin-top:10px;font-size:11pt;font-weight:normal;line-height:1.8;">{"<br>".join(bullets)}</div>
    </div>
    """


# =============================================================================
# HTML BUILDERS
# =============================================================================

def html_link(headline: str, url: str) -> str:
    h = html.escape(clean_text(headline))
    u = html.escape(clean_text(url))
    return f"<a href='{u}' style='color:#0563C1;text-decoration:underline;font-weight:bold;'>{h}</a>" if u else f"<b>{h}</b>"


def risk_color(risk: str) -> str:
    return {"상": "#C00000", "중": "#ED7D31", "하": "#5B9BD5"}.get(normalize_risk(risk), "#ED7D31")


def build_top3_block(top3: pd.DataFrame) -> str:
    blocks = []
    for i, (_, r) in enumerate(top3.iterrows(), start=1):
        score = clean_text(r.get("Samsung Impact Score", ""))
        blocks.append(f"""
        <div style="margin:16px 0 18px 0;padding:14px;border-left:5px solid #C00000;background:#FFF7F7;">
          <div style="font-size:15px;font-weight:bold;margin-bottom:6px;">{i}️⃣ {html_link(r['Headline'], r['URL'])}</div>
          <div style="font-size:12px;color:#555;margin-bottom:10px;">
            Publish Date: {html.escape(clean_text(r['Date']))} | Country: {html.escape(clean_text(r['Country']))} | Agency: {html.escape(clean_text(r['Agency']))} | Risk: {html.escape(clean_text(r['Risk']))} | Samsung Impact Score: {html.escape(score)}
          </div>
          <div style="margin-top:8px;"><b>Executive Impact</b><br>{html.escape(executive_impact_summary(r))}</div>
          <div style="margin-top:8px;"><b>Summary</b><br>{html.escape(clean_text(r['Summary']))}</div>
          <div style="margin-top:8px;"><b>Impact</b><br>{html.escape(clean_text(r['AI Analysis']))}</div>
          <div style="margin-top:8px;"><b>Action</b><br>{html.escape(clean_text(r['Action Plan']))}</div>
        </div>
        """)
    return "\n".join(blocks)


def build_event_rows(rest: pd.DataFrame) -> str:
    rows = []
    for _, r in rest.iterrows():
        color = risk_color(r["Risk"])
        rows.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{r['No']}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html_link(r['Headline'], r['URL'])}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean_text(r['Summary']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean_text(r['AI Analysis']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean_text(r['Action Plan']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{html.escape(clean_text(r['Country']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean_text(r['Agency']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;color:{color};font-weight:bold;">{html.escape(clean_text(r['Risk']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{html.escape(clean_text(r['Date']))}</td>
        </tr>
        """)
    return "\n".join(rows)


def build_html(top25: pd.DataFrame) -> str:
    top3 = top25.head(3).copy()
    rest = top25.iloc[3:].copy()
    return f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"><title>{html.escape(SUBJECT)}</title></head>
<body style="font-family:Arial,'Malgun Gothic',sans-serif;font-size:13px;color:#222;line-height:1.5;">
  <div style="max-width:1200px;margin:0 auto;">
    <h2 style="margin-bottom:4px;color:#1F4E78;">[GTI Radar] Global Trade Intelligence</h2>
    <div style="font-size:14px;margin-bottom:4px;"><b>Date:</b> {TODAY}</div>
    <div style="font-size:12px;color:#555;margin-bottom:16px;">Coverage: Last 24 Hours | Focus: Samsung Electronics Customs & Trade Intelligence</div>

    <h3 style="margin-top:18px;margin-bottom:6px;">총평</h3>
    {build_overall_summary(top3)}

    <h3 style="color:#C00000;margin-top:22px;">🔴 TOP POLICY EVENTS (Top 3)</h3>
    {build_top3_block(top3)}

    <h3 style="color:#1F4E78;margin-top:24px;">🟦 EVENT LIST</h3>
    <table style="border-collapse:collapse;width:100%;font-size:12px;">
      <tr style="background:#1F4E78;color:white;">
        <th style="padding:7px;border:1px solid #d9d9d9;">No</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Headline</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Summary</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Impact</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Action</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Country</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Agency</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Risk</th>
        <th style="padding:7px;border:1px solid #d9d9d9;">Publish Date</th>
      </tr>
      {build_event_rows(rest)}
    </table>
    <p style="margin-top:18px;color:#666;font-size:12px;">※ 첨부 Excel 파일에 전체 Top25 분석표가 포함되어 있습니다.</p>
  </div>
</body>
</html>"""


# =============================================================================
# OUTPUT / MAIL
# =============================================================================

def save_outputs(top25: pd.DataFrame, html_body: str) -> None:
    top25.to_excel(OUTPUT_XLSX, index=False)
    log(f"[SAVE] Excel: {OUTPUT_XLSX}")
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_body)
    log(f"[SAVE] HTML: {OUTPUT_HTML}")


def load_recipients() -> list[str]:
    if not RECIPIENT_FILE.exists():
        raise FileNotFoundError(f"수신자 파일 없음: {RECIPIENT_FILE}")
    df = pd.read_excel(RECIPIENT_FILE)
    text = "\n".join(df.astype(str).fillna("").values.ravel().tolist())
    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    emails = list(dict.fromkeys([e.strip() for e in emails if e.strip()]))
    if not emails:
        raise ValueError(f"{RECIPIENT_FILE}에서 이메일 주소를 찾지 못했습니다.")
    return emails


def send_email(html_body: str) -> None:
    if not SEND_EMAIL:
        log("[MAIL SKIP] SEND_EMAIL=False")
        return

    recipients = load_recipients()
    if not SMTP_USER:
        raise ValueError("SMTP_USER가 비어 있습니다.")
    if not SMTP_PASS:
        raise ValueError("SMTP 비밀번호가 없습니다. Windows 환경변수 GTI_SMTP_PASS 또는 코드 내 SMTP_PASS_DEFAULT를 설정하세요.")

    msg = EmailMessage()
    msg["Subject"] = SUBJECT
    msg["From"] = formataddr((MAIL_FROM_NAME, SMTP_USER))
    msg["To"] = ", ".join(recipients)
    msg.set_content("GTI Radar 메일입니다. HTML 메일을 지원하는 클라이언트에서 확인해 주세요.")
    msg.add_alternative(html_body, subtype="html")

    if OUTPUT_XLSX.exists():
        with open(OUTPUT_XLSX, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename=OUTPUT_XLSX.name,
            )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    log(f"[MAIL SENT] {len(recipients)} recipients")


def update_mail_cumulative(top25: pd.DataFrame) -> None:
    try:
        data = top25.copy()
        data.insert(0, "mail_date", TODAY)
        data.insert(1, "subject", SUBJECT)
        if MAIL_CUMULATIVE.exists():
            old = pd.read_excel(MAIL_CUMULATIVE)
            data = pd.concat([old, data], ignore_index=True)
            data = data.drop_duplicates(subset=["mail_date", "Headline", "URL"], keep="last")
        data.to_excel(MAIL_CUMULATIVE, index=False)
        log(f"[SAVE] Cumulative: {MAIL_CUMULATIVE}")
    except Exception as e:
        log(f"[WARN] cumulative save skipped: {e}")


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    log("[START] GTI Mail Engine FINAL - Samsung Impact Scoring")
    input_file = find_input_file()
    raw = read_input(input_file)
    log(f"[LOAD] {input_file} rows={len(raw)}")

    norm = normalize_columns(raw)
    top25 = prepare_top25(norm)
    log(f"[SELECT] Top rows={len(top25)}")

    if len(top25) > 0:
        log("[TOP3]")
        for _, r in top25.head(3).iterrows():
            log(f"  - score={r['Samsung Impact Score']} / risk={r['Risk']} / {r['Headline'][:80]}")

    html_body = build_html(top25)
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    save_outputs(top25, html_body)
    update_mail_cumulative(top25)
    send_email(html_body)
    log("[DONE] GTI Mail Engine")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"[ERROR] {e}")
        log(traceback.format_exc())
        raise

try:
    send_mail()
    print("MAIL SENT SUCCESS")
except Exception as e:
    print("MAIL ERROR")
    print(type(e).__name__)
    print(str(e))
    raise
