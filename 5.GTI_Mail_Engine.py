# -*- coding: utf-8 -*-
"""
5.GTI_Mail_Engine_v12.py
GTI STEP5 FINAL v12 - Executive Mail Delivery Engine

목적:
- STEP4 v12 구조화 결과(news_raw.xlsx)의 category / priority / issue_key / score를 그대로 존중
- STEP5에서 다시 임의 scoring하여 품질이 낮아지는 문제 제거
- Top3는 서로 다른 issue/theme/category 중심으로 선정
- 총평 bullet 반복 방지
- 면세점/커피/농산물/비료/방산행사/의약품/교육/관광/마약압수 등 강제 제외
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
INPUT_FILE = BASE_DIR / "news_raw.xlsx"
RECIPIENT_FILE = BASE_DIR / "00.xlsx"

TODAY = datetime.now().strftime("%Y-%m-%d")
SUBJECT = f"[GTI Radar] Global Trade Intelligence | {TODAY}"

OUTPUT_XLSX = BASE_DIR / f"GTI_Radar_{TODAY}_Top30.xlsx"
OUTPUT_HTML = BASE_DIR / f"GTI_Radar_{TODAY}_Top30_Email.html"
MAIL_CUMULATIVE = BASE_DIR / "mail_cumulative.xlsx"

TOP_N = 30
SEND_EMAIL = True

SMTP_HOST = "smtp.naver.com"
SMTP_PORT = 465
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS_DEFAULT = "3GKBVKMZMEKK"
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or SMTP_PASS_DEFAULT).strip()
MAIL_FROM_NAME = "GTI Radar"


CATEGORY_ORDER = {
    "A_LEGAL_OFFICIAL": 1,
    "B_SEMICONDUCTOR_TARIFF": 2,
    "C_ORIGIN_FTA_USMCA": 3,
    "D_EXPORT_CONTROL_SANCTION": 4,
    "E_AD_CVD_TRADE_REMEDY": 5,
    "F_CUSTOMS_AUDIT_VALUATION": 6,
    "G_CBAM_SUPPLY_CHAIN": 7,
    "H_SAMSUNG_GEO_POLICY": 8,
    "I_GENERAL_REFERENCE": 9,
}

CATEGORY_QUOTA = {
    "A_LEGAL_OFFICIAL": 8,
    "B_SEMICONDUCTOR_TARIFF": 5,
    "C_ORIGIN_FTA_USMCA": 5,
    "D_EXPORT_CONTROL_SANCTION": 4,
    "E_AD_CVD_TRADE_REMEDY": 4,
    "F_CUSTOMS_AUDIT_VALUATION": 3,
    "G_CBAM_SUPPLY_CHAIN": 3,
    "H_SAMSUNG_GEO_POLICY": 4,
    "I_GENERAL_REFERENCE": 2,
}

STRICT_REJECT_TERMS = [
    "롯데면세점", "면세점", "duty free", "관광", "tourism", "공공캐릭터", "팝업존",
    "교육", "세미나", "설명회", "컨퍼런스", "워크숍", "인재 양성", "전문가 배출",
    "합격", "training", "webinar", "seminar", "workshop", "conference",
    "커피", "coffee", "cocoa", "beef", "소고기", "쇠고기", "농산물", "farmer",
    "agriculture", "농업", "비료", "fertilizer", "seafood", "수산물", "fish",
    "마약", "fentanyl", "drug", "narcotic", "cocaine", "cannabis",
    "firearm", "gun", "weapon", "총기", "방산", "방위산업", "호위함", "잠수함", "무기",
    "dx korea", "의약품", "제약", "약품", "medicine", "pharma",
    "맛집", "연예", "드라마", "축구", "야구", "증시", "주가", "부동산",
    "채용", "모집", "수주", "보안검색 장비",
    "child porn", "porn", "아동 포르노", "입국 통관 중단", "immigration sanctuary",
]

REJECT_OVERRIDE_TERMS = [
    "federal register", "final rule", "proposed rule", "입법예고", "행정예고",
    "관보", "고시", "공고", "section 301", "section 232", "uflpa",
    "수출통제", "export control", "sanction", "제재", "반덤핑", "상계관세",
    "anti-dumping", "countervailing", "원산지", "usmca", "ustr", "cbp notice",
]

OFFICIAL_TERMS = [
    "ustr", "cbp", "federal register", "관세청", "wto", "eu commission",
    "european commission", "official journal", "taxud", "bis", "mofcom",
    "department of commerce", "상무부",
]

SEMICON_TERMS = ["semiconductor", "chip", "반도체", "hbm", "dram", "nand", "memory"]
EXPORT_CONTROL_TERMS = ["export control", "수출통제", "sanction", "제재", "ear", "bis", "entity list", "dual-use"]
FTA_ORIGIN_TERMS = ["fta", "cepa", "epa", "원산지", "origin", "rules of origin", "usmca", "coo"]
AD_CVD_TERMS = ["anti-dumping", "antidumping", "반덤핑", "countervailing", "상계관세"]
CUSTOMS_TERMS = ["customs", "세관", "통관", "cbp", "valuation", "과세가격", "classification", "품목분류", "audit", "심사"]
CBAM_SUPPLY_TERMS = ["cbam", "carbon", "탄소", "supply chain", "공급망", "rare earth", "희토류"]


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


def low(v) -> str:
    return clean_text(v).lower()


def contains_any(text: str, terms: list[str]) -> bool:
    t = low(text)
    return any(term.lower() in t for term in terms)


def safe_int(v, default: int = 9999) -> int:
    try:
        if pd.isna(v):
            return default
    except Exception:
        pass
    try:
        return int(float(str(v).strip()))
    except Exception:
        return default


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
    if "상" in s or "HIGH" in s:
        return "상"
    if "하" in s or "LOW" in s:
        return "하"
    if "중" in s or "MED" in s:
        return "중"
    return "중"


def safe_url(v) -> str:
    s = clean_text(v)
    return s if s.startswith(("http://", "https://")) else ""


def is_strict_noise(row: pd.Series) -> bool:
    text = " ".join(clean_text(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Source"])
    if contains_any(text, REJECT_OVERRIDE_TERMS):
        return False
    return contains_any(text, STRICT_REJECT_TERMS)


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


def read_input() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일 없음: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    df.columns = make_unique_columns(df.columns)
    return df.fillna("")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_date = pick_col(df, ["Date", "date", "publish date", "published", "뉴스게시일"])
    col_headline = pick_col(df, ["Headline", "headline", "Title", "title", "뉴스 제목", "Head"])
    col_summary = pick_col(df, ["Summary", "summary", "요약", "주요내용"])
    col_analysis = pick_col(df, ["AI Analysis", "AI_Analysis", "analysis", "Impact", "전문관세사 분석"])
    col_action = pick_col(df, ["Action Plan", "Action", "Action_Plan", "대응방안"])
    col_country = pick_col(df, ["Country", "country", "국가", "대상 국가"])
    col_agency = pick_col(df, ["agency", "Agency", "관련 기관", "관련기관", "정부기관"])
    col_risk = pick_col(df, ["Risk", "risk"])
    col_url = pick_col(df, ["URL", "url", "Link", "link", "출처URL"])
    col_source = pick_col(df, ["source", "Source", "출처", "date source"])
    col_score = pick_col(df, ["score"])
    col_priority = pick_col(df, ["priority"])
    col_category = pick_col(df, ["category"])
    col_news_type = pick_col(df, ["news_type"])
    col_issue_key = pick_col(df, ["issue_key", "cluster_key"])
    col_importance = pick_col(df, ["importance", "중요도"])

    out = pd.DataFrame()
    out["Date"] = df[col_date].apply(safe_date) if col_date else ""
    out["Headline"] = df[col_headline].apply(clean_text) if col_headline else ""
    out["Summary"] = df[col_summary].apply(clean_text) if col_summary else ""
    out["AI Analysis"] = df[col_analysis].apply(clean_text) if col_analysis else ""
    out["Action Plan"] = df[col_action].apply(clean_text) if col_action else ""
    out["Country"] = df[col_country].apply(clean_text) if col_country else "Global"
    out["Agency"] = df[col_agency].apply(clean_text) if col_agency else "Relevant customs/trade authority"
    out["Risk"] = df[col_risk].apply(normalize_risk) if col_risk else "중"
    out["URL"] = df[col_url].apply(safe_url) if col_url else ""
    out["Source"] = df[col_source].apply(clean_text) if col_source else ""
    out["score"] = df[col_score].apply(lambda x: safe_int(x, 0)) if col_score else 0
    out["priority"] = df[col_priority].apply(lambda x: safe_int(x, 9)) if col_priority else 9
    out["category"] = df[col_category].apply(clean_text) if col_category else ""
    out["news_type"] = df[col_news_type].apply(clean_text) if col_news_type else ""
    out["issue_key"] = df[col_issue_key].apply(clean_text) if col_issue_key else ""
    out["importance"] = df[col_importance].apply(clean_text) if col_importance else ""

    for i, r in out.iterrows():
        if not clean_text(r["Headline"]):
            s = clean_text(r["Summary"])
            out.at[i, "Headline"] = (re.split(r"[.!?。]\s+|\n", s)[0][:90] if s else "제목 확인 필요")
        if not clean_text(r["Country"]):
            out.at[i, "Country"] = "Global"
        if not clean_text(r["Agency"]):
            out.at[i, "Agency"] = "Relevant customs/trade authority"
        if not clean_text(r["category"]):
            out.at[i, "category"] = infer_category(r)
        if not clean_text(r["news_type"]):
            out.at[i, "news_type"] = infer_news_type(r)
        if not clean_text(r["issue_key"]):
            out.at[i, "issue_key"] = make_issue_key(r)
        if not clean_text(r["importance"]):
            out.at[i, "importance"] = infer_importance(out.at[i, "category"])

    return out


def infer_category(row: pd.Series) -> str:
    text = " ".join(clean_text(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Agency", "Source"])
    if contains_any(text, OFFICIAL_TERMS) and contains_any(text, ["관세", "tariff", "customs", "fta", "origin", "원산지", "rule", "notice", "regulation"]):
        return "A_LEGAL_OFFICIAL"
    if contains_any(text, SEMICON_TERMS) and contains_any(text, ["관세", "tariff", "232", "301", "ustr", "수출통제", "export control"]):
        return "B_SEMICONDUCTOR_TARIFF"
    if contains_any(text, FTA_ORIGIN_TERMS):
        return "C_ORIGIN_FTA_USMCA"
    if contains_any(text, EXPORT_CONTROL_TERMS):
        return "D_EXPORT_CONTROL_SANCTION"
    if contains_any(text, AD_CVD_TERMS):
        return "E_AD_CVD_TRADE_REMEDY"
    if contains_any(text, CUSTOMS_TERMS):
        return "F_CUSTOMS_AUDIT_VALUATION"
    if contains_any(text, CBAM_SUPPLY_TERMS):
        return "G_CBAM_SUPPLY_CHAIN"
    return "I_GENERAL_REFERENCE"


def infer_news_type(row: pd.Series) -> str:
    cat = clean_text(row.get("category", ""))
    if cat == "A_LEGAL_OFFICIAL":
        return "LEGAL_OFFICIAL"
    if cat == "B_SEMICONDUCTOR_TARIFF":
        return "SEMICONDUCTOR"
    if cat == "C_ORIGIN_FTA_USMCA":
        return "FTA_ORIGIN"
    if cat == "D_EXPORT_CONTROL_SANCTION":
        return "EXPORT_CONTROL"
    if cat == "E_AD_CVD_TRADE_REMEDY":
        return "AD_CVD"
    if cat == "F_CUSTOMS_AUDIT_VALUATION":
        return "CUSTOMS_AUDIT"
    if cat == "G_CBAM_SUPPLY_CHAIN":
        return "CBAM_SUPPLY_CHAIN"
    return "GENERAL_TRADE"


def make_issue_key(row: pd.Series) -> str:
    text = low(row.get("Headline", ""))
    rules = [
        ("USMCA_ORIGIN", ["usmca"]),
        ("USMCA_ORIGIN", ["원산지", "멕시코"]),
        ("USTR_CHIP_TARIFF", ["ustr", "chip"]),
        ("USTR_CHIP_TARIFF", ["반도체", "관세"]),
        ("SECTION_232_SEMICON", ["232", "semiconductor"]),
        ("US_CHINA_TARIFF", ["미국", "중국", "관세"]),
        ("US_CHINA_TARIFF", ["china", "us", "tariff"]),
        ("EXPORT_CONTROL_CHINA", ["수출통제", "중국"]),
        ("AD_CVD_STEEL", ["반덤핑", "철강"]),
        ("ASEAN_CUSTOMS_TRANSIT", ["아세안", "관세", "환승"]),
        ("MEXICO_EU_FTA", ["mexico", "eu", "tariff"]),
        ("MEXICO_EU_FTA", ["멕시코", "eu", "관세"]),
        ("EU_STEEL_TARIFF", ["eu", "철강", "관세"]),
        ("CHINA_SEMICON_EXPORT", ["중국", "반도체", "수출"]),
        ("RARE_EARTH_VIETNAM", ["베트남", "희토류"]),
    ]
    for key, terms in rules:
        if all(t in text for t in terms):
            return key
    return clean_text(row.get("category", "I_GENERAL_REFERENCE")) + "_" + re.sub(r"[^a-z0-9가-힣]+", "_", text)[:60]


def infer_importance(category: str) -> str:
    if category in ["A_LEGAL_OFFICIAL", "B_SEMICONDUCTOR_TARIFF", "C_ORIGIN_FTA_USMCA", "D_EXPORT_CONTROL_SANCTION", "E_AD_CVD_TRADE_REMEDY"]:
        return "상"
    if category in ["F_CUSTOMS_AUDIT_VALUATION", "G_CBAM_SUPPLY_CHAIN", "H_SAMSUNG_GEO_POLICY"]:
        return "중"
    return "하"


def final_rank_score(row: pd.Series) -> int:
    cat = clean_text(row.get("category", "I_GENERAL_REFERENCE"))
    priority = CATEGORY_ORDER.get(cat, 9)
    base = 100000 - priority * 10000
    score = safe_int(row.get("score", 0), 0)

    text = " ".join(clean_text(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency"])
    bonus = 0

    if contains_any(text, OFFICIAL_TERMS):
        bonus += 3000
    if contains_any(text, SEMICON_TERMS):
        bonus += 2600
    if contains_any(text, ["usmca", "원산지", "origin", "fta"]):
        bonus += 2300
    if contains_any(text, EXPORT_CONTROL_TERMS):
        bonus += 2200
    if contains_any(text, AD_CVD_TERMS):
        bonus += 2100
    if contains_any(text, ["section 301", "301", "section 232", "232", "추가관세"]):
        bonus += 2000
    if contains_any(text, ["SEV", "SEVT", "SIEL", "SAMEX", "베트남", "인도", "멕시코", "미국", "EU", "중국"]):
        bonus += 1200

    if is_strict_noise(row):
        bonus -= 1000000

    return base + score + bonus


def prepare_top30(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    before = len(df)
    df["_noise"] = df.apply(is_strict_noise, axis=1)
    df = df[~df["_noise"]].copy()
    log(f"[FILTER] strict noise removed={before - len(df)}")

    df["_rank_score"] = df.apply(final_rank_score, axis=1)
    df["_url_key"] = df["URL"].astype(str).str.lower().str.strip()
    df["_headline_key"] = (
        df["Headline"].astype(str).str.lower().str.strip()
        .str.replace(r"[^a-z0-9가-힣]+", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
    )

    if df["_url_key"].str.len().gt(0).any():
        df = df.drop_duplicates(subset=["_url_key"], keep="first")
    df = df.drop_duplicates(subset=["_headline_key"], keep="first")

    selected = []
    selected_issues = set()
    category_counts = {k: 0 for k in CATEGORY_QUOTA}
    source_counts = {}
    korea_count = 0

    ordered = df.sort_values(["priority", "_rank_score"], ascending=[True, False])

    for _, r in ordered.iterrows():
        category = clean_text(r.get("category", "I_GENERAL_REFERENCE"))
        issue = clean_text(r.get("issue_key", ""))
        source = clean_text(r.get("Source", ""))[:45]
        country = clean_text(r.get("Country", ""))

        if category_counts.get(category, 0) >= CATEGORY_QUOTA.get(category, 2):
            continue
        if issue and issue in selected_issues:
            continue
        if source and source_counts.get(source, 0) >= 5:
            continue
        if "한국" in country and korea_count >= 6:
            continue

        selected.append(r)
        if issue:
            selected_issues.add(issue)
        category_counts[category] = category_counts.get(category, 0) + 1
        source_counts[source] = source_counts.get(source, 0) + 1
        if "한국" in country:
            korea_count += 1

        if len(selected) >= TOP_N:
            break

    if len(selected) < TOP_N:
        for _, r in ordered.iterrows():
            issue = clean_text(r.get("issue_key", ""))
            if issue and issue in selected_issues:
                continue
            selected.append(r)
            if issue:
                selected_issues.add(issue)
            if len(selected) >= TOP_N:
                break

    top = pd.DataFrame(selected).reset_index(drop=True)
    top.insert(0, "No", range(1, len(top) + 1))
    top["Samsung Impact Score"] = top["_rank_score"].astype(int)

    return top[[
        "No", "Date", "Headline", "Summary", "AI Analysis", "Action Plan",
        "Country", "Agency", "Risk", "Samsung Impact Score", "URL", "Source",
        "category", "news_type", "issue_key", "priority", "score"
    ]]


def select_top3(top30: pd.DataFrame) -> pd.DataFrame:
    selected = []
    used_themes = set()
    used_issues = set()

    # 우선순위: 공식/반도체/원산지/수출통제/AD 순서로 다양성 확보
    preferred_categories = [
        "A_LEGAL_OFFICIAL",
        "B_SEMICONDUCTOR_TARIFF",
        "C_ORIGIN_FTA_USMCA",
        "D_EXPORT_CONTROL_SANCTION",
        "E_AD_CVD_TRADE_REMEDY",
        "F_CUSTOMS_AUDIT_VALUATION",
        "G_CBAM_SUPPLY_CHAIN",
        "H_SAMSUNG_GEO_POLICY",
    ]

    for cat in preferred_categories:
        cand = top30[top30["category"] == cat].sort_values("Samsung Impact Score", ascending=False)
        for _, r in cand.iterrows():
            theme = detect_main_theme(r)
            issue = clean_text(r.get("issue_key", ""))
            if theme in used_themes or issue in used_issues:
                continue
            selected.append(r)
            used_themes.add(theme)
            used_issues.add(issue)
            break
        if len(selected) >= 3:
            break

    if len(selected) < 3:
        for _, r in top30.sort_values("Samsung Impact Score", ascending=False).iterrows():
            theme = detect_main_theme(r)
            issue = clean_text(r.get("issue_key", ""))
            if issue in used_issues:
                continue
            if len(selected) < 3:
                selected.append(r)
                used_themes.add(theme)
                used_issues.add(issue)
            if len(selected) >= 3:
                break

    return pd.DataFrame(selected).reset_index(drop=True)


def detect_main_theme(row: pd.Series) -> str:
    cat = clean_text(row.get("category", ""))
    text = " ".join(clean_text(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "news_type"]).lower()

    if cat == "A_LEGAL_OFFICIAL":
        if contains_any(text, ["usmca", "원산지", "origin"]):
            return "공식 원산지/FTA"
        if contains_any(text, ["반덤핑", "anti-dumping", "countervailing", "상계관세"]):
            return "공식 AD/CVD"
        return "공식 정책/법령"
    if cat == "B_SEMICONDUCTOR_TARIFF":
        return "반도체 관세"
    if cat == "C_ORIGIN_FTA_USMCA":
        return "FTA/원산지"
    if cat == "D_EXPORT_CONTROL_SANCTION":
        return "수출통제"
    if cat == "E_AD_CVD_TRADE_REMEDY":
        return "반덤핑/상계관세"
    if cat == "F_CUSTOMS_AUDIT_VALUATION":
        return "통관/세관심사"
    if cat == "G_CBAM_SUPPLY_CHAIN":
        return "공급망/CBAM"
    if contains_any(text, ["301", "section 301"]):
        return "美 301조 관세"
    if contains_any(text, ["232", "section 232"]):
        return "美 232조 관세"
    if contains_any(text, ["semiconductor", "반도체"]):
        return "반도체 관세"
    if contains_any(text, ["export control", "수출통제", "ear"]):
        return "수출통제"
    if contains_any(text, ["fta", "원산지", "origin", "usmca"]):
        return "FTA/원산지"
    return "통상 리스크"


def detect_site_product(row: pd.Series) -> str:
    text = " ".join(clean_text(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Country"]).lower()
    sites = []
    if contains_any(text, ["vietnam", "베트남", "sev", "sevt"]):
        sites.append("SEV/SEVT")
    if contains_any(text, ["india", "인도", "siel"]):
        sites.append("SIEL")
    if contains_any(text, ["mexico", "멕시코", "samex"]):
        sites.append("SAMEX")
    if contains_any(text, ["china", "중국"]):
        sites.append("중국")
    if contains_any(text, ["united states", "u.s.", "usa", "미국"]):
        sites.append("북미 판매법인")
    if contains_any(text, ["eu", "europe", "유럽"]):
        sites.append("EU 판매법인")
    if not sites:
        sites.append("주요 생산거점")

    products = []
    if contains_any(text, SEMICON_TERMS):
        products.append("반도체")
    if contains_any(text, ["mobile", "smartphone", "스마트폰", "휴대폰"]):
        products.append("Mobile")
    if contains_any(text, ["consumer electronics", "가전", "tv"]):
        products.append("CE")
    if contains_any(text, ["display", "디스플레이"]):
        products.append("Display")
    if contains_any(text, ["component", "부품"]):
        products.append("Component")
    if not products:
        products.append("주요 제품")

    return f"{','.join(sites[:3])} {','.join(products[:3])}"


def executive_impact_summary(row: pd.Series) -> str:
    theme = detect_main_theme(row)
    target = detect_site_product(row)

    if theme == "반도체 관세":
        return f"반도체 관세 이슈 → {target}의 HS·관세율 시나리오, EAR·수출통제, 북미/EU 가격전가 영향 재산출"
    if theme in ["공식 원산지/FTA", "FTA/원산지"]:
        return f"FTA/원산지 변화 → {target}의 원산지 판정기준, BOM 충족률, CO 발급·특혜세율 적용 기준 재확정"
    if theme == "수출통제":
        return f"수출통제 강화 → {target}의 HS·ECCN·최종사용자 심사 및 출하통제 기준 업데이트"
    if theme == "공식 AD/CVD" or theme == "반덤핑/상계관세":
        return f"AD/CVD 조치 → {target}의 대상 HS·공급국 원산지·과세가격 및 대체 공급선 영향 재산출"
    if theme == "공식 정책/법령":
        return f"공식 정책·법령 변경 → {target}의 HS·원산지·FTA·수출통제 내부통제 기준 및 관세사 지침 개정"
    if theme == "통관/세관심사":
        return f"통관·세관심사 강화 → {target}의 HS·과세가격·원산지 신고 정확성 및 소명자료 패키지 정비"
    if theme == "공급망/CBAM":
        return f"공급망·CBAM 이슈 → {target}의 원재료 출처·원산지·탄소자료 증빙 체계 정렬"
    if theme == "美 301조 관세":
        return f"美 301조 관세 변동 → {target}의 대미 수출 관세부담·원산지·가격전가 영향 재산출"
    if theme == "美 232조 관세":
        return f"美 232조 관세 확대 → {target}의 원재료·부품 관세부담 및 북미 판매법인 원가 영향 점검"
    return f"{theme} → {target}의 관세·원산지·FTA·통관 내부통제 영향 점검"


def build_executive_total_line(top3: pd.DataFrame) -> str:
    themes = [detect_main_theme(r) for _, r in top3.iterrows()]
    theme_txt = "·".join(list(dict.fromkeys(themes))[:3])

    country_text = " ".join(clean_text(r.get("Country", "")) for _, r in top3.iterrows()).lower()
    areas = []
    if contains_any(country_text, ["미국", "us", "u.s.", "usa", "united states"]):
        areas.append("미국")
    if contains_any(country_text, ["멕시코", "mexico"]):
        areas.append("멕시코")
    if contains_any(country_text, ["eu", "유럽", "europe"]):
        areas.append("EU")
    if contains_any(country_text, ["중국", "china"]):
        areas.append("중국")
    if contains_any(country_text, ["베트남", "vietnam"]):
        areas.append("베트남")
    if contains_any(country_text, ["인도", "india"]):
        areas.append("인도")
    if not areas:
        areas = ["미국", "EU", "아시아"]

    return (
        f"{','.join(areas[:3])} 중심의 {theme_txt} 이슈가 삼성전자 주요 생산거점 및 판매법인의 "
        f"관세원가·원산지·FTA·수출통제·통관 내부통제 재점검 필요성을 높이고 있습니다."
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
            Publish Date: {html.escape(clean_text(r['Date']))} | Country: {html.escape(clean_text(r['Country']))} | Agency: {html.escape(clean_text(r['Agency']))} | Risk: {html.escape(clean_text(r['Risk']))} | Category: {html.escape(clean_text(r['category']))} | Score: {html.escape(score)}
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


def build_html(top30: pd.DataFrame, top3: pd.DataFrame) -> str:
    top3_nos = set(top3["No"].tolist())
    rest = top30[~top30["No"].isin(top3_nos)].copy()

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
    <p style="margin-top:18px;color:#666;font-size:12px;">※ 첨부 Excel 파일에 전체 Top30 분석표가 포함되어 있습니다.</p>
  </div>
</body>
</html>"""


def save_outputs(top30: pd.DataFrame, html_body: str) -> None:
    top30.to_excel(OUTPUT_XLSX, index=False)
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


def update_mail_cumulative(top30: pd.DataFrame) -> None:
    try:
        data = top30.copy()
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


def main() -> None:
    log("[START] GTI Mail Engine v12 - Respect STEP4 Structural Ranking")
    raw = read_input()
    log(f"[LOAD] {INPUT_FILE} rows={len(raw)}")

    norm = normalize_columns(raw)
    top30 = prepare_top30(norm)
    top3 = select_top3(top30)

    log(f"[SELECT] Top30 rows={len(top30)} / Top3 rows={len(top3)}")
    if len(top3) > 0:
        log("[TOP3]")
        for _, r in top3.iterrows():
            log(f"  - {r['category']} / {detect_main_theme(r)} / {r['Headline'][:80]}")

    html_body = build_html(top30, top3)
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    save_outputs(top30, html_body)
    update_mail_cumulative(top30)
    send_email(html_body)
    log("[DONE] GTI Mail Engine v12")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"[ERROR] {e}")
        log(traceback.format_exc())
        raise
