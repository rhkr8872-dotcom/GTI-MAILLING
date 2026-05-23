# -*- coding: utf-8 -*-
"""
GTI C_TYPE - Samsung HQ Customs Daily Sensing Mail Engine

Purpose
-------
Single-file final version that combines:
1. Step4-style news selection and clustering.
2. B_type executive mail layout.
3. A_type-style fallback writing quality for Summary / AI Analysis / Action.

Default input
-------------
    C:/Temp/3.news_ai_summary.xlsx

Default outputs
---------------
    C:/Temp/12345/c_type_outputs/4.news_ai_analysis.xlsx
    C:/Temp/12345/c_type_outputs/[GTI Radar] Global Trade Intelligence(YYYY-MM-DD).xlsx
    C:/Temp/12345/c_type_outputs/[GTI Radar] Global Trade Intelligence(YYYY-MM-DD).html

Environment variables
---------------------
    GTI_INPUT_FILE       default C:/Temp/3.news_ai_summary.xlsx
    GTI_OUTPUT_DIR       default C:/Temp/12345/c_type_outputs
    GTI_RUN_DATE         default today
    GTI_LOOKBACK_HOURS   default 24
    GTI_TOP_N            default 30
    GTI_USE_GEMINI       default Y
    GEMINI_API_KEY       optional
    GEMINI_MODEL         default gemini-1.5-flash
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# CONFIG
# =============================================================================
INPUT_FILE = Path(os.getenv("GTI_INPUT_FILE", r"C:\Temp\3.news_ai_summary.xlsx"))
OUTPUT_DIR = Path(os.getenv("GTI_OUTPUT_DIR", r"C:\Temp\12345\c_type_outputs"))
RUN_DATE = os.getenv("GTI_RUN_DATE", datetime.now().strftime("%Y-%m-%d"))
LOOKBACK_HOURS = int(os.getenv("GTI_LOOKBACK_HOURS", "24"))
TOP_N = int(os.getenv("GTI_TOP_N", "30"))

USE_GEMINI = os.getenv("GTI_USE_GEMINI", "Y").strip().upper() in {"Y", "YES", "TRUE", "1"}
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-1.5-flash").strip()
GEMINI_TEMPERATURE = float(os.getenv("GTI_GEMINI_TEMPERATURE", "0.2"))

SEND_EMAIL = os.getenv("GTI_SEND_EMAIL", "Y").strip().upper() in {"Y", "YES", "TRUE", "1"}
SMTP_HOST = os.getenv("GTI_SMTP_HOST", "smtp.naver.com")
SMTP_PORT = int(os.getenv("GTI_SMTP_PORT", "465"))
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or "3GKBVKMZMEKK").strip()
MAIL_TO = os.getenv("GTI_MAIL_TO", "").strip()
MAIL_FROM_NAME = os.getenv("GTI_MAIL_FROM_NAME", "GTI Radar").strip()
RECIPIENT_FILE = Path(os.getenv("GTI_RECIPIENT_FILE", r"C:\Temp\00.xlsx"))


def output_paths() -> dict[str, Path]:
    return {
        "analysis": OUTPUT_DIR / "4.news_ai_analysis.xlsx",
        "mail_xlsx": OUTPUT_DIR / f"[GTI Radar] Global Trade Intelligence({RUN_DATE}).xlsx",
        "mail_html": OUTPUT_DIR / f"[GTI Radar] Global Trade Intelligence({RUN_DATE}).html",
        "cumulative": OUTPUT_DIR / "gti_news_cumulative.xlsx",
    }


# =============================================================================
# KEYWORDS
# =============================================================================
RISK_ORDER = {"상": 1, "중": 2, "하": 3}

OFFICIAL_TERMS = [
    "ustr", "cbp", "federal register", "wto", "wco", "european commission", "eu commission",
    "mofcom", "gacc", "customs", "trade remedies authority", ".gov",
    "관세청", "관보", "법제처", "입법예고", "행정예고", "고시", "공고", "상무부", "무역대표부",
]

CORE_CUSTOMS_TERMS = [
    "tariff", "duty", "customs", "section 301", "section 232", "ustr", "cbp",
    "anti-dumping", "antidumping", "countervailing", "safeguard", "export control",
    "import restriction", "origin", "country of origin", "hs code", "classification",
    "customs valuation", "valuation", "fta", "cepa", "usmca", "wto", "cbam", "refund",
    "관세", "관세율", "추가관세", "상호관세", "301조", "232조", "반덤핑", "상계관세",
    "세이프가드", "수입규제", "수출통제", "원산지", "품목분류", "HS", "과세가격",
    "통관", "관세환급", "FTA", "무역협정", "관세소송",
]

HIGH_VALUE_TERMS = [
    "section 301", "section 232", "export control", "anti-dumping", "antidumping",
    "countervailing", "tariff refund", "trade court", "cit", "forced labor", "uflpa",
    "critical minerals", "rare earth", "semiconductor", "chip", "hbm",
    "301조", "232조", "수출통제", "반덤핑", "상계관세", "관세환급", "희토류", "핵심광물",
    "반도체", "칩", "원산지 규정",
]

SAMSUNG_PRODUCT_TERMS = [
    "samsung", "semiconductor", "chip", "hbm", "memory", "display", "battery",
    "smartphone", "mobile", "electronics", "consumer electronics", "appliance",
    "network equipment", "telecom", "5g", "component", "pcb", "mlcc",
    "삼성", "삼성전자", "반도체", "메모리", "디스플레이", "배터리", "스마트폰",
    "모바일", "전자", "가전", "네트워크", "통신장비", "부품",
]

SAMSUNG_SITE_TERMS = [
    "vietnam", "india", "mexico", "china", "korea", "poland", "slovakia", "turkey",
    "brazil", "indonesia", "united states", "eu", "europe", "sev", "sevt", "siel", "samex",
    "베트남", "인도", "멕시코", "중국", "한국", "폴란드", "슬로바키아", "튀르키예",
    "브라질", "인도네시아", "미국", "유럽", "EU",
]

NOISE_TERMS = [
    # entertainment / sport / market articles
    "sports", "football", "baseball", "celebrity", "movie", "drama", "concert",
    "stock price", "stock market", "shares", "earnings", "bitcoin", "crypto", "benzinga",
    "yahoo finance", "investing.com", "bitget", "coinness",
    "스포츠", "축구", "야구", "연예", "배우", "영화", "드라마", "주가", "증시", "급등", "급락", "코인",
    # food, retail, tourism, small business
    "coffee", "beef", "cattle", "seafood", "fertilizer", "farmers", "duty free", "retail",
    "fertiliser", "gold", "silver", "precious metals", "shopper", "shoppers",
    "chili powder", "online shopping", "market share", "electric vehicle", "tire", "tyre",
    "startup", "start-up", "mou with", "workforce skills", "logistics centre",
    "ag exporters", "farm policy", "cocoa", "fruit", "metallurgical coke", "solar panel", "solar panels",
    "steel", "hot-rolled", "special steel", "prestress", "posco", "animals", "animal", "ebola",
    "travel restrictions", "flight diverted",
    "cosmetic", "fashion", "apparel", "footwear", "tourism", "character goods", "popup",
    "커피", "소고기", "쇠고기", "수산", "비료", "농산물", "금 수요", "은 수입", "귀금속",
    "고추", "온라인 쇼핑", "시장 점유율", "전기차", "타이어", "스타트업",
    "철강", "열연강판", "특수강", "강선", "강케이블", "포스코", "현대제철", "보잉",
    "전문가 배출", "인재 양성",
    "면세점", "롯데면세점", "화장품",
    "패션", "의류", "신발", "관광", "캐릭터", "팝업", "농민",
    # crime / personal customs / procurement noise
    "fentanyl", "drug", "narcotic", "smuggling", "stolen", "weapon", "firearm", "suppressor",
    "cannabis", "cannabis resin",
    "child porn", "former u.s. customs", "sentenced for", "porn",
    "nissan patrol", "lucky baskhar", "dutquer", "dulquer", "personal baggage",
    "how much gold can you bring", "duty-free limits", "tender", "procurement", "security equipment",
    "sanctuary city", "sanctuary cities", "immigration", "airport staffing", "customs staffing",
    "telcos", "unregistered devices", "russian influence", "nigeria", "moldovan", "cameroon",
    "boeing", "枇杷",
    "마약", "밀수", "도난", "무기", "총기", "휴대품", "면세한도", "입찰", "조달", "보안검색 장비",
    # company/consumer disputes that are not Samsung customs operating signals
    "amazon faces", "class-action lawsuit", "walmart to use tariff refunds", "lawsuit over alleged retention",
]

TOP3_STRONG_NOISE = [
    "amazon faces lawsuit", "class-action lawsuit", "coffee", "fertilizer", "fertiliser", "gold demand",
    "silver imports", "duty-free limits", "retail", "tourism", "면세점", "커피", "비료", "금 수요",
    "은 수입", "롯데", "관광",
]


# =============================================================================
# BASIC HELPERS
# =============================================================================
def log(msg: str) -> None:
    print(msg, flush=True)


def clean(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = html.unescape(str(value)).replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return "" if text.lower() in {"nan", "none", "nat", "0"} else text


def contains_any(text: str, terms: list[str]) -> bool:
    low = text.lower()
    return any(term.lower() in low for term in terms)


def count_hits(text: str, terms: list[str]) -> int:
    low = text.lower()
    return sum(1 for term in terms if term.lower() in low)


def normalize_title(title: str) -> str:
    text = clean(title).lower()
    text = re.sub(r"\s*[-|–]\s*[^-|–]{2,50}$", "", text)
    text = re.sub(r"[^0-9a-z가-힣一-龥 ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def similar_title(a: str, b: str) -> bool:
    if not a or not b:
        return False
    if a[:34] == b[:34]:
        return True
    return SequenceMatcher(None, a, b).ratio() >= 0.84


def unwrap_url(url: str) -> str:
    value = clean(url)
    if not value:
        return ""
    try:
        parsed = urlparse(value)
        query = parse_qs(parsed.query)
        for key in ("url", "u"):
            if query.get(key):
                return unquote(query[key][0])
    except Exception:
        pass
    return value


def parse_date(value):
    if isinstance(value, datetime):
        return value
    text = clean(value)
    if not text:
        return None
    try:
        dt = pd.to_datetime(text, errors="coerce")
        if not pd.isna(dt):
            return dt.to_pydatetime()
    except Exception:
        return None
    return None


def is_recent(value) -> bool:
    dt = parse_date(value)
    if dt is None:
        return True
    return dt >= datetime.now() - timedelta(hours=LOOKBACK_HOURS)


def cut_sentence(text: str, limit: int = 120) -> str:
    text = clean(text)
    if len(text) <= limit:
        return text
    return text[:limit].rstrip() + "..."


# =============================================================================
# INPUT NORMALIZATION
# =============================================================================
def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for key in candidates:
        if key.lower() in lower_map:
            return lower_map[key.lower()]
    for col in df.columns:
        low = str(col).strip().lower()
        if any(key.lower() in low for key in candidates):
            return col
    return None


def read_input() -> pd.DataFrame:
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"입력 파일 없음: {INPUT_FILE}")
    df = pd.read_excel(INPUT_FILE)
    df.columns = [str(c).strip() for c in df.columns]

    col_date = pick_col(df, ["date", "publish date", "published"])
    col_title = pick_col(df, ["title", "headline"])
    col_url = pick_col(df, ["url", "link"])
    col_source = pick_col(df, ["source_file", "source", "publisher"])
    col_keyword = pick_col(df, ["keyword", "category"])
    col_score = pick_col(df, ["score", "importance"])
    col_agency = pick_col(df, ["agency"])

    out = pd.DataFrame()
    out["DateRaw"] = df[col_date] if col_date else ""
    out["Date"] = [format_date(v) for v in out["DateRaw"]]
    out["Headline"] = df[col_title].apply(clean) if col_title else ""
    out["URL"] = df[col_url].apply(unwrap_url) if col_url else ""
    out["Source"] = df[col_source].apply(clean) if col_source else ""
    out["Keyword"] = df[col_keyword].apply(clean) if col_keyword else ""
    out["AgencyRaw"] = df[col_agency].apply(clean) if col_agency else ""
    out["Step3Score"] = df[col_score] if col_score else 0
    out = out[(out["Headline"] != "") & (out["URL"] != "")]
    out = out[out["DateRaw"].apply(is_recent)]
    return out.reset_index(drop=True)


def format_date(value) -> str:
    dt = parse_date(value)
    if dt is None:
        return clean(value)[:16]
    return dt.strftime("%Y-%m-%d %H:%M")


# =============================================================================
# STEP4-STYLE DECISION AND SCORING
# =============================================================================
def row_text(row: pd.Series) -> str:
    return " ".join(clean(row.get(c, "")) for c in ["Headline", "Source", "Keyword", "AgencyRaw"])


def is_noise(row: pd.Series) -> bool:
    text = row_text(row)
    if contains_any(text, NOISE_TERMS):
        non_overridable = [
            "gold", "silver", "coffee", "fertilizer", "fertiliser", "amazon", "walmart",
            "sanctuary city", "immigration", "airport", "telcos", "unregistered devices",
            "cannabis", "chili powder", "online shopping", "market share", "tire", "tyre",
            "startup", "start-up", "workforce skills", "logistics centre", "ag exporters",
            "farm policy", "cocoa", "fruit", "metallurgical coke", "solar panel", "child porn",
            "former u.s. customs", "boeing", "steel", "prestress", "animals", "ebola",
            "travel restrictions", "flight diverted", "枇杷",
            "금 수요", "은 수입", "커피", "비료", "아마존", "월마트",
            "고추", "온라인 쇼핑", "시장 점유율", "타이어", "스타트업",
            "철강", "열연강판", "특수강", "강선", "강케이블", "포스코", "현대제철", "보잉",
            "전문가 배출", "인재 양성",
        ]
        if contains_any(text, non_overridable):
            return True
        if contains_any(text, ["section 301", "301조", "ustr", "cbp", "federal register", "반덤핑", "상계관세", "수출통제", "semiconductor", "반도체"]):
            return False
        return True
    return False


def infer_country(text: str) -> str:
    rules = [
        ("United States", ["united states", "u.s.", "usa", "us ", "미국", "美", "ustr", "cbp", "trump"]),
        ("China", ["china", "중국", "中", "mofcom", "gacc", "beijing", "xi", "시진핑"]),
        ("EU", ["eu", "europe", "european", "유럽", "eu commission"]),
        ("Vietnam", ["vietnam", "베트남", "sev", "sevt"]),
        ("India", ["india", "인도", "siel"]),
        ("Mexico", ["mexico", "멕시코", "samex", "usmca"]),
        ("Korea", ["korea", "한국", "관세청", "산업통상", "무역위"]),
        ("Japan", ["japan", "일본"]),
        ("Brazil", ["brazil", "브라질"]),
        ("Indonesia", ["indonesia", "인도네시아"]),
    ]
    found = []
    low = text.lower()
    for country, keys in rules:
        if any(k.lower() in low for k in keys) and country not in found:
            found.append(country)
    return ", ".join(found[:2]) if found else "Global"


def infer_agency(text: str) -> str:
    rules = [
        ("USTR", ["ustr", "u.s. trade representative", "무역대표"]),
        ("U.S. Customs and Border Protection (CBP)", ["cbp", "customs and border protection"]),
        ("U.S. Department of Commerce", ["department of commerce", "u.s. commerce"]),
        ("European Commission", ["european commission", "eu commission"]),
        ("WTO", ["wto"]),
        ("MOFCOM", ["mofcom", "중국 상무부"]),
        ("Korea Customs Service", ["관세청"]),
        ("Korea Trade Commission", ["무역위", "무역위원회"]),
        ("Vietnam Customs / Trade Remedies Authority", ["vietnam customs", "vietnam.vn", "trade remedies authority"]),
    ]
    low = text.lower()
    for agency, keys in rules:
        if any(k.lower() in low for k in keys):
            return agency
    return "Relevant customs/trade authority"


def classify_issue(text: str) -> str:
    low = text.lower()
    if contains_any(low, ["section 301", "301조"]):
        return "SECTION_301"
    if contains_any(low, ["section 232", "232조"]):
        return "SECTION_232"
    if contains_any(low, ["export control", "수출통제", "sanction", "제재"]):
        return "EXPORT_CONTROL"
    if contains_any(low, ["anti-dumping", "antidumping", "반덤핑", "countervailing", "상계관세"]):
        return "AD_CVD"
    if contains_any(low, ["origin", "원산지", "usmca", "fta", "cepa"]):
        return "FTA_ORIGIN"
    if contains_any(low, ["customs", "통관", "cbp", "관세청", "valuation", "과세가격"]):
        return "CUSTOMS"
    if contains_any(low, ["semiconductor", "반도체", "chip", "hbm", "rare earth", "희토류"]):
        return "SEMICONDUCTOR_SUPPLY"
    if contains_any(low, ["tariff", "관세", "duty"]):
        return "TARIFF"
    return "GENERAL_TRADE"


def cluster_key(row: pd.Series) -> str:
    text = row_text(row).lower()
    if (
        contains_any(text, ["chip tariff", "chip tariffs", "semiconductor tariff", "semiconductor tariffs", "반도체 관세"])
        or (contains_any(text, ["ustr", "greer", "무역대표"]) and contains_any(text, ["chip", "chips", "semiconductor", "반도체", "arancel"]))
    ):
        return "CHIP_TARIFF"
    if contains_any(text, ["usmca", "원산지 규정", "rules of origin"]):
        return "USMCA_ORIGIN"
    if contains_any(text, ["trade court", "section 122", "tariff refund", "관세환급", "관세소송"]):
        return "US_TARIFF_LITIGATION_REFUND"
    issue = classify_issue(text)
    country = infer_country(text).split(",")[0]
    if issue in {"SECTION_301", "SECTION_232", "EXPORT_CONTROL", "AD_CVD", "FTA_ORIGIN", "CUSTOMS", "SEMICONDUCTOR_SUPPLY"}:
        return f"{issue}_{country}"
    title = normalize_title(row.get("Headline", ""))
    return f"{issue}_{title[:40]}"


def base_score(row: pd.Series) -> int:
    text = row_text(row)
    if is_noise(row):
        return -999

    score = 0
    official_hits = count_hits(text, OFFICIAL_TERMS)
    core_hits = count_hits(text, CORE_CUSTOMS_TERMS)
    high_hits = count_hits(text, HIGH_VALUE_TERMS)
    product_hits = count_hits(text, SAMSUNG_PRODUCT_TERMS)
    site_hits = count_hits(text, SAMSUNG_SITE_TERMS)
    if contains_any(text, ["steel", "철강", "강선", "강케이블", "prestress"]) and not contains_any(text, ["semiconductor", "chip", "hbm", "반도체", "전자", "electronics"]):
        score -= 700

    score += min(official_hits, 4) * 350
    score += min(core_hits, 8) * 260
    score += min(high_hits, 6) * 420
    score += min(product_hits, 5) * 260
    score += min(site_hits, 5) * 160

    if core_hits and high_hits:
        score += 900
    if product_hits and core_hits:
        score += 700
    if site_hits and core_hits:
        score += 500
    if product_hits and site_hits:
        score += 400

    issue = classify_issue(text)
    priority_bonus = {
        "SECTION_301": 1500,
        "SECTION_232": 1300,
        "EXPORT_CONTROL": 1450,
        "AD_CVD": 1350,
        "FTA_ORIGIN": 1050,
        "CUSTOMS": 950,
        "SEMICONDUCTOR_SUPPLY": 1250,
        "TARIFF": 800,
        "GENERAL_TRADE": 200,
    }
    score += priority_bonus.get(issue, 0)

    try:
        score += min(int(float(row.get("Step3Score", 0) or 0)), 50)
    except Exception:
        pass

    return int(score)


def infer_risk(score: int, text: str) -> str:
    if score >= 4200 or contains_any(text, ["section 301", "301조", "section 232", "수출통제", "anti-dumping", "반덤핑", "상계관세"]):
        return "상"
    if score >= 2300 or contains_any(text, ["관세", "통관", "원산지", "fta", "customs", "tariff", "origin"]):
        return "중"
    return "하"


def top3_eligible(row: pd.Series) -> bool:
    text = row_text(row)
    if contains_any(text, TOP3_STRONG_NOISE):
        return False
    if is_noise(row):
        return False
    issue = classify_issue(text)
    product_or_site = contains_any(text, SAMSUNG_PRODUCT_TERMS) or contains_any(text, ["vietnam", "india", "mexico", "china", "eu", "베트남", "인도", "멕시코", "중국", "미국", "유럽"])
    if issue == "AD_CVD" and not contains_any(text, SAMSUNG_PRODUCT_TERMS):
        return False
    if issue in {"TARIFF", "CUSTOMS", "FTA_ORIGIN"} and not product_or_site and not contains_any(text, ["ustr", "cbp", "federal register", "관세청", "wto"]):
        return False
    return issue in {"SECTION_301", "SECTION_232", "EXPORT_CONTROL", "AD_CVD", "FTA_ORIGIN", "CUSTOMS", "SEMICONDUCTOR_SUPPLY", "TARIFF"}


def select_news(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df["_text"] = df.apply(row_text, axis=1)
    df["_score"] = df.apply(base_score, axis=1)
    df["_noise"] = df.apply(is_noise, axis=1)
    df["_cluster"] = df.apply(cluster_key, axis=1)
    df["_title_norm"] = df["Headline"].apply(normalize_title)
    df["_url_norm"] = df["URL"].str.lower().str.strip()

    before = len(df)
    df = df[(df["_score"] > 0) & (~df["_noise"])].copy()
    log(f"[FILTER] input={before}, selected_candidates={len(df)}, removed={before - len(df)}")
    if df.empty:
        raise RuntimeError("선정 후보가 없습니다. 입력 파일 또는 필터 조건을 확인해 주세요.")

    df = df.sort_values(["_score", "Date"], ascending=[False, False])
    df = df.drop_duplicates(subset=["_url_norm"], keep="first")

    unique_rows = []
    seen_titles = []
    for _, row in df.iterrows():
        title_norm = row["_title_norm"]
        if any(similar_title(title_norm, old) for old in seen_titles):
            continue
        seen_titles.append(title_norm)
        unique_rows.append(row)
    df = pd.DataFrame(unique_rows)

    # Top3 first: high impact with topic diversity.
    top3 = []
    used_clusters = set()
    for _, row in df.iterrows():
        if not top3_eligible(row):
            continue
        key = row["_cluster"]
        if key in used_clusters:
            continue
        top3.append(row)
        used_clusters.add(key)
        if len(top3) == 3:
            break
    for _, row in df.iterrows():
        if len(top3) == 3:
            break
        if any(row["URL"] == picked["URL"] for picked in top3):
            continue
        top3.append(row)

    selected = list(top3)
    selected_urls = {r["URL"] for r in selected}
    cluster_count = {}
    for r in selected:
        cluster_count[r["_cluster"]] = cluster_count.get(r["_cluster"], 0) + 1

    for _, row in df.iterrows():
        if row["URL"] in selected_urls:
            continue
        cluster = row["_cluster"]
        if cluster_count.get(cluster, 0) >= 1:
            continue
        selected.append(row)
        selected_urls.add(row["URL"])
        cluster_count[cluster] = cluster_count.get(cluster, 0) + 1
        if len(selected) >= TOP_N:
            break

    # If strict cluster cap produces fewer than TOP_N, fill remaining by score.
    if len(selected) < TOP_N:
        for _, row in df.iterrows():
            if row["URL"] in selected_urls:
                continue
            selected.append(row)
            selected_urls.add(row["URL"])
            if len(selected) >= TOP_N:
                break

    out = pd.DataFrame(selected).reset_index(drop=True)
    rows = []
    for i, row in out.iterrows():
        text = row["_text"]
        score = int(row["_score"])
        country = infer_country(text)
        agency = clean(row.get("AgencyRaw", "")) or infer_agency(text)
        risk = infer_risk(score, text)
        products = infer_products(text)
        summary = fallback_summary(row["Headline"], country, agency)
        analysis = fallback_analysis(row["Headline"], country, risk, products)
        action = fallback_action(row["Headline"], country, agency, risk)

        item = {
            "No": i + 1,
            "Date": row["Date"],
            "Headline": row["Headline"],
            "Summary": summary,
            "AI Analysis": analysis,
            "Action Plan": action,
            "Country": country,
            "Agency": agency,
            "Risk": risk,
            "Samsung Impact Score": score,
            "URL": row["URL"],
            "Source": row["Source"],
            "Issue": classify_issue(text),
            "Cluster": row["_cluster"],
        }
        ai = analyze_with_gemini(item)
        if ai:
            item["Summary"] = clean(ai.get("Summary")) or item["Summary"]
            item["AI Analysis"] = clean(ai.get("AI Analysis")) or item["AI Analysis"]
            item["Action Plan"] = clean(ai.get("Action Plan")) or item["Action Plan"]
        rows.append(item)

    return pd.DataFrame(rows)


def infer_products(text: str) -> str:
    products = []
    rules = [
        ("Semiconductor/Component", ["semiconductor", "chip", "hbm", "memory", "반도체", "칩", "메모리", "부품"]),
        ("Mobile", ["smartphone", "mobile", "phone", "스마트폰", "모바일"]),
        ("Consumer Electronics", ["consumer electronics", "appliance", "tv", "가전", "tv"]),
        ("Display", ["display", "디스플레이"]),
        ("Network Equipment", ["network", "telecom", "5g", "네트워크", "통신장비"]),
    ]
    low = text.lower()
    for name, keys in rules:
        if any(k.lower() in low for k in keys):
            products.append(name)
    return ", ".join(products[:3]) if products else "Mobile, Consumer Electronics, Semiconductor/Component"


# =============================================================================
# WRITING QUALITY
# =============================================================================
def fallback_summary(headline: str, country: str, agency: str) -> str:
    return (
        f"{headline} 관련 정책·집행 변화입니다. 본사 관세팀 관점에서는 {country}의 {agency} 발표가 "
        "관세율, 통관요건, 원산지 또는 수입규제 변화로 연결되는지 확인할 필요가 있습니다."
    )


def fallback_analysis(headline: str, country: str, risk: str, products: str) -> str:
    low = headline.lower()
    if contains_any(low, ["section 301", "301조", "tariff", "관세"]):
        return (
            f"{country} 관세 변화는 {products}의 대미·대EU 수출입 원가와 가격전가 판단에 영향을 줄 수 있습니다. "
            "SEV/SEVT, SIEL, SAMEX 등 생산법인별 HS와 원산지 기준을 함께 점검해야 합니다."
        )
    if contains_any(low, ["export control", "수출통제", "sanction", "제재"]):
        return (
            f"{country} 수출통제 이슈는 {products}의 출하 가능 여부, 최종사용자 확인, 라이선스 필요성에 영향을 줄 수 있습니다. "
            "반도체·부품 공급망과 해외 생산법인 거래선을 우선 점검해야 합니다."
        )
    if contains_any(low, ["anti-dumping", "antidumping", "반덤핑", "countervailing", "상계관세"]):
        return (
            f"{country} 반덤핑·상계관세 이슈는 우회수출, 공급국 전환, 과세가격 방어자료 요구로 확대될 수 있습니다. "
            f"{products} 관련 원재료·부품 조달 구조와 거래가격 증빙을 확인해야 합니다."
        )
    if contains_any(low, ["origin", "원산지", "fta", "usmca"]):
        return (
            f"{country} 원산지·FTA 이슈는 특혜관세 적용과 비특혜 원산지 판정에 직접 영향을 줄 수 있습니다. "
            "생산공정, BOM, 원산지증명서 발급 기준을 생산지별로 재점검해야 합니다."
        )
    return (
        f"{country} 통상정책 변화는 {products}의 수입원가, 수출통관, 원산지 입증 부담에 영향을 줄 수 있습니다. "
        "본사 관세팀 모니터링 항목으로 등록해 시행일과 적용 품목을 확인해야 합니다."
    )


def fallback_action(headline: str, country: str, agency: str, risk: str) -> str:
    base = (
        f"{agency} 원문, 시행일, 적용 품목과 HS 범위를 확인하고 관련 법인에 영향 가능 품목 리스트를 요청합니다. "
        "원산지 증빙, FTA 적용 여부, 과세가격 자료를 함께 점검합니다."
    )
    if risk == "상":
        return base + " 관세율 변경 시나리오별 원가 영향을 즉시 산출해 사업부와 공유합니다."
    if risk == "중":
        return base + " 조사·협의·입법 진행 일정을 추적하고 요청 가능 자료를 사전에 준비합니다."
    return base + " 단기 조치보다 모니터링 리스트에 등록해 후속 공지를 확인합니다."


def analyze_with_gemini(row: dict) -> dict | None:
    if not USE_GEMINI or not GEMINI_API_KEY:
        return None
    try:
        from google import genai

        client = genai.Client(api_key=GEMINI_API_KEY)
        prompt = f"""
You are a Samsung Electronics HQ customs and trade compliance manager.

Samsung context:
- Production sites: Korea, China, Vietnam(SEV/SEVT), India(SIEL), Mexico(SAMEX), Indonesia, Poland, Slovakia, Turkey, Brazil.
- Products: Mobile smartphones, Consumer Electronics, Network Equipment, Semiconductor/Component, Display, Medical.
- Job scope: tariff rate, Section 301/232, anti-dumping/countervailing duties, export control, origin, HS classification, FTA, customs valuation, import/export clearance.

News:
Headline: {row['Headline']}
Country: {row['Country']}
Agency: {row['Agency']}
Risk: {row['Risk']}
Issue: {row['Issue']}

Return only JSON in Korean:
{{
  "Summary": "뉴스 요약 2문장 이내",
  "AI Analysis": "삼성전자 본사 관세담당자 관점 영향 2문장 이내",
  "Action Plan": "관세담당자가 실행할 조치 1-2문장"
}}
"""
        res = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={"temperature": GEMINI_TEMPERATURE, "max_output_tokens": 900},
        )
        text = res.text.strip()
        return json.loads(text[text.find("{") : text.rfind("}") + 1])
    except Exception:
        return None


def build_gemini_client():
    if not USE_GEMINI or not GEMINI_API_KEY:
        return None
    try:
        from google import genai

        return genai.Client(api_key=GEMINI_API_KEY)
    except Exception:
        return None


def parse_json_object(text: str) -> dict | None:
    try:
        text = clean(text)
        start = text.find("{")
        end = text.rfind("}")
        if start < 0 or end < start:
            return None
        return json.loads(text[start : end + 1])
    except Exception:
        return None


def optimize_rows_with_gemini(rows: pd.DataFrame) -> pd.DataFrame:
    """
    Gemini is used only after deterministic selection is finished.
    It must not add/remove/re-rank news. It only improves wording.
    """
    client = build_gemini_client()
    if client is None:
        log("[GEMINI] skipped: API key/client unavailable")
        return rows

    improved = rows.copy()
    for idx, row in improved.iterrows():
        prompt = f"""
You are Samsung Electronics HQ customs and trade compliance manager.

Rewrite the selected news analysis in Korean for a daily executive customs/trade intelligence email.

Rules:
- Do not change the headline, country, agency, risk, score, issue, or URL.
- Do not exaggerate Samsung impact. If impact is indirect, say it is indirect.
- Use Samsung customs wording: SEV/SEVT, SIEL, SAMEX, HS, origin, customs valuation, FTA, ECCN, tariff rate, import/export clearance.
- Summary: max 2 concise sentences.
- AI Analysis: max 2 concise sentences, Samsung HQ customs perspective.
- Action Plan: max 2 concrete action sentences.
- Avoid repeated boilerplate.

News:
Headline: {row['Headline']}
Country: {row['Country']}
Agency: {row['Agency']}
Risk: {row['Risk']}
Issue: {row['Issue']}
Current Summary: {row['Summary']}
Current AI Analysis: {row['AI Analysis']}
Current Action Plan: {row['Action Plan']}

Return only JSON:
{{
  "Summary": "...",
  "AI Analysis": "...",
  "Action Plan": "..."
}}
"""
        try:
            res = client.models.generate_content(
                model=GEMINI_MODEL,
                contents=prompt,
                config={"temperature": GEMINI_TEMPERATURE, "max_output_tokens": 900},
            )
            data = parse_json_object(res.text)
            if data:
                for col in ["Summary", "AI Analysis", "Action Plan"]:
                    value = clean(data.get(col, ""))
                    if value:
                        improved.at[idx, col] = value
        except Exception as exc:
            log(f"[GEMINI] row {idx + 1} skipped: {type(exc).__name__}")
    return improved


def optimize_review_with_gemini(top3: pd.DataFrame) -> dict | None:
    client = build_gemini_client()
    if client is None or top3.empty:
        return None

    items = []
    for i, (_, row) in enumerate(top3.iterrows(), start=1):
        items.append({
            "rank": i,
            "headline": clean(row.get("Headline", "")),
            "country": clean(row.get("Country", "")),
            "agency": clean(row.get("Agency", "")),
            "risk": clean(row.get("Risk", "")),
            "issue": clean(row.get("Issue", "")),
            "summary": clean(row.get("Summary", "")),
            "analysis": clean(row.get("AI Analysis", "")),
            "action": clean(row.get("Action Plan", "")),
        })

    prompt = f"""
You are writing the opening executive review for Samsung Electronics HQ customs team.

Input is the already-selected Top3. Do not question or change selection.

Write in Korean:
1. total_review: one concise paragraph, 1-2 sentences.
2. bullets: exactly 3 bullets, one for each Top3, each one sentence.

Rules:
- Perspective: Samsung Electronics HQ customs/trade compliance manager.
- Mention SEV/SEVT, SIEL, SAMEX only when relevant; otherwise use "주요 생산법인".
- Use concrete customs terms: HS, origin, customs valuation, FTA, ECCN, tariff rate, import/export clearance.
- Do not mention FTA unless the issue is explicitly FTA/origin/trade agreement.
- Do not say Samsung should consult foreign customs authorities unless the news is an official customs procedure requiring consultation.
- Do not invent plant-specific impact. If uncertain, say "주요 생산법인" or "대미/대EU 수출 품목".
- For USTR semiconductor tariff news, focus on HS mapping, tariff-rate scenarios, origin, and US-bound shipment impact.
- For US-China semiconductor export-control news, focus on ECCN, end-user, re-export, and shipment-control checks.
- For EU-China semiconductor supply news, focus on alternative sourcing, import restriction changes, origin, and customs clearance monitoring.
- Do not repeat the same sentence pattern.
- No vague phrase like "영향 점검 필요" repeated across bullets.
- Do not mention unrelated industries unless needed to explain indirect risk.
- Keep each bullet under 85 Korean characters if possible.

Top3 JSON:
{json.dumps(items, ensure_ascii=False)}

Return only JSON:
{{
  "total_review": "...",
  "bullets": ["...", "...", "..."]
}}
"""
    try:
        res = client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config={"temperature": GEMINI_TEMPERATURE, "max_output_tokens": 1000},
        )
        data = parse_json_object(res.text)
        if not data:
            return None
        bullets = data.get("bullets", [])
        if not isinstance(bullets, list) or len(bullets) != 3:
            return None
        total = clean(data.get("total_review", ""))
        bullets = [clean(x) for x in bullets if clean(x)]
        if not total or len(bullets) != 3:
            return None
        bullets = [postprocess_review_bullet(top3.iloc[i], bullets[i]) for i in range(3)]
        total = postprocess_total_review(total)
        return {"total_review": total, "bullets": bullets}
    except Exception as exc:
        log(f"[GEMINI] review skipped: {type(exc).__name__}")
        return None


def postprocess_total_review(text: str) -> str:
    text = clean(text)
    text = text.replace("지속적인 모니터링 및 대응이 필요합니다", "HS·원산지·수출통제 기준을 우선 재점검해야 합니다")
    text = text.replace("지속적인 모니터링이 필요합니다", "후속 시행일과 적용 품목을 추적해야 합니다")
    return text


def postprocess_review_bullet(row: pd.Series, text: str) -> str:
    text = clean(text)
    issue = clean(row.get("Issue", ""))
    headline = clean(row.get("Headline", "")).lower()

    if issue not in {"FTA_ORIGIN"} and not contains_any(headline, ["fta", "origin", "원산지", "무역협정", "usmca"]):
        text = text.replace("FTA 활용 전략", "원산지·관세율 영향")
        text = text.replace("FTA 활용 가능성", "원산지·관세율 영향")
        text = text.replace("FTA 활용", "원산지 기준")

    text = text.replace("EU 관세 당국과의 협의를 통해", "EU 집행위·관세당국 후속 공지를 확인해")
    text = text.replace("관세 당국과의 협의를 통해", "관세당국 후속 공지를 확인해")
    text = text.replace("통관 관련 불확실성을 해소해야 합니다", "통관 리스크를 사전에 정리해야 합니다")

    if issue == "SEMICONDUCTOR_SUPPLY" and contains_any(headline, ["ustr", "chip tariff", "반도체 관세", "tariff"]):
        return "미국 반도체 관세가 즉시 시행되지 않더라도 HS별 관세율 시나리오와 원산지 영향표를 선제 갱신해야 합니다."
    if issue == "EXPORT_CONTROL":
        return "미·중 수출통제 이슈는 반도체 부품의 ECCN, 최종사용자, 우회수출 가능성을 재확인해야 합니다."
    if issue == "SEMICONDUCTOR_SUPPLY" and contains_any(headline, ["eu", "중국", "china", "반도체"]):
        return "중국/EU 반도체 공급망 변화는 대체조달, 원산지 판정, 수입규제 리스크로 나누어 봐야 합니다."

    return text


# =============================================================================
# EXECUTIVE HTML
# =============================================================================
def html_link(headline: str, url: str) -> str:
    title = html.escape(clean(headline))
    href = html.escape(clean(url))
    if href.startswith("http"):
        return f'<a href="{href}" style="color:#0563C1;text-decoration:underline;font-weight:bold;">{title}</a>'
    return f"<b>{title}</b>"


def risk_color(risk: str) -> str:
    return {"상": "#C00000", "중": "#ED7D31", "하": "#5B9BD5"}.get(risk, "#666666")


def detect_theme(row: pd.Series) -> str:
    issue = clean(row.get("Issue", ""))
    return {
        "SECTION_301": "미 301조 관세",
        "SECTION_232": "미 232조 관세",
        "EXPORT_CONTROL": "수출통제",
        "AD_CVD": "반덤핑·상계관세",
        "FTA_ORIGIN": "FTA·원산지",
        "CUSTOMS": "통관·관세집행",
        "SEMICONDUCTOR_SUPPLY": "반도체 공급망",
        "TARIFF": "관세율 변동",
    }.get(issue, "통상정책")


def executive_impact(row: pd.Series) -> str:
    theme = detect_theme(row)
    country = clean(row.get("Country", "관련국"))
    text = " ".join(clean(row.get(c, "")) for c in ["Headline", "Summary", "AI Analysis", "Action Plan", "Country"])
    products = infer_products(text)
    low = text.lower()
    site = "주요 생산법인"
    if contains_any(low, ["vietnam", "베트남", "sev", "sevt"]):
        site = "SEV/SEVT"
    elif contains_any(low, ["india", "인도", "siel"]):
        site = "SIEL"
    elif contains_any(low, ["mexico", "멕시코", "samex"]):
        site = "SAMEX"
    elif contains_any(low, ["china", "중국"]):
        site = "중국 생산·조달망"

    if "301조" in theme or "232조" in theme or "관세" in theme:
        return f"{theme}: {site} 기준 {products}의 HS별 관세율·원산지·가격전가 영향표를 갱신해야 합니다."
    if "수출통제" in theme:
        return f"{theme}: {products}의 ECCN, 최종사용자, 우회수출 가능성을 점검하고 출하통제 기준을 재확인해야 합니다."
    if "반덤핑" in theme:
        return f"{theme}: 대상 품목이 삼성 부품·설비 조달망과 겹치는지 확인하고 원산지·가격 방어자료를 준비해야 합니다."
    if "FTA" in theme:
        return f"{theme}: {site}의 원산지 판정, CO 발급 요건, 특혜세율 적용 가능성을 재검토해야 합니다."
    if "반도체" in theme:
        if contains_any(country, ["United States", "미국"]):
            return f"{theme}: 미국 반도체 관세가 즉시 시행되지 않더라도 품목별 HS·원산지 영향표를 선제 갱신해야 합니다."
        if contains_any(country, ["China", "EU", "중국", "유럽"]):
            return f"{theme}: 중국/EU 반도체 공급 의존도 변화가 대체조달·원산지 판정·수입규제 리스크로 이어지는지 봐야 합니다."
        if contains_any(country, ["Korea", "한국"]):
            return f"{theme}: 한국 수출 쏠림이 커진 만큼 반도체 주요 품목의 관세율·통관 리스크를 별도 관리해야 합니다."
        return f"{theme}: 반도체 부품·장비 조달선 변화가 HS·원산지 판정과 수출통관 리스크로 이어지는지 확인해야 합니다."
    if "통관" in theme:
        return f"{theme}: 신고자료, HS 분류, 과세가격 증빙을 정비해 법인별 통관심사 대응 수준을 맞춰야 합니다."
    return f"{theme}: 후속 고시와 시행일을 추적하고 삼성 적용 품목 여부를 관세 리스크 리스트에 반영해야 합니다."


def build_total_review(top3: pd.DataFrame) -> str:
    countries = []
    themes = []
    for _, row in top3.iterrows():
        for country in clean(row.get("Country", "")).split(","):
            country = country.strip()
            if country and country not in countries:
                countries.append(country)
        theme = detect_theme(row)
        if theme not in themes:
            themes.append(theme)
    country_text = ", ".join(countries[:3]) if countries else "주요국"
    theme_text = "·".join(themes[:3]) if themes else "관세·통상"
    has_export_control = any("수출통제" in theme for theme in themes)
    has_semiconductor = any("반도체" in theme for theme in themes)
    has_tariff = any(("관세" in theme or "301조" in theme or "232조" in theme) for theme in themes)

    if has_semiconductor and has_export_control:
        return (
            f"{country_text}에서 반도체 관세와 수출통제 신호가 동시에 감지됩니다. "
            "본사 관세팀은 SEV/SEVT·SIEL·SAMEX 기준으로 HS, 원산지, ECCN, 최종사용자 통제를 함께 점검해야 합니다."
        )
    if has_tariff:
        return (
            f"{country_text} 중심으로 {theme_text} 변화가 이어지고 있습니다. "
            "본사 관세팀은 적용 품목, 시행일, 관세율 시나리오를 법인별 원가 영향으로 즉시 연결해야 합니다."
        )
    return (
        f"{country_text} 관련 {theme_text} 이슈가 확인됐습니다. "
        "본사 관세팀은 삼성 적용 품목 여부를 먼저 가르고, 필요한 경우 HS·원산지·과세가격 증빙을 보강해야 합니다."
    )


def build_html(rows: pd.DataFrame, review: dict | None = None) -> str:
    subject = f"[GTI Radar] Global Trade Intelligence | {RUN_DATE}"
    top3 = rows.head(3).copy()
    rest = rows.iloc[3:].copy()
    if review:
        total_review = clean(review.get("total_review", "")) or build_total_review(top3)
        review_bullets = review.get("bullets", [])
        if not isinstance(review_bullets, list) or len(review_bullets) != 3:
            review_bullets = [executive_impact(row) for _, row in top3.iterrows()]
    else:
        total_review = build_total_review(top3)
        review_bullets = [executive_impact(row) for _, row in top3.iterrows()]
    bullets = "".join(f"· {html.escape(clean(line))}<br>" for line in review_bullets)

    top_blocks = []
    for idx, row in top3.iterrows():
        top_blocks.append(f"""
        <div style="margin:16px 0 18px 0;padding:14px;border-left:5px solid #C00000;background:#FFF7F7;">
          <div style="font-size:15px;font-weight:bold;margin-bottom:6px;">Top {idx + 1}. {html_link(row['Headline'], row['URL'])}</div>
          <div style="font-size:12px;color:#555;margin-bottom:10px;">Publish Date: {html.escape(clean(row['Date']))} | Country: {html.escape(clean(row['Country']))} | Agency: {html.escape(clean(row['Agency']))} | Risk: <span style="color:{risk_color(row['Risk'])};font-weight:bold;">{html.escape(clean(row['Risk']))}</span> | Samsung Impact Score: {row['Samsung Impact Score']}</div>
          <div style="margin-top:8px;"><b>Executive Impact</b><br>{html.escape(executive_impact(row))}</div>
          <div style="margin-top:8px;"><b>Summary</b><br>{html.escape(clean(row['Summary']))}</div>
          <div style="margin-top:8px;"><b>Impact</b><br>{html.escape(clean(row['AI Analysis']))}</div>
          <div style="margin-top:8px;"><b>Action</b><br>{html.escape(clean(row['Action Plan']))}</div>
        </div>
        """)

    event_rows = []
    for _, row in rest.iterrows():
        event_rows.append(f"""
        <tr>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{row['No']}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html_link(row['Headline'], row['URL'])}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean(row['Summary']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean(row['AI Analysis']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean(row['Action Plan']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{html.escape(clean(row['Country']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;">{html.escape(clean(row['Agency']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;color:{risk_color(row['Risk'])};font-weight:bold;">{html.escape(clean(row['Risk']))}</td>
          <td style="padding:7px;border:1px solid #d9d9d9;text-align:center;">{html.escape(clean(row['Date']))}</td>
        </tr>
        """)

    return f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"><title>{html.escape(subject)}</title></head>
<body style="font-family:Arial,'Malgun Gothic',sans-serif;font-size:13px;color:#222;line-height:1.5;">
  <div style="max-width:1200px;margin:0 auto;">
    <h2 style="margin-bottom:4px;color:#1F4E78;">[GTI Radar] Global Trade Intelligence</h2>
    <div style="font-size:14px;margin-bottom:4px;"><b>Date:</b> {RUN_DATE}</div>
    <div style="font-size:12px;color:#555;margin-bottom:16px;">Coverage: Last {LOOKBACK_HOURS} Hours | Focus: Samsung Electronics Customs & Trade Intelligence</div>

    <h3 style="margin-top:18px;margin-bottom:6px;">총평</h3>
    <div style="margin-top:10px;margin-bottom:20px;padding:16px;background:#F4F6F8;border-left:6px solid #1F4E78;color:#222;">
      <div style="font-size:12pt;font-weight:bold;line-height:1.7;margin-bottom:10px;">{html.escape(total_review)}</div>
      <div style="margin-top:10px;font-size:11pt;font-weight:normal;line-height:1.8;">{bullets}</div>
    </div>

    <h3 style="color:#C00000;margin-top:22px;">TOP POLICY EVENTS (Top 3)</h3>
    {''.join(top_blocks)}

    <h3 style="color:#1F4E78;margin-top:24px;">EVENT LIST</h3>
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
      {''.join(event_rows)}
    </table>
    <p style="margin-top:18px;color:#666;font-size:12px;">첨부 Excel 파일에 전체 Top30 분석표가 포함되어 있습니다.</p>
  </div>
</body>
</html>"""


# =============================================================================
# EXCEL / MAIL
# =============================================================================
def make_headline_formula(title: str, url: str) -> str:
    if clean(url).startswith("http"):
        return f'=HYPERLINK("{url.replace(chr(34), "%22")}","{title.replace(chr(34), "'")}")'
    return title


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
        ws.row_dimensions[row[0].row].height = 84
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def save_excel(rows: pd.DataFrame, paths: dict[str, Path]) -> None:
    headers = ["No", "Date", "Headline", "Summary", "AI Analysis", "Action Plan", "Country", "Agency", "Risk", "Samsung Impact Score", "URL", "Source", "Issue", "Cluster"]
    wb = Workbook()
    ws = wb.active
    ws.title = "GTI Radar Top30"
    ws.append(headers)
    for _, row in rows.iterrows():
        ws.append([row.get(h, "") if h != "Headline" else make_headline_formula(row["Headline"], row["URL"]) for h in headers])
        cell = ws.cell(row=ws.max_row, column=3)
        if clean(row["URL"]).startswith("http"):
            cell.hyperlink = row["URL"]
            cell.font = Font(color="0563C1", underline="single", bold=True)
    widths = {
        "A": 6, "B": 18, "C": 58, "D": 58, "E": 62, "F": 58, "G": 20,
        "H": 34, "I": 9, "J": 18, "K": 42, "L": 24, "M": 20, "N": 28,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    style_sheet(ws)

    ws2 = wb.create_sheet("Top3")
    ws2.append(headers)
    for _, row in rows.head(3).iterrows():
        ws2.append([row.get(h, "") if h != "Headline" else make_headline_formula(row["Headline"], row["URL"]) for h in headers])
    for idx, width in enumerate([6, 18, 58, 58, 62, 58, 20, 34, 9, 18, 42, 24, 20, 28], start=1):
        ws2.column_dimensions[get_column_letter(idx)].width = width
    style_sheet(ws2)

    ws3 = wb.create_sheet("Run Log")
    ws3.append(["item", "value"])
    ws3.append(["input", str(INPUT_FILE)])
    ws3.append(["run_date", RUN_DATE])
    ws3.append(["lookback_hours", LOOKBACK_HOURS])
    ws3.append(["selected_rows", len(rows)])
    ws3.append(["gemini", "ON" if USE_GEMINI and GEMINI_API_KEY else "OFF"])
    style_sheet(ws3)

    wb.save(paths["mail_xlsx"])
    wb.save(paths["analysis"])

    cumul = paths["cumulative"]
    if cumul.exists():
        old = pd.read_excel(cumul)
        combined = pd.concat([old, rows], ignore_index=True)
        combined = combined.drop_duplicates(subset=["URL"], keep="last")
    else:
        combined = rows.copy()
    combined.to_excel(cumul, index=False)


def load_recipients() -> list[str]:
    recipients = [
        x.strip()
        for x in re.split(r"[;,]", MAIL_TO)
        if re.fullmatch(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", x.strip())
    ]
    if recipients:
        return list(dict.fromkeys(recipients))

    for fp in [RECIPIENT_FILE, Path(r"C:\Temp\00.xlsx"), Path(r"C:\Temp\mail.xlsx")]:
        if not fp.exists():
            continue
        try:
            df = pd.read_excel(fp, dtype=str).fillna("")
            text = "\n".join(df.astype(str).values.ravel().tolist())
            found = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
            recipients.extend(found)
        except Exception as exc:
            log(f"[MAIL] recipient file read failed: {fp} / {type(exc).__name__}")
    return list(dict.fromkeys([x.strip() for x in recipients if x.strip()]))


def send_email(html_body: str, attachment: Path) -> None:
    if not SEND_EMAIL:
        log("[MAIL] skipped: GTI_SEND_EMAIL is not Y")
        return
    recipients = load_recipients()
    if not SMTP_USER or not SMTP_PASS:
        log("[MAIL] skipped: GTI_SMTP_USER / GTI_SMTP_PASS is missing")
        return
    if not recipients:
        log(f"[MAIL] skipped: no recipients. Set GTI_MAIL_TO or check {RECIPIENT_FILE}")
        return
    log(f"[MAIL] sending: {len(recipients)} recipients via {SMTP_HOST}:{SMTP_PORT}")
    msg = EmailMessage()
    msg["Subject"] = f"[GTI Radar] Global Trade Intelligence | {RUN_DATE}"
    msg["From"] = formataddr((MAIL_FROM_NAME, SMTP_USER))
    msg["To"] = ", ".join(recipients)
    msg.set_content("GTI Radar HTML 메일입니다. HTML 지원 메일 클라이언트에서 확인해 주세요.")
    msg.add_alternative(html_body, subtype="html")
    with open(attachment, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment.name,
        )
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    log(f"[MAIL SENT] {len(recipients)} recipients")


def main() -> None:
    paths = output_paths()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    raw = read_input()
    log(f"[LOAD] {INPUT_FILE} rows={len(raw)}")
    rows = select_news(raw)
    if rows.empty:
        raise RuntimeError("메일 대상 뉴스가 없습니다.")
    rows = optimize_rows_with_gemini(rows)
    review = optimize_review_with_gemini(rows.head(3))
    html_body = build_html(rows, review)
    save_excel(rows, paths)
    paths["mail_html"].write_text(html_body, encoding="utf-8")
    send_email(html_body, paths["mail_xlsx"])
    log(f"[SELECTED] {len(rows)} rows")
    for path in paths.values():
        log(f"[SAVE] {path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=None)
    parser.add_argument("--input", default=None)
    parser.add_argument("--output-dir", default=None)
    args = parser.parse_args()
    if args.date:
        RUN_DATE = args.date
    if args.input:
        INPUT_FILE = Path(args.input)
    if args.output_dir:
        OUTPUT_DIR = Path(args.output_dir)
    main()
