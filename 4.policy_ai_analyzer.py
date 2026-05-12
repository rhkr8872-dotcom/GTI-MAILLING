# -*- coding: utf-8 -*-
"""
GTI STEP4 Policy AI Analyzer

Input:
  - 3.news_ai_summary.xlsx, or 3.news_master_raw.xlsx
Output:
  - 4.news_ai_analysis.xlsx
  - 4.news_cumulative.xlsx

The script keeps the pipeline usable without a live AI call. If you set
GTI_USE_GEMINI=Y and GEMINI_API_KEY, the Gemini hook can be extended later;
the deterministic expert rules below remain the safety fallback.
"""

from __future__ import annotations

import os
import re
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\temp"))
OUTPUT_DIR = Path(os.getenv("GTI_OUTPUT_DIR", ".")).resolve()
TOP_N = int(os.getenv("GTI_TOP_N", "30"))

INPUT_CANDIDATES = [
    BASE_DIR / "3.news_ai_summary.xlsx",
    BASE_DIR / "3.news_master_raw.xlsx",
    BASE_DIR / "news_ai_summary.xlsx",
    BASE_DIR / "news_raw.xlsx",
]
CUMULATIVE_CANDIDATES = [
    BASE_DIR / "4.news_cumulative.xlsx",
    BASE_DIR / "news_cumulative.xlsx",
]

OUT_ANALYSIS = OUTPUT_DIR / "4.news_ai_analysis.xlsx"
OUT_CUMULATIVE = OUTPUT_DIR / "4.news_cumulative.xlsx"

FINAL_COLUMNS = [
    "Date",
    "Headline",
    "importance",
    "출처URL",
    "source",
    "last_checked",
    "Summary",
    "AI Analysis",
    "Action Plan",
    "Country",
    "agency",
]

FOCUS_COUNTRIES = {
    "Korea": ["korea", "korean", "관세청", "한국", "산업통상", "기획재정부"],
    "United States": ["united states", "u.s.", " us ", "america", "ustr", "cbp", "federal register", "commerce department"],
    "China": ["china", "chinese", "mofcom", "gacc", "중국"],
    "Vietnam": ["vietnam", "vietnamese", "베트남"],
    "India": ["india", "indian", "인도"],
    "Mexico": ["mexico", "mexican", "멕시코"],
    "Brazil": ["brazil", "brazilian", "브라질"],
    "EU": ["european commission", "taxud", " eu ", "european union", "유럽연합"],
    "Indonesia": ["indonesia", "인도네시아"],
    "Turkey": ["turkey", "turkiye", "튀르키예"],
    "Poland": ["poland", "폴란드"],
    "Slovakia": ["slovakia", "슬로바키아"],
}

AGENCY_PATTERNS = [
    ("USTR", ["ustr", "trade representative"]),
    ("U.S. Customs and Border Protection (CBP)", ["cbp", "customs and border protection"]),
    ("U.S. Department of Commerce", ["commerce department", "department of commerce"]),
    ("U.S. Federal Register", ["federal register"]),
    ("Korea Customs Service", ["관세청", "korea customs"]),
    ("Korea MOTIE", ["산업통상자원부", "motie"]),
    ("European Commission / DG TAXUD", ["european commission", "taxud"]),
    ("China MOFCOM", ["mofcom", "중국 상무부"]),
    ("China GACC", ["gacc", "중국 해관"]),
    ("Vietnam Customs", ["vietnam customs", "customs department is piloting", "hai quan"]),
    ("India Ministry of Commerce & Industry", ["india", "indian"]),
    ("Mexico Ministry of Economy / SAT", ["mexico"]),
    ("Brazil MDIC / Receita Federal", ["brazil"]),
    ("WTO", ["wto", "world trade organization"]),
    ("WCO", ["wco", "world customs organization"]),
]

POLICY_TERMS = [
    "tariff", "customs", "fta", "free trade", "trade agreement", "import", "export",
    "origin", "hs code", "valuation", "anti-dumping", "countervailing", "safeguard",
    "section 232", "section 301", "cbam", "sanction", "export control", "clearance",
    "duty", "refund", "excise", "de minimis", "supply chain", "관세", "통관", "수출",
    "수입", "무역", "원산지", "덤핑", "상계관세", "세이프가드", "FTA", "운임", "과세",
]

NOISE_TERMS = [
    "fire", "stolen", "lamborghini", "rolls-royce", "marijuana", "narcotics", "drug",
    "hezbollah", "customs official who fled", "office complex", "motorcycle", "scout customs",
    "custom bike", "pakistan customs fire", "마약", "화재", "람보르기니", "롤스로이스",
]

OFF_TOPIC_PRODUCT_TERMS = [
    "beef", "meat", "seafood", "fruit", "vegetable", "gold", "silver", "jewelry",
    "tire", "tyre", "cotton", "soybean", "pork", "chicken", "wine", "fishery",
    "crypto", "luxury car", "whiskey", "nike", "west bank", "settlers", "gm romulus",
    "royal whiskey", "violent israeli",
    "쇠고기", "수산", "과일", "채소", "금 ", "은 ", "타이어", "농산물",
]

SAMSUNG_RELEVANT_TERMS = [
    "electronics", "semiconductor", "chip", "smartphone", "mobile", "appliance",
    "network equipment", "display", "battery", "steel", "aluminum", "customs procedure",
    "clearance", "origin", "valuation", "supply chain", "logistics", "export control",
    "전자", "반도체", "스마트폰", "가전", "디스플레이", "배터리", "철강", "알루미늄",
    "통관", "원산지", "과세가격", "물류", "수출통제", "공급망", "운임",
]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).replace("\r", "\n")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def norm_title(title: str) -> str:
    text = clean_text(title).lower()
    text = re.sub(r"[-–|].*$", "", text)
    text = re.sub(r"[^0-9a-z가-힣]+", " ", text)
    stop = {"the", "a", "an", "and", "to", "of", "in", "for", "on", "with", "from", "says"}
    return " ".join(w for w in text.split() if w not in stop)


def source_name(url: str, source: str) -> str:
    if clean_text(source) and not str(source).startswith("http"):
        return clean_text(source)
    host = urlparse(clean_text(url)).netloc.replace("www.", "")
    return host or clean_text(source)


def load_inputs() -> tuple[pd.DataFrame, str]:
    frames = []
    used = []
    for path in INPUT_CANDIDATES:
        if path.exists():
            frames.append(pd.read_excel(path))
            used.append(str(path))
    if not frames:
        raise FileNotFoundError("STEP4 input file was not found.")
    return pd.concat(frames, ignore_index=True), ", ".join(used)


def pick_cumulative() -> Path | None:
    for path in CUMULATIVE_CANDIDATES:
        if path.exists():
            return path
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    lower = {str(c).strip().lower(): c for c in df.columns}
    mapping = {
        lower.get("date", "date"): "Date",
        lower.get("title", lower.get("headline", "title")): "Headline",
        lower.get("url", "url"): "출처URL",
        lower.get("source", "source"): "source",
        lower.get("collected_at", lower.get("last_checked", "collected_at")): "last_checked",
        lower.get("agency", "agency"): "agency",
        lower.get("importance", "importance"): "importance_raw",
        lower.get("score", "score"): "score_raw",
    }
    existing = {k: v for k, v in mapping.items() if k in df.columns}
    out = df.rename(columns=existing).copy()
    for col in ["Date", "Headline", "출처URL", "source", "last_checked", "agency", "importance_raw", "score_raw"]:
        if col not in out.columns:
            out[col] = ""
    return out


def countries_for(text: str) -> list[str]:
    raw = f" {text.lower()} "
    found = []
    for country, terms in FOCUS_COUNTRIES.items():
        if any(term.lower() in raw for term in terms):
            found.append(country)
    return found[:2] or ["Global"]


def agency_for(text: str, current: str) -> str:
    raw = f" {text.lower()} "
    for agency, terms in AGENCY_PATTERNS:
        if any(term.lower() in raw for term in terms):
            return agency
    cur = clean_text(current)
    return cur if cur and cur.lower() != "nan" else "Trade / Customs authority"


def topic_for(text: str) -> str:
    raw = text.lower()
    if any(x in raw for x in ["tariff", "duty", "section 232", "section 301", "refund", "관세율", "과세"]):
        return "tariff"
    if any(x in raw for x in ["fta", "free trade", "trade agreement", "customs union"]):
        return "fta"
    if any(x in raw for x in ["clearance", "customs procedure", "origin", "valuation", "hs code", "통관", "원산지", "운임"]):
        return "customs"
    if any(x in raw for x in ["export control", "sanction", "anti-dumping", "countervailing", "safeguard", "수출통제", "덤핑"]):
        return "trade_remedy"
    if any(x in raw for x in ["export", "import", "supply chain", "수출입"]):
        return "trade_flow"
    return "general"


def risk_for(row: pd.Series, topic: str, countries: list[str]) -> str:
    raw = f"{row.get('Headline','')} {row.get('importance_raw','')} {row.get('score_raw','')}".lower()
    focused = any(c in {"United States", "China", "Vietnam", "India", "Mexico", "Brazil", "Korea", "EU"} for c in countries)
    direct = any(term.lower() in raw for term in SAMSUNG_RELEVANT_TERMS)
    official = any(x in raw for x in ["customs.go.kr", "taxation-customs.ec.europa.eu", "federal register", "ustr", "cbp"])
    if topic in {"tariff", "trade_remedy"} and focused and (direct or official):
        return "상"
    if topic in {"tariff", "trade_remedy"} and focused:
        return "중"
    if topic in {"fta", "customs"} and focused:
        return "중"
    return "하"


def summary_for(headline: str, topic: str, countries: list[str], agency: str) -> str:
    country_txt = ", ".join(countries)
    if topic == "tariff":
        return f"{country_txt} 관련 관세·수입부담 변화 가능성이 제기된 사안입니다. {agency} 발표 또는 보도 내용을 기준으로 적용 품목, 시행시점, 환급·감면 조건 확인이 필요합니다."
    if topic == "fta":
        return f"{country_txt}의 FTA 또는 관세동맹 협력 방향에 관한 뉴스입니다. 향후 원산지 기준, 특혜관세 적용 가능성, 역내 공급망 활용 조건이 달라질 수 있습니다."
    if topic == "customs":
        return f"{country_txt} 통관 절차, 과세가격, 원산지 또는 세관 운영 변경과 관련된 사안입니다. 수입신고·서류·심사 대응 방식에 실무 영향이 있을 수 있습니다."
    if topic == "trade_remedy":
        return f"{country_txt} 수출통제·무역구제·제재 관련 정책 신호입니다. 대상 품목과 거래상대, 원산지 기준에 따라 수출입 제한 또는 추가 비용이 발생할 수 있습니다."
    if topic == "trade_flow":
        return f"{country_txt} 수출입 흐름과 공급망 여건에 관한 동향입니다. 직접적인 세율 변경은 아니지만 물류·통관 리드타임과 비용 추이를 확인할 필요가 있습니다."
    return f"{country_txt} 관세·통상 관련 일반 동향입니다. 삼성전자 업무 관련성은 원문에서 대상 품목과 정책 근거를 추가 확인해야 합니다."


def analysis_for(topic: str, countries: list[str]) -> str:
    c = set(countries)
    sites = []
    if "Vietnam" in c:
        sites.append("베트남(SEV/SEVT)")
    if "India" in c:
        sites.append("인도(SIEL)")
    if "Mexico" in c:
        sites.append("멕시코(SAMEX)")
    if "China" in c:
        sites.append("중국")
    if "Korea" in c:
        sites.append("한국 본사/수출입")
    site_txt = ", ".join(sites) if sites else "주요 생산·판매 법인"

    if topic == "tariff":
        return f"{site_txt}의 모바일·가전·네트워크 장비 관련 수입원가와 통관세액에 직접 영향 가능성이 있습니다. 세율, 환급, 감면 요건이 바뀌면 가격·원산지·HS 운영을 즉시 재점검해야 합니다."
    if topic == "fta":
        return f"{site_txt} 공급망에서 특혜관세 활용 또는 원산지 충족 전략에 영향이 있을 수 있습니다. 인도·베트남·멕시코 생산품의 역내 조달 구조와 증빙 체계 점검이 필요합니다."
    if topic == "customs":
        return f"{site_txt}의 수출입 신고, 과세가격, 운임 가산, 원산지 증빙 업무에 운영 리스크가 있습니다. 통관 지연 또는 사후심사 이슈가 생산·출하 일정에 연결될 수 있습니다."
    if topic == "trade_remedy":
        return f"{site_txt} 제품이 조사·제재·수출통제 대상 품목과 겹칠 경우 선적 제한, 추가관세, 고객 납기 리스크가 발생할 수 있습니다. 반도체·네트워크 장비는 특히 민감 품목 여부 확인이 필요합니다."
    return f"{site_txt} 관점에서 즉각적인 직접 영향은 제한적이나, 관세·통상 정책 신호로 모니터링 가치가 있습니다. 관련 국가 법인의 품목·거래선 노출 여부를 확인해야 합니다."


def action_for(topic: str, countries: list[str], agency: str) -> str:
    base = f"{agency} 원문과 시행일을 확인하고, 대상 HS·원산지·법인별 거래 흐름을 매핑하십시오."
    if topic == "tariff":
        return base + " 예상 세액 변동, 환급 가능성, 가격 전가 여부를 재무·물류와 함께 산정해야 합니다."
    if topic == "fta":
        return base + " BOM 원산지, 공급업체 증빙, FTA 판정 로직을 사전 검증해 특혜관세 활용안을 정리하십시오."
    if topic == "customs":
        return base + " 신고가격·운임·통관서류 샘플을 점검하고 현지 법인 브로커 대응 가이드를 갱신하십시오."
    if topic == "trade_remedy":
        return base + " 제재·조사 대상 품목과 고객을 대조하고 필요 시 선적 보류 또는 라이선스 검토 절차를 준비하십시오."
    return base + " 직접 영향이 확인되면 차기 GTI 메일에서 후속조치 항목으로 승격하십시오."


def is_policy_relevant(row: pd.Series) -> bool:
    text = f"{row.get('Headline','')} {row.get('source','')} {row.get('agency','')}".lower()
    has_policy = any(term.lower() in text for term in POLICY_TERMS)
    has_noise = any(term.lower() in text for term in NOISE_TERMS)
    if has_noise and not any(x in text for x in ["tariff", "fta", "trade", "origin", "valuation", "refund", "procedure"]):
        return False
    off_topic = any(term.lower() in text for term in OFF_TOPIC_PRODUCT_TERMS)
    samsung_relevant = any(term.lower() in text for term in SAMSUNG_RELEVANT_TERMS)
    official_source = any(x in text for x in ["customs.go.kr", "taxation-customs.ec.europa.eu", "federal register", "ustr", "cbp", "mofcom", "gacc"])
    if off_topic and not samsung_relevant:
        return False
    return has_policy


def dedupe_similar(df: pd.DataFrame) -> pd.DataFrame:
    seen_urls = set()
    kept = []
    titles = []
    for _, row in df.iterrows():
        url = clean_text(row["출처URL"])
        title = norm_title(row["Headline"])
        if not title:
            continue
        if url and url in seen_urls:
            continue
        duplicate = False
        for prev in titles:
            if SequenceMatcher(None, title, prev).ratio() >= 0.86:
                duplicate = True
                break
        if duplicate:
            continue
        seen_urls.add(url)
        titles.append(title)
        kept.append(row)
    return pd.DataFrame(kept)


def run() -> None:
    input_df, input_label = load_inputs()
    raw = normalize_columns(input_df)
    raw["Date"] = pd.to_datetime(raw["Date"], errors="coerce")
    raw["last_checked"] = pd.to_datetime(raw["last_checked"], errors="coerce").fillna(pd.Timestamp.now())
    raw["Headline"] = raw["Headline"].map(clean_text)
    raw["출처URL"] = raw["출처URL"].map(clean_text)
    raw["source"] = [source_name(u, s) for u, s in zip(raw["출처URL"], raw["source"])]
    raw = raw.dropna(subset=["Date"])

    run_date = pd.to_datetime(os.getenv("GTI_RUN_DATE", datetime.now().strftime("%Y-%m-%d")))
    cutoff = run_date - timedelta(days=1)
    recent = raw[raw["Date"] >= cutoff].copy()
    if recent.empty:
        recent = raw.sort_values("Date", ascending=False).head(200).copy()

    recent = recent[recent.apply(is_policy_relevant, axis=1)].copy()

    cumulative_path = pick_cumulative()
    if cumulative_path:
        old = normalize_columns(pd.read_excel(cumulative_path))
        old_urls = set(old["출처URL"].dropna().map(clean_text))
        old_titles = set(old["Headline"].dropna().map(norm_title))
        recent = recent[
            ~recent["출처URL"].map(clean_text).isin(old_urls)
            | ~recent["Headline"].map(norm_title).isin(old_titles)
        ].copy()

    recent["score_num"] = pd.to_numeric(recent["score_raw"], errors="coerce").fillna(0)
    recent = recent.sort_values(["score_num", "Date"], ascending=[False, False])
    recent = dedupe_similar(recent)

    records = []
    for _, row in recent.iterrows():
        text = f"{row['Headline']} {row['source']} {row['agency']}"
        countries = countries_for(text)
        agency = agency_for(text, row.get("agency", ""))
        topic = topic_for(text)
        importance = risk_for(row, topic, countries)
        records.append(
            {
                "Date": row["Date"].strftime("%Y:%m:%d:%H"),
                "Headline": clean_text(row["Headline"]),
                "importance": importance,
                "출처URL": clean_text(row["출처URL"]),
                "source": clean_text(row["source"]),
                "last_checked": pd.to_datetime(row["last_checked"]).strftime("%Y-%m-%d %H:%M"),
                "Summary": summary_for(row["Headline"], topic, countries, agency),
                "AI Analysis": analysis_for(topic, countries),
                "Action Plan": action_for(topic, countries, agency),
                "Country": " / ".join(countries[:2]),
                "agency": agency,
                "_rank": {"상": 1, "중": 2, "하": 3}[importance],
                "_score": float(row.get("score_num", 0)),
            }
        )

    result = pd.DataFrame(records)
    if result.empty:
        result = pd.DataFrame(columns=FINAL_COLUMNS + ["_rank", "_score"])
    result = result.sort_values(["_rank", "_score", "Date"], ascending=[True, False, False]).head(TOP_N)
    result = result[FINAL_COLUMNS]

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    result.to_excel(OUT_ANALYSIS, index=False)

    cumulative = result.copy()
    if OUT_CUMULATIVE.exists() and os.getenv("GTI_RESET_CUMULATIVE", "N").upper() != "Y":
        previous = pd.read_excel(OUT_CUMULATIVE)
        cumulative = pd.concat([previous, cumulative], ignore_index=True)
    cumulative = cumulative.drop_duplicates(subset=["출처URL"], keep="last")
    cumulative.to_excel(OUT_CUMULATIVE, index=False)

    for path in [OUT_ANALYSIS, OUT_CUMULATIVE]:
        style_workbook(path)

    print(f"[DONE] rows={len(result)} input={input_label}")
    print(f"[SAVE] {OUT_ANALYSIS}")
    print(f"[SAVE] {OUT_CUMULATIVE}")


def style_workbook(path: Path) -> None:
    wb = load_workbook(path)
    ws = wb.active
    ws.title = "GTI Analysis"
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    widths = [15, 48, 10, 42, 22, 18, 56, 60, 58, 20, 32]
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
        ws.row_dimensions[row[0].row].height = 92
        headline = row[1]
        url = row[3].value
        if url:
            headline.hyperlink = url
            headline.font = Font(color="0563C1", bold=True, underline="single")
        risk = row[2].value
        fill = {"상": "F4CCCC", "중": "FFF2CC", "하": "D9EAD3"}.get(risk, "FFFFFF")
        row[2].fill = PatternFill("solid", fgColor=fill)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


if __name__ == "__main__":
    run()
