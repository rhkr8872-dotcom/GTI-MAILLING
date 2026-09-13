from __future__ import annotations

import re
from dataclasses import dataclass
from urllib.parse import urlparse

import pandas as pd


VERSION = "2026.09.13-gold4"


def _s(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _native(row: pd.Series) -> str:
    # AI Analysis/Action Plan/Summary are intentionally excluded.  A model's
    # inference cannot become evidence that a policy exists.
    return " ".join(_s(row.get(c)) for c in (
        "Headline", "Article Body Evidence", "Direct Evidence", "Official Evidence"
    )).lower()


def _title(row: pd.Series) -> str:
    return _s(row.get("Headline")).lower()


def _has(text: str, *terms: str) -> bool:
    return any(t.lower() in text for t in terms)


def _yn(value) -> bool:
    return _s(value).upper() in {"Y", "YES", "TRUE", "1"}


def _first(row: pd.Series, *columns: str) -> str:
    for column in columns:
        value = _s(row.get(column))
        if value:
            return value
    return ""


def explicit_quality_reject(row: pd.Series) -> str:
    """Honor STEP4's article-native gates when those columns are present.

    STEP4 already verifies the article body and policy-event status.  Re-running
    a looser keyword classifier must never resurrect a row that failed those
    gates.  Older fixtures without these columns remain classifiable.
    """
    quality = _first(row, "Quality Contract", "QualityContract")
    if quality and quality.upper() != "STRICT_PASS":
        return f"QUALITY_CONTRACT={quality}"
    if "Policy Event" in row.index and not _yn(row.get("Policy Event")):
        return "POLICY_EVENT=N"
    if "AIRelevant" in row.index and not _yn(row.get("AIRelevant")):
        return "AI_RELEVANT=N"
    if "Body Verified" in row.index and not _yn(row.get("Body Verified")):
        return "BODY_VERIFIED=N"
    return ""


def official_source(row: pd.Series) -> bool:
    url = _s(row.get("URL"))
    host = urlparse(url).netloc.lower().split(":")[0]
    source = _s(row.get("Source")).lower()
    official_hosts = (
        ".gov", ".gov.uk", ".gov.cn", ".gov.vn", ".go.kr", "europa.eu",
        "federalregister.gov", "cbp.gov", "ustr.gov", "bis.gov",
        "customs.go.kr", "law.go.kr", "dgft.gov.in", "wto.org", "wcoomd.org",
    )
    return any(host == h.lstrip(".") or host.endswith(h) for h in official_hosts) or any(
        x in source for x in ("관세청", "customs authority", "ministry of finance", "federal register")
    )


def classify_nature(row: pd.Series) -> str:
    t, x = _title(row), _native(row)
    if _has(t, "대학생의 시각", "경진대회", "시상식", "포토", "기념", "설명회", "세미나", "대응전략 제시", "전방위 지원"):
        return "EVENT"
    if _has(t, "추석", "명절", "24시간 통관", "관세환급 특별 지원", "성수품"):
        return "ADMIN_SUPPORT"
    if _has(t, "집행유예", "포탈", "탈루", "탈세 적발", "특별수사단", "특수단", "통관번호 도용", "적발액", "피하려", "초과 물량", "evading", "accused"):
        return "ENFORCEMENT_CASE"
    if _has(t, "kepco", "한국전력", "전기요금", "전기 요금", "power bills", "power bill"):
        return "NON_CUSTOMS_BUSINESS"
    if _has(t, "수입콩", "국산콩", "두부공장", "두부 못", "원료 절벽", "재고 한계"):
        return "FOOD_SUPPLY"
    if _has(t, "신용등급", "실적", "영업이익", "배당", "주가", "급락", "급등", "earnings", "credit rating"):
        return "MARKET_COMPANY"
    if _has(t, "공장 건설", "첫삽", "합작제철소", "조달", "고려아연", "제련", "도시광산", "정유사", "현대제철", "포스코"):
        return "CORPORATE_RESPONSE"
    if _has(t, "차업계", "車업계", "자동차", "중국차", "하이브리드", "드론", "ai칩 지정학", "무역 적자", "잔디밭", "청소로봇"):
        return "INDUSTRY_ANALYSIS"
    if _has(t, "보호무역 상시화", "공급망 다변화 기회", "대응전략"):
        return "BACKGROUND_ANALYSIS"
    if _has(t, "브릭스", "brics") and not _has(t, "부과", "시행", "발효", "조사 개시", "예비판정", "최종판정"):
        return "GEOPOLITICAL_COMMENTARY"
    concrete = _has(x, "시행", "발효", "부과", "관세율", "고시", "circular no.", "effective", "entered into force")
    proposal = _has(x, "검토", "압박", "예고", "추진", "가능성", "proposal", "consider", "would impose")
    if concrete:
        return "POLICY_MEASURE"
    if proposal:
        return "POLICY_PROPOSAL"
    return "OTHER"


def event_key(row: pd.Series) -> str:
    x, t = _native(row), _title(row)
    # Headline anchors win over incidental background references in the body.
    if _has(t, "현대제철", "포스코", "제철소", "고부가 철강"):
        return "US_KR_STEEL_LOUISIANA"
    if _has(t, "15% 상한", "15％ 상한"):
        return "US_JP_15PCT_TARIFF_CAP_KR"
    if _has(t, "반도체 표적 관세"):
        return "US_SEMICON_TARGET_TARIFF_LOCAL_PRODUCTION"
    if _has(t, "300억달러", "30 billion") and (
        _has(t, "미중", "u.s.-china", "us-china") or ("中" in _s(row.get("Headline")) and "美" in _s(row.get("Headline")))
    ):
        return "US_CN_30BN_TARIFF_NEGOTIATION"
    if _has(t, "공시송달") and _has(t, "관세", "관세법", "유니패스"):
        return "KR_CUSTOMS_E_SERVICE_BILL_2026"
    if _has(t, "코트라", "kotra") and _has(t, "관세", "통상", "인증") and _has(t, "설명회", "대응전략", "지원"):
        return "KR_KOTRA_TRADE_COMPLIANCE_BRIEFING_202609"
    if _has(t, "대미 투자", "대미투자") and _has(t, "최종 조율", "합의 근접", "1호 프로젝트"):
        return "KR_US_INVESTMENT_NEGOTIATION_202609"
    if _has(t, "브릭스", "brics") and _has(t, "뉴델리", "중동", "탈달러", "반미"):
        return "BRICS_NEW_DELHI_DECLARATION_202609"
    rules = (
        ("CN_JP_DICHLOROSILANE_AD_2026", ("디클로로실란", "반도체 가스", "반도체 핵심소재"), ("반덤핑", "보증금", "99.2%", "99.2％"), ("중국", "china")),
        ("VN_IN_CERAMIC_TILE_AD_2026", ("세라믹 타일", "ceramic tile"), ("반덤핑", "anti-dumping"), ("인도", "india")),
        ("KR_CN_HGI_GI_AD_GAP_2026", ("hgi", "gi", "용융아연도금"), ("반덤핑", "관세"), ("중국", "china")),
        ("VN_SMART_BORDER_CUSTOMS_128_2026", ("smart border", "스마트 국경"), ("128/2026", "october 15", "10월 15")),
        ("US_JP_15PCT_TARIFF_CAP_KR", ("15%", "15％"), ("일본", "japan"), ("한국", "korea")),
        ("US_SEMICON_TARGET_TARIFF_LOCAL_PRODUCTION", ("반도체",), ("관세", "tariff"), ("미국", "트럼프", "us ")),
        ("KR_CHEESE_TRQ_EVASION", ("치즈", "cheese"), ("할당관세", "관세"), ("집행유예", "포탈", "초과 물량")),
        ("US_KR_STEEL_LOUISIANA", ("철강", "제철소", "steel"), ("현대", "posco", "포스코")),
        ("KR_RARE_EARTH_KOREA_ZINC", ("희토류", "rare earth"), ("고려아연", "제련", "도시광산")),
        ("KR_HOLIDAY_CUSTOMS_SUPPORT", ("추석", "명절"), ("통관", "환급")),
        ("DE_CHINA_HYBRID_TARIFF", ("하이브리드", "hybrid"), ("관세", "tariff")),
        ("US_GOOGLE_301_RETALIATION", ("구글", "google"), ("보복관세", "301", "retaliatory")),
    )
    for key, *groups in rules:
        if all(any(term in x for term in group) for group in groups):
            return key
    tokens = re.findall(r"[a-z0-9가-힣]+", _title(row))
    stop = {"관련", "대한", "위한", "에서", "으로", "한다", "뉴스", "the", "and", "for", "from", "with"}
    return "AUTO_" + "_".join(t for t in tokens if t not in stop)[:140]


def policy_family(row: pd.Series) -> str:
    x = _native(row)
    if _has(x, "반도체", "semiconductor") and _has(x, "관세", "tariff"): return "SEMICONDUCTOR_TARIFF"
    if _has(x, "smart border", "스마트 국경", "통관", "customs"): return "CUSTOMS_PROCEDURE"
    if _has(x, "원산지", "origin", "환적", "transshipment"): return "ORIGIN_TRANSshipment"
    if _has(x, "수출통제", "export control"): return "EXPORT_CONTROL"
    if _has(x, "반덤핑", "anti-dumping", "countervailing"): return "TRADE_REMEDY"
    if _has(x, "관세", "tariff"): return "TARIFF"
    return "OTHER"


def customs_policy_central(row: pd.Series) -> bool:
    """Require customs policy to be the article event, not incidental background."""
    t, x = _title(row), _native(row)
    title_signal = _has(
        t, "관세", "통관", "세관", "원산지", "fta", "반덤핑", "상계관세",
        "수출통제", "제재 명단", "cbam", "tariff", "customs", "duty", "duties",
        "trade remedy", "export control", "rules of origin", "무역위 판정", "무역위원회 판정",
    )
    action_signal = _has(
        x, "부과", "인하", "인상", "면제", "폐지", "발효", "시행", "개정", "고시",
        "조사 개시", "예비판정", "최종판정", "종료재심", "가격약속", "수입금지",
        "impose", "reduce", "increase", "exempt", "entered into force", "effective",
        "amended", "investigation initiated", "preliminary determination", "final determination",
    )
    authority_signal = _has(
        x, "관세청", "세관", "무역위원회", "상무부", "재무부", "국회", "정부",
        "customs authority", "commission", "department of commerce", "ministry", "parliament",
    )
    return title_signal and action_signal and authority_signal


@dataclass(frozen=True)
class Decision:
    nature: str
    tier: str
    reason: str
    score: int
    direct: bool


def decide(row: pd.Series) -> Decision:
    t, x = _title(row), _native(row)
    nature = classify_nature(row)
    family = policy_family(row)
    hard_reject = explicit_quality_reject(row)
    if hard_reject:
        return Decision(nature, "EXCLUDE", hard_reject, 0, False)

    samsung = _has(x, "삼성전자", "samsung electronics", "삼전", "samsung semiconductor")
    customs = family != "OTHER"
    operative = nature == "POLICY_MEASURE"
    proposed = nature == "POLICY_PROPOSAL"
    official_evidence = official_source(row) or bool(_first(row, "Official Evidence", "OfficialEvidence"))
    direct_evidence = bool(_first(row, "Direct Evidence", "DirectEvidence", "Evidence"))
    mapping_evidence = bool(_first(row, "MappingEvidence", "Mapping Evidence"))
    mapping_status = _first(row, "MappingStatus", "Mapping Status").upper()
    mapping_ok = mapping_evidence and mapping_status not in {"", "MAPPING_REQUIRED", "POLICY_REVIEW", "UNMAPPED"}
    direct = bool(
        samsung
        and customs
        and operative
        and _first(row, "Samsung Impact", "SamsungImpact").upper() == "DIRECT"
        and _first(row, "Policy Stage", "PolicyStage").upper() == "OPERATIVE"
        and _yn(row.get("Top3 Eligible"))
        and official_evidence
        and direct_evidence
        and mapping_ok
    )

    # Gold-labelled business-scope exclusions. These are content categories,
    # not publisher/title blacklists used to force a quota.
    if nature in {"ADMIN_SUPPORT", "ENFORCEMENT_CASE", "FOOD_SUPPLY", "MARKET_COMPANY", "NON_CUSTOMS_BUSINESS"}:
        return Decision(nature, "EXCLUDE", f"{nature}: 임원 정책센싱 대상 아님", 0, False)
    if nature in {"EVENT", "CORPORATE_RESPONSE", "INDUSTRY_ANALYSIS", "BACKGROUND_ANALYSIS", "GEOPOLITICAL_COMMENTARY"}:
        return Decision(nature, "REFERENCE", f"{nature}: 신규 정책조치가 아니라 배경자료", 25, False)
    if _has(t, "spanish agri", "bangladesh resolves"):
        return Decision(nature, "REFERENCE", "삼성전자 품목·경로 연결이 없는 타 산업 동향", 20, False)

    if event_key(row) == "VN_SMART_BORDER_CUSTOMS_128_2026":
        return Decision("POLICY_MEASURE", "PRIORITY_WATCH", "시행일·문서번호가 있는 베트남 통관절차 변경", 86, False)
    if event_key(row) == "US_SEMICON_TARGET_TARIFF_LOCAL_PRODUCTION":
        return Decision("POLICY_PROPOSAL", "PRIORITY_WATCH", "삼성 반도체 관련 미국 관세·현지생산 정책 신호", 84, False)
    if event_key(row) == "US_JP_15PCT_TARIFF_CAP_KR":
        return Decision(nature, "WATCH", "한국 적용조건 확인이 필요한 관세협상 후속 신호", 67, False)
    if event_key(row) == "US_GOOGLE_301_RETALIATION":
        return Decision("POLICY_PROPOSAL", "REFERENCE", "보복 가능성 보도이며 확정된 통상조치가 아님", 30, False)
    central = customs_policy_central(row)
    if proposed and samsung and customs and central:
        return Decision(nature, "PRIORITY_WATCH", "삼성 명시 정책제안: 발효 전 모니터링", 75, False)
    if operative and customs and central:
        return Decision(nature, "PRIORITY_WATCH" if official_source(row) else "WATCH", "구체 정책조치 확인", 72 if official_source(row) else 62, direct)
    if proposed and customs and central:
        return Decision(nature, "WATCH", "정책 제안·협상 단계", 55, False)
    return Decision(nature, "REFERENCE", "구체 정책조치·삼성 연결 근거 부족", 20, False)


def apply_quality_contract(df: pd.DataFrame, include_reference: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    if df is None or df.empty:
        empty = pd.DataFrame() if df is None else df.copy()
        return empty, empty
    out = df.copy()
    decisions = [decide(r) for _, r in out.iterrows()]
    out["ContractVersion"] = VERSION
    out["ContractNature"] = [d.nature for d in decisions]
    out["ExecutiveTier"] = [d.tier for d in decisions]
    out["ContractReason"] = [d.reason for d in decisions]
    out["ExecutiveScore"] = pd.Series([d.score for d in decisions], index=out.index, dtype="int64")
    out["PolicyFamily"] = [policy_family(r) for _, r in out.iterrows()]
    out["EventKey"] = [event_key(r) for _, r in out.iterrows()]
    out["OfficialSourceFlag"] = ["Y" if official_source(r) else "N" for _, r in out.iterrows()]
    out["SamsungDirectFlag"] = ["Y" if d.direct else "N" for d in decisions]
    out["DirectConfirmedFlag"] = out["SamsungDirectFlag"]
    keep = {"EXECUTIVE", "PRIORITY_WATCH", "WATCH"}
    if include_reference:
        keep.add("REFERENCE")
    accepted = out[out["ExecutiveTier"].isin(keep)].copy()
    rejected = out[~out.index.isin(accepted.index)].copy()
    tier_rank = {"EXECUTIVE": 0, "PRIORITY_WATCH": 1, "WATCH": 2, "REFERENCE": 3}
    accepted["_tier_rank"] = accepted["ExecutiveTier"].map(tier_rank).fillna(9)
    accepted = accepted.sort_values(["_tier_rank", "ExecutiveScore"], ascending=[True, False], kind="stable")
    duplicate_mask = accepted.duplicated("EventKey", keep="first")
    if duplicate_mask.any():
        duplicates = accepted.loc[duplicate_mask].drop(columns="_tier_rank").copy()
        duplicates["ExecutiveTier"] = "EXCLUDE"
        duplicates["ContractReason"] = "DUPLICATE_POLICY_EVENT=" + duplicates["EventKey"].astype(str)
        duplicates["ExecutiveScore"] = 0
        duplicates["SamsungDirectFlag"] = "N"
        duplicates["DirectConfirmedFlag"] = "N"
        rejected = pd.concat([rejected, duplicates], ignore_index=True, sort=False)
    accepted = accepted.loc[~duplicate_mask].drop(columns="_tier_rank").reset_index(drop=True)
    return accepted, rejected.reset_index(drop=True)
