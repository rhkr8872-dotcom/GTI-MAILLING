# -*- coding: utf-8 -*-
"""Deterministic Samsung Customs Action Queue contract.

Only an evidenced policy/operational change with a confirmed Samsung entity,
product/HS or trade route, and an event-specific action may enter Action Queue.
Everything else is retained in Policy Radar / Global Context or excluded when
it is an administrative notice.
"""
from __future__ import annotations

import re
from typing import Iterable

import pandas as pd


VERSION = "2026.09.14-action-queue-1"

ADMIN_PATTERNS = (
    r"채용|채용공고|기간제|인사발령|인사 공고|입찰공고|용역공고|구매공고|"
    r"교육생 모집|교육 안내|세미나|설명회|행사 안내|공모전|장학생|"
    r"업무협약|mou|recruit|vacanc(?:y|ies)|tender notice|procurement notice"
)
OPERATIONAL_PATTERNS = (
    r"일시정지|일시 중지|서비스 중단|시스템 점검|전산 점검|전산장애|"
    r"납부.*중지|수납.*정지|maintenance window|service interruption|system outage"
)
PROPOSAL_PATTERNS = r"법률안|입법예고|행정예고|개정안|초안|proposal|proposed|draft|검토 중|예고"


def _s(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _flag(row: pd.Series, names: Iterable[str]) -> bool:
    return any(_s(row.get(name)).upper() in {"Y", "YES", "TRUE", "1"} for name in names)


def _first(row: pd.Series, names: Iterable[str]) -> str:
    for name in names:
        value = _s(row.get(name))
        if value and value.lower() not in {"nan", "none", "관련 법인 검토", "본문에서 확인 불가"}:
            return value
    return ""


def _native_text(row: pd.Series) -> str:
    return " ".join(_s(row.get(c)) for c in (
        "Headline", "Summary", "SummaryAI", "ArticleBody", "article_body",
        "Evidence", "Official Evidence", "Change Type", "Issue", "PolicyFamily",
    )).lower()


def classify_event(row: pd.Series) -> str:
    text = _native_text(row)
    if re.search(ADMIN_PATTERNS, text, re.I):
        return "ADMINISTRATIVE_NOTICE"
    if re.search(OPERATIONAL_PATTERNS, text, re.I):
        return "CUSTOMS_OPERATION"
    return "POLICY_CHANGE"


def official_evidence(row: pd.Series) -> bool:
    content_type = _s(row.get("Content Type")).lower()
    status = _s(row.get("OfficialSourceStatus")).upper()
    return (
        status == "VERIFIED_EXACT"
        or _flag(row, ("EvidenceGateFlag", "OfficialDocumentVerified", "OfficialSourceFlag"))
        or (content_type == "regulation" and _flag(row, ("Body Verified",)))
    )


def samsung_transaction_mapped(row: pd.Series) -> bool:
    entity = _first(row, ("MappedEntity", "SamsungEntity", "Affected Subsidiary"))
    product = _first(row, ("MappedProduct", "Product", "affected_products"))
    hs_code = _first(row, ("MappedHS", "HSCode", "hs_hint"))
    route = _first(row, ("TradeRoute", "ImportExportRoute"))
    mapping_status = _s(row.get("MappingStatus")).upper()
    explicit_gate = _flag(row, ("SamsungTradeGateFlag", "SamsungTransactionMatchedFlag"))
    confirmed_status = mapping_status in {"MAPPED", "ITEM_1TO1_MAPPED", "ENTITY_CONFIRMED", "TRANSACTION_CONFIRMED"}
    # Entity alone or a company name in an article is not a transaction map.
    return bool(entity and (product or hs_code or route) and (explicit_gate or confirmed_status))


def required_action(row: pd.Series, event_type: str) -> str:
    issue = (_s(row.get("Issue")) + " " + _s(row.get("PolicyFamily")) + " " + _s(row.get("Headline"))).lower()
    if event_type == "CUSTOMS_OPERATION":
        return "정지시간과 삼성 납부·신고 예정 건의 중첩 여부를 확인하고, 선납·신고일정 조정 및 통관지연 대상 법인에 공지"
    if "품목분류" in issue or "hs_classification" in issue or "hs code" in issue:
        return "변경 HS 대상 삼성 자재·제품을 확인하고 ERP/GTS 품목분류 마스터, 적용일·경과규정 및 환급 가능 신고 건을 점검"
    if any(x in issue for x in ("ad_cvd", "반덤핑", "덤핑방지", "상계관세", "trade_remedy")):
        return "대상국·공급자·HS별 삼성 수입실적을 대조하고 적용세율·시행일·예치금 및 대체조달 비용을 산출"
    if any(x in issue for x in ("수출통제", "전략물자", "export_control", "export control")):
        return "대상 품목·ECCN/전략물자번호와 삼성 수출경로를 대조하고 허가·최종사용자 증빙 및 출하보류 기준을 점검"
    if any(x in issue for x in ("원산지", "fta", "origin")):
        return "대상 거래의 원산지 기준·CO·BOM·공급자확인서를 대조하고 FTA 적용 및 사후검증 증빙을 점검"
    if any(x in issue for x in ("관세", "tariff", "customs")):
        return "공식 문서의 대상국·품목·세율·시행일을 삼성 법인별 거래실적과 대조하고 관세비용 및 통관조치를 확정"
    return "공식 적용범위와 삼성 법인·품목·거래경로를 대조하고 담당자·기한이 포함된 실행조치를 확정"


def apply_action_queue_contract(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return frame.copy() if frame is not None else pd.DataFrame()
    out = frame.copy()
    for col in ("DecisionStatus", "ReportLayer", "Action Plan", "Policy Stage"):
        if col not in out:
            out[col] = ""

    for idx, row in out.iterrows():
        event_type = classify_event(row)
        evidence_ok = official_evidence(row)
        transaction_ok = samsung_transaction_mapped(row)
        delta_ok = _flag(row, ("PolicyDeltaFlag",)) or _s(row.get("Content Type")).lower() == "regulation"
        proposed = _s(row.get("Policy Stage")).upper() == "PROPOSED_OR_MONITORING" or bool(re.search(PROPOSAL_PATTERNS, _native_text(row), re.I))
        action = required_action(row, event_type)

        out.at[idx, "EventType"] = event_type
        out.at[idx, "OfficialDocumentVerified"] = "Y" if evidence_ok else "N"
        out.at[idx, "SamsungTransactionMapped"] = "Y" if transaction_ok else "N"
        out.at[idx, "RequiredActionDefined"] = "Y" if action else "N"
        out.at[idx, "ActionQueueEligible"] = "N"
        # Never trust a legacy/direct flag from an upstream scoring model.
        # This contract is the sole authority for final direct confirmation.
        out.at[idx, "DirectConfirmedFlag"] = "N"

        if event_type == "ADMINISTRATIVE_NOTICE":
            out.at[idx, "DecisionStatus"] = "Excluded"
            out.at[idx, "ReportLayer"] = "EXCLUDED"
            out.at[idx, "PriorityEligible"] = "N"
            out.at[idx, "ContractReason"] = "채용·입찰·교육·행사 등 행정성 게시물로 관세정책 및 삼성 Action Queue 대상이 아님"
            out.at[idx, "Action Plan"] = ""
            continue

        if event_type == "CUSTOMS_OPERATION":
            out.at[idx, "DecisionStatus"] = "Operational Alert"
            out.at[idx, "ReportLayer"] = "OPERATIONAL_ALERT"
            out.at[idx, "ContractReason"] = "통관·납부 운영변경: 삼성 예정 거래와 시간 중첩이 확인될 때만 Action Queue 진입"
            out.at[idx, "Action Plan"] = action
            overlap_ok = _flag(row, ("ScheduledTransactionOverlapFlag", "PaymentDeadlineOverlapFlag"))
            if evidence_ok and transaction_ok and overlap_ok:
                out.at[idx, "ActionQueueEligible"] = "Y"
                out.at[idx, "PriorityEligible"] = "Y"
            else:
                out.at[idx, "PriorityEligible"] = "N"
            continue

        if not delta_ok:
            out.at[idx, "DecisionStatus"] = "Global Context"
            out.at[idx, "ReportLayer"] = "REFERENCE"
            out.at[idx, "PriorityEligible"] = "N"
            out.at[idx, "ContractReason"] = "현재 발생한 신규 정책변화가 확인되지 않아 Global Context로 관리"
            continue

        out.at[idx, "ReportLayer"] = "POLICY_RADAR"
        out.at[idx, "Action Plan"] = action
        if not evidence_ok:
            out.at[idx, "DecisionStatus"] = "Verification Pending"
            out.at[idx, "PriorityEligible"] = "N"
            out.at[idx, "ContractReason"] = "해당 사건의 공식 문서 원문 확인 필요"
        elif not transaction_ok:
            out.at[idx, "DecisionStatus"] = "Policy Radar"
            out.at[idx, "PriorityEligible"] = "N"
            out.at[idx, "ContractReason"] = "신규 정책과 공식근거는 확인됐으나 삼성 법인·품목·HS·거래경로 연결이 미확인"
        else:
            out.at[idx, "DecisionStatus"] = "Scenario Analysis" if proposed else "Action Required"
            out.at[idx, "ActionQueueEligible"] = "Y"
            out.at[idx, "PriorityEligible"] = "Y"
            out.at[idx, "DirectConfirmedFlag"] = "N" if proposed else "Y"
            out.at[idx, "ContractReason"] = "신규 정책·공식 문서·삼성 거래 연결·실행조치 확인"
    return out
