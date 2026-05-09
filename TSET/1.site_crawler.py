# -*- coding: utf-8 -*-
# GTI STEP1 FINAL - sites.xlsx 운영형
# input : C:\temp\sites.xlsx
# output: C:\temp\1.site_news_raw.xlsx
# form  : date / title / url / source / collected_at / agency

import re
import time
import warnings
import pandas as pd
import requests

from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import urlencode, urljoin, urlparse, parse_qs, parse_qsl, urlunparse
from urllib.request import Request, urlopen
from urllib.error import URLError
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

BASE_DIR = Path(r"C:\temp")
SITE_FILE = BASE_DIR / "sites.xlsx"
OUT_FILE = BASE_DIR / "1.site_news_raw.xlsx"
CUM_FILE = BASE_DIR / "1.site_news_cumulative.xlsx"
CANDIDATE_FILE = BASE_DIR / "1.site_news_candidates.xlsx"
RUN_LOG_FILE = BASE_DIR / "1.site_crawler_run_log.txt"

HOURS_BACK = 24
MAX_PER_SITE = 20
SIMILAR_TITLE_THRESHOLD = 0.86
# Core collection rule:
# collect every valid new post from sites.xlsx in the recent 24-hour window.
# Keyword filtering should be used later for classification/analysis, not for excluding site posts.
COLLECT_ALL_SITE_POSTS = True

results = []

BAD_TITLE_CONTAINS = [
    "로그인", "회원가입", "사이트맵", "skip", "menu", "home",
    "privacy", "cookie", "contact", "about us", "accessibility",
    "facebook", "twitter", "youtube", "instagram", "linkedin",
    "검색", "전체메뉴", "본문", "바로가기", "이전", "다음",
    "처음", "마지막", "다운로드", "첨부파일", "자주묻는질문", "faq",
    "네이버 블로그", "블로그", "blog.naver.com",
    "관보보기", "일자별 기간별", "마이페이지", "관심 관보",
    "발행예고보기", "내일관보", "관보분석",
    "등록·채용 신고", "관세사 · 법인 징계현황"
]

BAD_TITLE_EXACT = {
    "", "-", "0", "new", "more", "보기", "상세보기", "검색",
    "공지사항", "보도자료", "고시", "공고", "훈령", "예규",
    "뉴스", "news", "home", "menu"
}

TRADE_WORDS = [
    "관세", "통관", "수입", "수출", "무역", "통상", "고시", "공고",
    "훈령", "예규", "입법예고", "행정예고", "FTA", "원산지",
    "customs", "tariff", "trade", "import", "export", "notice",
    "regulation", "announcement", "directive", "policy",
    "anti-dumping", "antidumping", "countervailing", "safeguard",
    "hs code", "classification", "valuation", "rules of origin",
    "export control", "sanction", "supply chain",
    "arancel", "aduana", "comercio", "importación", "exportación",
    "thuế quan", "hải quan", "thương mại", "nhập khẩu", "xuất khẩu",
    "सीमा शुल्क", "व्यापार", "आयात", "निर्यात"
]

POLICY_WORDS = [
    "고시", "공고", "훈령", "예규", "입법예고", "행정예고", "시행령", "시행규칙",
    "법률", "개정", "규정", "지침", "regulation", "notice", "announcement",
    "directive", "rule", "rules", "amendment", "decree", "ordinance",
    "investigation", "determination", "measure", "policy", "law"
]


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", unescape(str(value))).strip()


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_extra_keywords():
    """Load custom query words when a keyword workbook exists in C:\\Temp."""
    candidates = [
        BASE_DIR / "custom_queries.xlsx",
        BASE_DIR / "keyword.xlsx",
        BASE_DIR / "KEYWORD.xlsx",
        BASE_DIR / "keyword_master.xlsx",
    ]
    words = []