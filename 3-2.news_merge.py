# -*- coding: utf-8 -*-
# GTI v5.4 STEP3-2 - Balanced Customs/Trade Candidate Engine / Same-day rerun safe
"""
STEP3-2 : news_merge.py

Input:
- C:\\Temp\\1-2.site_news_raw.xlsx   # STEP1 non-LAW1 official/news rows
- C:\\Temp\\2-1.naver_news_raw.xlsx
- C:\\Temp\\2-2.google_news_raw.xlsx
- C:\\Temp\\2-3.rss_news_raw.xlsx

Output:
- C:\\Temp\\3-2.news_summary.xlsx
- C:\\Temp\\3-2.news_cumulative.xlsx

Role:
- 언론/포털/RSS/사이트뉴스 후보 정리
- 법규 원문 전용 로직 제거
- 48시간 기준은 collected_at 우선 적용
- keyword.xlsx 기준 무역/통상/관세 관련 뉴스 선별
- 삼성 영향도와 중복 제거 scoring 강화
- 공공요금 tariff / 여행·스포츠·범죄 등 오탐 필터 강화
- Google Alert/RSS Agency를 실제 기사 URL 도메인 기준으로 보정
- Direct 영향도 과대 산정 완화(생산국가 단독 Direct 방지)
- cumulative는 URL 기준으로만 비교하며 기존 행을 절대 줄이지 않고 신규 행만 추가
"""

from __future__ import annotations

import argparse
import os
import re
import zipfile
import shutil
from difflib import SequenceMatcher
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import parse_qs, parse_qsl, unquote, urlencode, urlparse

import pandas as pd
from gti_report_window import previous_kst_day_mask
import requests