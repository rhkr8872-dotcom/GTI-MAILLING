# -*- coding: utf-8 -*-
"""
GTI STEP4-1 REGULATION AI ANALYSIS - GUARDRAIL v4.1

Fixes
- Exclude stale regulations/notices older than GTI_STEP4_REG_MAX_AGE_DAYS (default 90).
- Exclude webinar/seminar/tender/opening ceremony/event notices.
- Exclude bad URLs such as fonts.googleapis / analytics.
- Do not misread arbitrary percentages as tariff rates.
- Keep only customs/trade/FTA/export-control/CBAM/AD-CVD/HS regulation items.
"""
from __future__ import annotations

import os
import re
import json
import ssl
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from urllib.parse import quote, unquote, urlparse

import pandas as pd
from gti_report_window import previous_kst_day_mask

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_FILE = BASE_DIR / "3-1.regulation_article_summary.xlsx"
FALLBACK_INPUT_FILE = BASE_DIR / "3-1.regulation_summary.xlsx"
KEYWORD_FILE = BASE_DIR / "keyword.xlsx"
OUT_SUMMARY = BASE_DIR / "4-1.regulation_ai_summary.xlsx"
MAX_AGE_HOURS = int(os.getenv("GTI_STEP4_REG_MAX_AGE_HOURS", "24"))
OUT_CUMULATIVE = BASE_DIR / "4-1.regulation_ai_cumulative.xlsx"
OUT_CUMULATIVE_REMOVED = BASE_DIR / "4-1.regulation_ai_cumulative_removed.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-1.regulation_ai_excluded.xlsx"

MAX_AGE_DAYS = int(os.getenv("GTI_STEP4_REG_MAX_AGE_DAYS", "90"))
TOP_N_MAX = int(os.getenv("GTI_STEP4_REG_TOP_N_MAX", "9999"))
MIN_SCORE = int(os.getenv("GTI_STEP4_REG_MIN_SCORE", "70"))
KEYWORD_MIN_LEN = int(os.getenv("GTI_STEP4_REG_KEYWORD_MIN_LEN", "2"))

BAD_URL_PATTERNS = ["google-analytics.com", "googletagmanager.com", "doubleclick.net", "analytics.js", "fonts.googleapis.com", "fonts.gstatic.com", "googleusercontent.com", "googleadservices.com"]