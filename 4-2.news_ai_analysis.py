# -*- coding: utf-8 -*-
"""GTI STEP5 v45 evidence-gated executive report engine.

One preparation path, one quality contract, one send path.  --preview and
--no-email never mutate cumulative history.
"""
from __future__ import annotations

import argparse
import html
import os
import re
import smtplib
import ssl
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd
from gti_quality_contract import apply_quality_contract, VERSION as CONTRACT_VERSION
from gti_report_window import previous_kst_day_mask


BASE = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
OUT_DIR = Path(os.getenv("GTI_OUTPUT_DIR", str(BASE / "12345" / "c_type_outputs")))
NEWS_FILE = BASE / "4-2.news_ai_summary.xlsx"
REG_FILE = BASE / "4-1.regulation_ai_summary.xlsx"
CUM_FILE = BASE / "gti_news_cumulative.xlsx"
RECIPIENT_FILE = BASE / "00.xlsx"
SMTP_HOST = os.getenv("GTI_SMTP_HOST", "smtp.naver.com")
SMTP_PORT = int(os.getenv("GTI_SMTP_PORT", "465"))
SMTP_USER = os.getenv("GTI_SMTP_USER", "kch8872@naver.com").strip()
SMTP_PASS = (os.getenv("GTI_SMTP_PASS") or os.getenv("GTI_MAIL_PW") or "").strip()


def s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()