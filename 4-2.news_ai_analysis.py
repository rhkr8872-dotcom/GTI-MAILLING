# -*- coding: utf-8 -*-
"""
GTI STEP4-2 NEWS AI v47.2 KST PRIOR-DAY QUALITY ENGINE
- Input: 3-2.news_summary.xlsx
- Strict published-date 24h guard
- No legacy v18/v20/v23/v24 override chain
- Gemini: customs/trade YES/NO + Samsung customs impact analysis
- Final: maximum 30 news items
"""

from __future__ import annotations
import os, re, json, time, html as html_lib
from difflib import SequenceMatcher
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse
import pandas as pd
from gti_action_queue_contract import apply_action_queue_contract
import requests
from gti_report_window import previous_kst_day_mask
try:
    from gti_quality_contract import apply_quality_contract, VERSION as CONTRACT_VERSION
except ImportError as exc:
    raise RuntimeError(
        "gti_quality_contract.py must be placed in the same folder as this script"
    ) from exc

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_FILE = BASE_DIR / "3-2.news_summary.xlsx"
REQUIRED_COLLECTOR_FILES = [
    BASE_DIR / "2-1.naver_news_raw.xlsx",
    BASE_DIR / "2-2.google_news_raw.xlsx",
    BASE_DIR / "2-3.rss_news_raw.xlsx",
]
OUT_SUMMARY = BASE_DIR / "4-2.news_ai_summary.xlsx"
OUT_CUMULATIVE = BASE_DIR / "4-2.news_ai_cumulative.xlsx"
OUT_AUDIT = BASE_DIR / "4-2.news_ai_audit_candidates.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-2.news_ai_excluded.xlsx"
OUT_LEGACY = BASE_DIR / "4.news_ai_analysis.xlsx"
MAPPING_MASTER_XLSX = BASE_DIR / "gti_samsung_customs_mapping.xlsx"
MAPPING_MASTER_CSV = BASE_DIR / "gti_samsung_customs_mapping.csv"# -*- coding: utf-8 -*-
"""
GTI STEP4-2 NEWS AI v47.2 KST PRIOR-DAY QUALITY ENGINE
- Input: 3-2.news_summary.xlsx
- Strict published-date 24h guard
- No legacy v18/v20/v23/v24 override chain
- Gemini: customs/trade YES/NO + Samsung customs impact analysis
- Final: maximum 30 news items
"""

from __future__ import annotations
import os, re, json, time, html as html_lib
from difflib import SequenceMatcher
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse
import pandas as pd
from gti_action_queue_contract import apply_action_queue_contract
import requests
from gti_report_window import previous_kst_day_mask
try:
    from gti_quality_contract import apply_quality_contract, VERSION as CONTRACT_VERSION
except ImportError as exc:
    raise RuntimeError(
        "gti_quality_contract.py must be placed in the same folder as this script"
    ) from exc

BASE_DIR = Path(os.getenv("GTI_BASE_DIR", r"C:\Temp"))
INPUT_FILE = BASE_DIR / "3-2.news_summary.xlsx"
REQUIRED_COLLECTOR_FILES = [
    BASE_DIR / "2-1.naver_news_raw.xlsx",
    BASE_DIR / "2-2.google_news_raw.xlsx",
    BASE_DIR / "2-3.rss_news_raw.xlsx",
]
OUT_SUMMARY = BASE_DIR / "4-2.news_ai_summary.xlsx"
OUT_CUMULATIVE = BASE_DIR / "4-2.news_ai_cumulative.xlsx"
OUT_AUDIT = BASE_DIR / "4-2.news_ai_audit_candidates.xlsx"
OUT_EXCLUDED = BASE_DIR / "4-2.news_ai_excluded.xlsx"
OUT_LEGACY = BASE_DIR / "4.news_ai_analysis.xlsx"
MAPPING_MASTER_XLSX = BASE_DIR / "gti_samsung_customs_mapping.xlsx"
MAPPING_MASTER_CSV = BASE_DIR / "gti_samsung_customs_mapping.csv"