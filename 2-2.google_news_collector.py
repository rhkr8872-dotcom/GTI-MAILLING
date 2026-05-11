# 2-2.google_news_collector.py

# GitHub / Windows Compatible FINAL

from pathlib import Path
import pandas as pd
import feedparser
from datetime import datetime

print("🚀 GTI STEP2 GOOGLE RSS START v2.1")

# =====================================================

# PATH CONFIG

# =====================================================

BASE_DIR = Path(**file**).resolve().parent

DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"

DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

KEYWORD_FILE = DATA_DIR / "keyword.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "2-2.google_news_raw.xlsx"

# =====================================================

# LOAD KEYWORDS

# =====================================================

if not KEYWORD_FILE.exists():
raise FileNotFoundError(f"❌ keyword file not found: {KEYWORD_FILE}")

kw_df = pd.read_excel(KEYWORD_FILE)

keyword_col = None

for c in kw_df.columns:
if str(c).strip().lower() in [
"keyword",
"keywords",
"키워드"
]:
keyword_col = c
break

if keyword_col is None:
keyword_col = kw_df.columns[0]

keywords = (
kw_df[keyword_col]
.dropna()
.astype(str)
.str.strip()
.unique()
.tolist()
)

print(f"✅ Keywords Loaded: {len(keywords)}")

# =====================================================

# GOOGLE RSS COLLECT

# =====================================================

rows = []

for kw in keywords:

```
try:

    rss_url = (
        "https://news.google.com/rss/search?q="
        + kw
        + "&hl=ko&gl=KR&ceid=KR:ko"
    )

    feed = feedparser.parse(rss_url)

    print(f"🔍 {kw} → {len(feed.entries)}")

    for e in feed.entries:

        rows.append({
            "date": getattr(e, "published", ""),
            "title": getattr(e, "title", ""),
            "url": getattr(e, "link", ""),
            "source": "GoogleNews",
            "keyword": kw,
            "collected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

except Exception as ex:
    print(f"❌ ERROR {kw}: {ex}")
```

# =====================================================

# SAVE

# =====================================================

df = pd.DataFrame(rows)

if len(df) == 0:
print("❌ No Google news collected")
else:

```
df = df.drop_duplicates(
    subset=["title"]
).reset_index(drop=True)

df.to_excel(OUTPUT_FILE, index=False)

print(f"✅ SAVED: {OUTPUT_FILE}")
print(f"✅ ROWS : {len(df)}")
```

print("🏁 GTI STEP2 GOOGLE RSS END")
