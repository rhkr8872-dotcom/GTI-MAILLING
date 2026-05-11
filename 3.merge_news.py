from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "output"

OUT_DIR.mkdir(exist_ok=True)

SITE_FILE = DATA_DIR / "site.xlsx"
KEYWORD_FILE = DATA_DIR / "keyword.xlsx"
MAIL_FILE = DATA_DIR / "00.xlsx"

RAW1 = OUT_DIR / "1.site_news_raw.xlsx"
RAW2 = OUT_DIR / "2-1.naver_news_raw.xlsx"
RAW3 = OUT_DIR / "2-2.google_news_raw.xlsx"
RAW4 = OUT_DIR / "2-3.rss_news_raw.xlsx"

MERGED = OUT_DIR / "news_ai_summary.xlsx"
FINAL = OUT_DIR / "news_raw.xlsx"


import os
import re
import pandas as pd
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from pathlib import Path

BASE_DIR = Path(".")

INPUT_FILES = [
    "1.site_news_raw.xlsx",
    "2-1.naver_news_raw.xlsx",
    "2-2.google_news_raw.xlsx",
    "2-3.rss_news_raw.xlsx"
]

KEYWORD_CANDIDATES = [
    "keyword_master_trade_policy_150.xlsx",
    "KEYWORD.xlsx",
    "keyword.xlsx"
]

RAW_FILE = os.path.join(BASE_DIR, "3.news_master_raw.xlsx")
SUMMARY_FILE = os.path.join(BASE_DIR, "3.news_ai_summary.xlsx")
CUMULATIVE_FILE = os.path.join(BASE_DIR, "3.news_ai_cumulative.xlsx")

MAX_OUTPUT = 200
RECENT_HOURS = 24
TITLE_SIM_THRESHOLD = 0.86
COSINE_SIM_THRESHOLD = 0.78
CUMULATIVE_SIM_THRESHOLD = 0.92


def load_data():
    dfs = []
    for file_name in INPUT_FILES:
        path = os.path.join(BASE_DIR, file_name)
        if os.path.exists(path):
            df = pd.read_excel(path)
            df["source_file"] = file_name
            dfs.append(df)
            print(f"load: {file_name} / rows={len(df)}")

    if not dfs:
        raise Exception("No input files found")

    return pd.concat(dfs, ignore_index=True)


def load_keywords():
    keywords = []

    for file_name in KEYWORD_CANDIDATES:
        path = os.path.join(BASE_DIR, file_name)
        if not os.path.exists(path):
            continue

        df = pd.read_excel(path)
        df.columns = [str(c).strip().lower() for c in df.columns]

        for col in df.columns:
            if "keyword" in col:
                keywords = df[col].dropna().astype(str).str.strip().str.lower().tolist()
                print(f"keyword file loaded: {file_name} / keywords={len(keywords)}")
                break

        if keywords:
            break

    base_keywords = [
        "tariff", "customs", "duty", "trade", "export", "import",
        "fta", "anti-dumping", "countervailing", "sanction",
        "restriction", "regulation", "supply chain", "export control",
        "import restriction", "hs code", "de minimis", "wto",

        "관세", "통상", "무역", "수출", "수입", "수출입",
        "fta", "반덤핑", "상계관세", "수입규제", "통관",
        "규제", "공급망", "수출통제", "무역협정", "관세율",
        "보복관세", "비관세", "원산지", "세관", "hs코드",
        "품목분류", "품목번호", "기업무역활동", "수출입 현황",
        "관세청", "국제관세협력", "심사", "조사", "분류원"
    ]

    return list(set(keywords + base_keywords))


def standardize(df):
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]

    for col in ["date", "title", "url", "source", "collected_at", "source_file"]:
        if col not in df.columns:
            df[col] = ""

    df["title"] = df["title"].astype(str).str.strip()
    df["url"] = df["url"].astype(str).str.strip()
    df["source"] = df["source"].astype(str).str.strip()
    df["source_file"] = df["source_file"].astype(str).str.strip()

    df["date"] = pd.to_datetime(df["date"], errors="coerce", utc=True)
    df["date"] = df["date"].dt.tz_convert(None)

    df = df[df["title"].astype(str).str.len() > 5]

    return df.reset_index(drop=True)


def filter_recent(df):
    cutoff = datetime.now() - timedelta(hours=RECENT_HOURS)
    return df[df["date"] >= cutoff].reset_index(drop=True)


def is_must_keep(title, source=""):
    t = str(title).lower()
    s = str(source).lower()

    must_keep_keywords = [
        "관세청", "customs", "korea customs",
        "[국제관세협력]", "[통관]", "[심사]", "[조사]",
        "[정보데이터]", "[분류원]", "[품목분류]",
        "품목분류", "품목번호", "관세율", "수출입 현황",
        "기업무역활동", "통계 공표", "수출입기업",
        "대미 수출", "철강제품", "품목관세",
        "지식재산권", "위해물품", "국제관세협력"
    ]

    return any(k.lower() in t or k.lower() in s for k in must_keep_keywords)


def remove_real_noise(df):
    noise_keywords = [
        "youtube", "facebook", "instagram", "tiktok", "reddit",
        "threads", "shorts", "reels",

        "stock", "stocks", "share price", "shares", "nasdaq", "nyse",
        "brokerage", "brokerages", "consensus recommendation",
        "eps", "dividend", "earnings call", "marketbeat", "simplywall",
        "zacks", "analyst rating", "price target",
        "주가", "증권", "투자", "실적", "배당", "목표주가", "매수", "매도",

        "baseball", "softball", "lacrosse", "basketball", "football",
        "tournament", "playoffs", "roundup", "high school",
        "hs baseball", "hs softball", "hs roundup",
        "야구", "축구", "농구", "대회", "토너먼트",

        "celebrity", "movie", "music", "entertainment", "fashion",
        "shopping", "review", "coupon", "sale", "deal",
        "연예", "방송", "영화", "음악", "패션", "쇼핑", "광고",
        "홍보", "이벤트", "맛집",

        "strike", "union", "wage", "inheritance", "restructuring",
        "hiring", "election", "opinion", "editorial", "column",
        "파업", "노조", "임금", "상속", "구조조정", "채용",
        "노동", "선거", "정치", "칼럼", "사설",

        "peace pole", "donate", "donation", "festival",
        "행사", "축제", "기부"
    ]

    keep_rows = []

    for idx, row in df.iterrows():
        title = str(row.get("title", ""))
        source = str(row.get("source", ""))

        if is_must_keep(title, source):
            keep_rows.append(idx)
            continue

        t = title.lower()

        if any(k in t for k in noise_keywords):
            continue

        keep_rows.append(idx)

    return df.loc[keep_rows].reset_index(drop=True)


def policy_filter(df, keywords):
    keep_rows = []

    for idx, row in df.iterrows():
        title = str(row.get("title", "")).lower()
        source = str(row.get("source", "")).lower()
        text = f"{title} {source}"

        if is_must_keep(title, source):
            keep_rows.append(idx)
            continue

        if any(k in text for k in keywords):
            keep_rows.append(idx)

    return df.loc[keep_rows].reset_index(drop=True)


def normalize_title(title):
    t = str(title).lower()

    t = re.sub(r"&quot;|&#39;|&amp;", " ", t)

    for sep in [" - ", " | ", " : ", " — ", " – "]:
        if sep in t:
            t = t.split(sep)[0]

    t = re.sub(r"[^a-z0-9가-힣一-龥ぁ-ゔァ-ヴー\s]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()

    remove_words = [
        "the", "a", "an", "new", "latest", "update", "updates",
        "said", "says", "may", "could", "would", "will",
        "breaking", "exclusive", "report"
    ]

    tokens = [w for w in t.split() if w not in remove_words]
    return " ".join(tokens).strip()


def dedup_exact(df):
    df["title_norm"] = df["title"].apply(normalize_title)

    before = len(df)
    df = df.drop_duplicates(subset=["url"], keep="first")
    print("URL dedup 제거:", before - len(df))

    before = len(df)
    df = df.drop_duplicates(subset=["title_norm"], keep="first")
    print("Title exact dedup 제거:", before - len(df))

    return df.reset_index(drop=True)


def is_similar(a, b, threshold=TITLE_SIM_THRESHOLD):
    a = str(a)
    b = str(b)

    if len(a) < 8 or len(b) < 8:
        return False

    return SequenceMatcher(None, a, b).ratio() >= threshold


def dedup_sentence_similarity(df):
    seen_titles = []
    keep_rows = []

    for idx, row in df.iterrows():
        title_norm = str(row.get("title_norm", ""))

        duplicate = False

        for seen in seen_titles[-500:]:
            if is_similar(title_norm, seen):
                duplicate = True
                break

            if len(title_norm) >= 24 and title_norm[:24] == seen[:24]:
                duplicate = True
                break

        if not duplicate:
            seen_titles.append(title_norm)
            keep_rows.append(idx)

    print("Sentence similar dedup 제거:", len(df) - len(keep_rows))
    return df.loc[keep_rows].reset_index(drop=True)


def dedup_cosine(df):
    if len(df) < 2:
        return df

    titles = df["title_norm"].fillna("").astype(str).tolist()

    try:
        vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5))
        tfidf = vectorizer.fit_transform(titles)
        sim_matrix = cosine_similarity(tfidf)
    except Exception as e:
        print(f"Cosine dedup skip: {e}")
        return df

    keep = []
    removed = set()

    for i in range(len(titles)):
        if i in removed:
            continue

        keep.append(i)

        for j in range(i + 1, len(titles)):
            if j in removed:
                continue

            if sim_matrix[i, j] >= COSINE_SIM_THRESHOLD:
                removed.add(j)

    print("Cosine dedup 제거:", len(df) - len(keep))
    return df.iloc[keep].reset_index(drop=True)


def remove_cumulative(df):
    if not os.path.exists(CUMULATIVE_FILE):
        print("cumulative 없음 → skip")
        return df

    old = pd.read_excel(CUMULATIVE_FILE)
    old = standardize(old)

    if len(old) == 0:
        print("cumulative empty → skip")
        return df

    old["title_norm"] = old["title"].apply(normalize_title)

    before = len(df)

    old_urls = set(old["url"].dropna().astype(str))
    old_titles = old["title_norm"].dropna().astype(str).tolist()

    keep_rows = []

    for idx, row in df.iterrows():
        url = str(row.get("url", ""))
        title_norm = str(row.get("title_norm", ""))

        if url and url in old_urls:
            continue

        duplicate = False
        for old_title in old_titles[-3000:]:
            if is_similar(title_norm, old_title, CUMULATIVE_SIM_THRESHOLD):
                duplicate = True
                break

        if not duplicate:
            keep_rows.append(idx)

    df = df.loc[keep_rows].reset_index(drop=True)

    print("cumulative 제거:", before - len(df))
    return df


def update_cumulative(df):
    save_df = df.copy()

    if os.path.exists(CUMULATIVE_FILE):
        old = pd.read_excel(CUMULATIVE_FILE)
        save_df = pd.concat([old, save_df], ignore_index=True)

    save_df.columns = [str(c).strip().lower() for c in save_df.columns]

    if "url" in save_df.columns:
        save_df = save_df.drop_duplicates(subset=["url"], keep="first")

    save_df.to_excel(CUMULATIVE_FILE, index=False)


def samsung_policy_score(title):
    t = str(title).lower()
    score = 0

    policy_high = [
        "tariff", "customs", "duty", "anti-dumping", "countervailing",
        "sanction", "restriction", "export control", "import restriction",
        "관세", "통관", "반덤핑", "상계관세", "수입규제", "수출통제",
        "품목분류", "품목번호", "관세율", "심사"
    ]

    policy_mid = [
        "trade", "export", "import", "fta", "regulation", "supply chain",
        "무역", "수출", "수입", "수출입", "통상", "규제", "공급망",
        "원산지", "기업무역활동", "수출입 현황"
    ]

    products = [
        "mobile", "smartphone", "phone",
        "consumer electronics", "tv", "appliance",
        "network", "telecom",
        "medical", "healthcare",
        "semiconductor", "chip", "memory",
        "반도체", "스마트폰", "모바일", "가전", "네트워크", "의료기기"
    ]

    production_sites = [
        "korea", "china", "vietnam", "india", "indonesia",
        "turkey", "slovakia", "poland", "mexico", "brazil",
        "한국", "중국", "베트남", "인도", "인도네시아",
        "튀르키예", "터키", "슬로바키아", "폴란드", "멕시코", "브라질"
    ]

    if any(k in t for k in policy_high):
        score += 5
    elif any(k in t for k in policy_mid):
        score += 3

    if any(k in t for k in products):
        score += 4

    if any(k in t for k in production_sites):
        score += 3

    if is_must_keep(t):
        score += 5

    return score


def add_score(df):
    df["score"] = df["title"].apply(samsung_policy_score)
    return df.sort_values(["score", "date"], ascending=[False, False]).reset_index(drop=True)


def filter_recent_relaxed(df):
    cutoff = datetime.now() - timedelta(hours=36)
    return df[df["date"] >= cutoff].reset_index(drop=True)


def fallback_if_zero_or_too_low(original_df, current_df, keywords, min_rows=50):
    if len(current_df) >= min_rows:
        return current_df

    print(f"fallback 작동: 현재 {len(current_df)}건 → 최소 {min_rows}건 확보 시도")

    df = original_df.copy()
    df = filter_recent_relaxed(df)
    df = remove_real_noise(df)
    df = policy_filter(df, keywords)
    df = dedup_exact(df)
    df = dedup_sentence_similarity(df)
    df = dedup_cosine(df)
    df = add_score(df)

    combined = pd.concat([current_df, df], ignore_index=True)
    combined = dedup_exact(combined)
    combined = dedup_sentence_similarity(combined)
    combined = add_score(combined)

    return combined


def main():
    print("STEP3 FINAL - POLICY CANDIDATE ENGINE v2")

    df = load_data()
    print("Loaded:", len(df))

    df = standardize(df)
    df.to_excel(RAW_FILE, index=False)

    original_df = df.copy()

    df = filter_recent(df)
    print("24h:", len(df))

    keywords = load_keywords()

    df = remove_real_noise(df)
    print("real noise 제거:", len(df))

    df = policy_filter(df, keywords)
    print("policy:", len(df))

    df = dedup_exact(df)
    print("exact dedup 후:", len(df))

    df = dedup_sentence_similarity(df)
    print("sentence dedup 후:", len(df))

    df = dedup_cosine(df)
    print("cosine dedup 후:", len(df))

    df = remove_cumulative(df)
    print("cumulative 후:", len(df))

    df = add_score(df)

    df = fallback_if_zero_or_too_low(original_df, df, keywords, min_rows=50)
    print("fallback 후:", len(df))

    df = add_score(df)
    df = df.head(MAX_OUTPUT)

    df.to_excel(SUMMARY_FILE, index=False)
    update_cumulative(df)

    print("STEP3 COMPLETE:", len(df))
    print("SAVE:", SUMMARY_FILE)


if __name__ == "__main__":
    main()
