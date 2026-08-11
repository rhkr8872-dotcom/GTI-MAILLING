# =========================================================
# GTI STEP2-2 - GOOGLE NEWS RAW + ORIGINAL URL PARALLEL FINAL v3.4
# 목적: 빠른 수집 + Google News 원문 URL 병렬 복구 + 실행시간 로그
# 원칙: url 컬럼은 원문 URL 우선, google_url 컬럼은 Google News 원본 URL 보존
# =========================================================

import os
import re
import json
import time
import pandas as pd
import feedparser
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from urllib.parse import quote, unquote, urlparse, parse_qs
from bs4 import BeautifulSoup

print("🚀 GTI v5.0 STEP2-2 GOOGLE BROAD NEWS + ORIGINAL URL START")

# =============================
# PATH / CONFIG
# =============================
BASE_PATH = "C:\\temp\\"
KEYWORD_FILE = BASE_PATH + "keyword.xlsx"
RAW_FILE = BASE_PATH + "2-2.google_news_raw.xlsx"

LOOKBACK_HOURS = int(os.getenv("GTI_LOOKBACK_HOURS", "72"))
SLEEP_SEC = 0.05

# 원문 URL 복구 옵션
# - Y: 2-2 수집 단계에서 Google News URL을 병렬로 원문 URL로 복구
# - N: 기존 FAST 방식처럼 Google URL만 저장
ENABLE_ORIGINAL_URL_RESOLVE = os.getenv("GTI_STEP2_RESOLVE_ORIGINAL_URL", "Y").strip().upper() != "N"
URL_RESOLVE_WORKERS = int(os.getenv("GTI_STEP2_URL_WORKERS", "30"))
URL_RESOLVE_TIMEOUT = int(os.getenv("GTI_STEP2_URL_TIMEOUT", "8"))
URL_RESOLVE_RETRY = int(os.getenv("GTI_STEP2_URL_RETRY", "1"))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}



def safe_to_excel(df, output_path, index=False):
    """Excel 저장 안전 함수.
    - 기존 xlsx가 Excel에서 열려 있어 PermissionError가 나면 timestamp 백업 파일로 저장
    - Windows 파일 잠금 때문에 전체 수집 성공 후 마지막 저장에서 실패하는 문제 방지
    """
    try:
        df.to_excel(output_path, index=index)
        return output_path, "OK"
    except PermissionError:
        base, ext = os.path.splitext(output_path)
        alt_path = f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        df.to_excel(alt_path, index=index)
        return alt_path, "PERMISSION_LOCKED_SAVED_AS_ALT"
    except OSError as e:
        # Excel/OneDrive/백신 잠금 등 Windows 저장 오류 대비
        base, ext = os.path.splitext(output_path)
        alt_path = f"{base}_{datetime.now():%Y%m%d_%H%M%S}{ext}"
        try:
            df.to_excel(alt_path, index=index)
            return alt_path, f"OSERROR_SAVED_AS_ALT:{type(e).__name__}"
        except Exception:
            raise

FINAL_COLS = [
    "date",
    "title",
    "url",
    "google_url",
    "source",
    "summary",
    "collected_at",
    "keyword",
    "language",
    "publisher",
    "category",
    "importance",
    "importance_score",
    "score_reason",
    "url_decode_status",
    "original_url_candidate",
    "rss_url",
]

# =============================
# UTILS
# =============================

def clean_html(text):
    if pd.isna(text):
        return ""
    soup = BeautifulSoup(str(text), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def normalize_text(text):
    return re.sub(r"\s+", " ", clean_html(text)).lower().strip()


def contains_any(text, terms):
    t = normalize_text(text)
    return any(str(term).lower() in t for term in terms if str(term).strip())


def keyword_equals_any(keyword, terms):
    k = normalize_text(keyword)
    return any(k == normalize_text(term) for term in terms)


def parse_datetime(entry):
    for key in ["published_parsed", "updated_parsed"]:
        try:
            v = getattr(entry, key, None)
            if v:
                return datetime(*v[:6])
        except Exception:
            pass
    return datetime.now()


def is_recent(dt):
    return dt >= datetime.now() - timedelta(hours=LOOKBACK_HOURS)


def normalize_title(title):
    title = clean_html(title).lower()

    # Google News title usually: "Article title - Publisher"
    # dedup 목적상 뒤 publisher 제거
    if " - " in title:
        title = title.rsplit(" - ", 1)[0]

    title = re.sub(r"[^0-9a-z가-힣一-龥ぁ-ゔァ-ヴー\s]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def get_google_locale(language):
    lang = str(language).upper().strip()
    locale_map = {
        "EN": {"hl": "en", "gl": "US", "ceid": "US:en"},
        "KR": {"hl": "ko", "gl": "KR", "ceid": "KR:ko"},
        "CN": {"hl": "zh-CN", "gl": "CN", "ceid": "CN:zh-Hans"},
        "ES": {"hl": "es", "gl": "ES", "ceid": "ES:es"},
        "PT": {"hl": "pt", "gl": "BR", "ceid": "BR:pt-419"},
        "TR": {"hl": "tr", "gl": "TR", "ceid": "TR:tr"},
        "VI": {"hl": "vi", "gl": "VN", "ceid": "VN:vi"},
        "HI": {"hl": "hi", "gl": "IN", "ceid": "IN:hi"},
    }
    return locale_map.get(lang, locale_map["EN"])


def importance_to_score(v):
    s = str(v).strip().upper()
    if s in ["100", "HIGH", "H", "상", "중요", "A"]:
        return 100
    if s in ["70", "80", "MEDIUM", "M", "중", "B"]:
        return 70
    if s in ["50", "LOW", "L", "하", "C"]:
        return 50
    try:
        return int(float(s))
    except Exception:
        return 50


def extract_publisher(entry, title):
    try:
        src = entry.get("source", {})
        if isinstance(src, dict):
            src_title = src.get("title", "")
            if src_title:
                return clean_html(src_title)
    except Exception:
        pass

    try:
        if " - " in str(title):
            return str(title).rsplit(" - ", 1)[-1].strip()
    except Exception:
        pass

    return ""


# =============================
# SCORE / URL HINT
# =============================

STRONG_KEEP_TERMS = [
    "tariff", "tariffs", "customs duty", "customs duties", "customs clearance",
    "forced labor", "uflpa", "section 301", "section 232", "export control",
    "export controls", "entity list", "denied persons", "anti-dumping",
    "antidumping", "countervailing", "countervailing duty", "ad/cvd",
    "관세", "통관", "수출통제", "수출 통제", "전략물자", "강제노동",
    "반덤핑", "덤핑방지관세", "상계관세", "무역구제", "세이프가드",
]

WEAK_SINGLE_KEYWORDS = [
    "epa", "sta", "수입", "customs",
]

WEAK_SINGLE_SUPPORT_CONTEXT = [
    "tariff", "customs duty", "trade agreement", "economic partnership agreement",
    "fta", "origin", "rules of origin", "export control", "forced labor",
    "section 301", "section 232", "ad/cvd", "anti-dumping", "countervailing",
    "관세", "통관", "원산지", "자유무역협정", "경제동반자협정", "수출통제",
    "반덤핑", "상계관세", "무역구제",
]

TRADE_REMEDY_TERMS = [
    "anti-dumping", "antidumping", "countervailing", "countervailing duty",
    "ad/cvd", "trade remedy", "safeguard", "반덤핑", "덤핑방지관세",
    "상계관세", "무역구제", "세이프가드",
]

GENERIC_IMPORT_NOISE_TERMS = [
    "import car", "imported car", "import beer", "imported beer", "import food",
    "import price", "import prices", "luxury import", "수입차", "수입맥주",
    "수입식품", "수입물가", "수입 가격", "병행수입", "수입 브랜드",
]


def adjust_importance_score(keyword, title, summary, base_score):
    text = f"{title} {summary}"
    score = int(base_score)
    reasons = []

    strong = contains_any(text, STRONG_KEEP_TERMS)
    trade_remedy = contains_any(keyword, TRADE_REMEDY_TERMS) or contains_any(text, TRADE_REMEDY_TERMS)
    weak_single = keyword_equals_any(keyword, WEAK_SINGLE_KEYWORDS)
    weak_supported = contains_any(text, WEAK_SINGLE_SUPPORT_CONTEXT)

    if strong:
        score += 25
        reasons.append("strong_trade_policy_context")

    if trade_remedy:
        score += 35
        reasons.append("trade_remedy_forced_high")

    if weak_single and not weak_supported:
        score -= 35
        reasons.append("weak_single_keyword_penalty")

    if keyword_equals_any(keyword, ["수입"]) and contains_any(text, GENERIC_IMPORT_NOISE_TERMS) and not strong:
        score -= 35
        reasons.append("generic_import_noise_penalty")

    return max(0, min(score, 150)), ", ".join(reasons) or "base"


def extract_original_url_candidate(entry, google_url):
    """
    Google RSS URLs are often encoded article links. STEP2 keeps the Google URL,
    but this stores cheap hints so STEP3 can attempt stronger restoration first.
    """
    candidates = []

    for key in ["id", "guid", "link"]:
        value = str(entry.get(key, "") or "").strip()
        if value:
            candidates.append(value)

    try:
        for link in entry.get("links", []) or []:
            href = str(link.get("href", "") or "").strip()
            if href:
                candidates.append(href)
    except Exception:
        pass

    for value in candidates:
        parsed = urlparse(value)
        qs = parse_qs(parsed.query)
        for param in ["url", "u", "q"]:
            if param in qs and qs[param]:
                decoded = unquote(str(qs[param][0]))
                if decoded.startswith("http") and "news.google." not in decoded:
                    return decoded

    if google_url and "news.google." not in google_url:
        return google_url

    return ""


def elapsed_text(seconds):
    seconds = float(seconds)
    if seconds < 60:
        return f"{seconds:.1f}초"
    return f"{seconds/60:.1f}분"


def is_google_news_url(value):
    u = str(value or "").lower().strip()
    return "news.google.com/rss/articles/" in u or "news.google.com/articles/" in u


def is_bad_original_url(value):
    u = str(value or "").lower().strip()
    if not u or u in {"nan", "none", "null", "-"}:
        return True
    if not (u.startswith("http://") or u.startswith("https://")):
        return True
    bad_patterns = [
        "news.google.com/rss/articles/", "news.google.com/articles/", "news.google.com/",
        "googleusercontent.com", "gstatic.com", "ggpht.com",
        # Google News article pages contain analytics/script URLs. These are NOT article originals.
        "google-analytics.com", "googletagmanager.com", "doubleclick.net",
        "google.com/analytics", "analytics.js", "gtag/js", "googlesyndication.com",
        "googleadservices.com", "googleapis.com", "google.com/pagead",
    ]
    if any(x in u for x in bad_patterns):
        return True
    if re.search(r"\.(png|jpg|jpeg|gif|webp|svg)(\?|$)", u):
        return True
    return False


def _google_news_article_id(value):
    try:
        parsed = urlparse(str(value or ""))
        parts = [p for p in parsed.path.split("/") if p]
        return parts[-1] if parts else ""
    except Exception:
        return ""


def _extract_article_url_from_google_text(text):
    if not text:
        return ""
    variants = [text]
    try:
        variants.append(text.encode("utf-8", errors="ignore").decode("unicode_escape", errors="ignore"))  # harmless decode fallback
    except Exception:
        pass

    patterns = [
        r"data-n-au=[\"'](https?://[^\"']+)[\"']",
        r"data-url=[\"'](https?://[^\"']+)[\"']",
        r"href=[\"'](https?://[^\"']+)[\"']",
        r"url=(https?%3A%2F%2F[^&\"'<>]+)",
        r"(https?:\\/\\/[^\"'<>\\]+)",
        r"(https?://[^\"'<>\s]+)",
    ]
    for t in variants:
        for pat in patterns:
            for m in re.finditer(pat, t, flags=re.I):
                cand = unquote(m.group(1).replace("\\/", "/")).rstrip(".,;?)\"")
                if cand and not is_bad_original_url(cand):
                    return cand
    return ""


def _decode_google_news_batchexecute(article_id, page_text):
    if not article_id or not page_text:
        return ""
    sg = ""
    ts = ""
    m = re.search(r'data-n-a-sg=["\']([^"\']+)["\']', page_text)
    if m:
        sg = m.group(1)
    m = re.search(r'data-n-a-ts=["\']([^"\']+)["\']', page_text)
    if m:
        ts = m.group(1)
    if not sg or not ts:
        return ""
    try:
        req_obj = [
            "garturlreq",
            [["en-US", "US", ["FINANCE_TOP_INDICES", "WEB_TEST_1_0_0"], None, None, 1, 1, "US:en", None, 180, None, None, None, None, None, 0, None, None, [int(ts), 0]],
             "en-US", "US", 1, [2, 3, 4, 8], 1, 0, "655000234", 0, 0, None, 0],
            article_id,
            int(ts),
            sg,
        ]
        f_req = [[[
            "Fbv4je",
            json.dumps(req_obj, ensure_ascii=False, separators=(",", ":")),
            None,
            "generic",
        ]]]
        resp = requests.post(
            "https://news.google.com/_/DotsSplashUi/data/batchexecute?rpcids=Fbv4je",
            data={"f.req": json.dumps(f_req, ensure_ascii=False, separators=(",", ":"))},
            headers={**HEADERS, "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8", "Referer": "https://news.google.com/"},
            timeout=URL_RESOLVE_TIMEOUT,
        )
        if resp.status_code == 200:
            return _extract_article_url_from_google_text(resp.text)
    except Exception:
        return ""
    return ""


def resolve_google_news_original_url(google_url):
    """Google News URL 1건을 원문 URL로 복구한다.

    우선순위:
    1) requests redirect 최종 URL
    2) Google News article page HTML 안의 원문 URL
    3) batchexecute 방식 decode
    """
    u = str(google_url or "").strip()
    if not u:
        return "", "EMPTY_URL"
    if not is_google_news_url(u):
        if not is_bad_original_url(u):
            return u, "OK_ALREADY_ORIGINAL"
        return "", "BAD_NON_GOOGLE_URL"

    last_error = ""
    for attempt in range(URL_RESOLVE_RETRY + 1):
        try:
            resp = requests.get(u, headers=HEADERS, allow_redirects=True, timeout=URL_RESOLVE_TIMEOUT)
            final_url = str(resp.url or "").strip()
            if final_url and not is_bad_original_url(final_url):
                return final_url, "OK_REDIRECT"

            page_text = (resp.text or "")[:500000]
            found = _extract_article_url_from_google_text(page_text)
            if found and not is_bad_original_url(found):
                return found, "OK_HTML"

            found = _decode_google_news_batchexecute(_google_news_article_id(u), page_text)
            if found and not is_bad_original_url(found):
                return found, "OK_BATCHEXECUTE"

            last_error = "STILL_GOOGLE_OR_NO_URL"
        except Exception as e:
            last_error = f"ERROR_{type(e).__name__}"
            time.sleep(0.2)

    return "", last_error or "FAILED"


def resolve_original_urls_parallel(df):
    if df.empty:
        return df
    if not ENABLE_ORIGINAL_URL_RESOLVE:
        df["url_decode_status"] = "SKIPPED_STEP2_FAST"
        return df

    started = time.perf_counter()

    # 이미 cheap hint로 원문 후보가 있는 건은 우선 사용
    df["original_url_candidate"] = df.get("original_url_candidate", "").fillna("").astype(str).str.strip()
    has_hint = df["original_url_candidate"].apply(lambda x: bool(x) and not is_bad_original_url(x))
    df.loc[has_hint, "url"] = df.loc[has_hint, "original_url_candidate"]
    df.loc[has_hint, "url_decode_status"] = "OK_HINT"

    need_mask = ~has_hint & df["google_url"].astype(str).apply(is_google_news_url)
    urls = df.loc[need_mask, "google_url"].dropna().astype(str).str.strip().unique().tolist()

    print(f"🔗 URL RESOLVE START: target={len(urls)} / workers={URL_RESOLVE_WORKERS} / timeout={URL_RESOLVE_TIMEOUT}s")
    if not urls:
        print("🔗 URL RESOLVE SKIP: no Google News URL target")
        return df

    results = {}
    done = 0
    ok = 0
    with ThreadPoolExecutor(max_workers=URL_RESOLVE_WORKERS) as executor:
        future_map = {executor.submit(resolve_google_news_original_url, url): url for url in urls}
        for future in as_completed(future_map):
            url = future_map[future]
            try:
                original_url, status = future.result()
            except Exception as e:
                original_url, status = "", f"ERROR_{type(e).__name__}"
            results[url] = (original_url, status)
            done += 1
            if original_url:
                ok += 1
            if done % 100 == 0 or done == len(urls):
                print(f"   - URL RESOLVE progress: {done}/{len(urls)} / success={ok}")

    for idx, row in df.loc[need_mask].iterrows():
        google_url = str(row.get("google_url", "")).strip()
        original_url, status = results.get(google_url, ("", "NOT_TRIED"))
        if original_url and not is_bad_original_url(original_url):
            df.at[idx, "original_url_candidate"] = original_url
            df.at[idx, "url"] = original_url
            df.at[idx, "url_decode_status"] = status
        else:
            # 실패 시 Google URL은 google_url 컬럼에 보존하고, url도 fallback으로 유지
            df.at[idx, "url"] = google_url
            df.at[idx, "url_decode_status"] = status or "FAILED"

    elapsed = time.perf_counter() - started
    total = len(urls)
    success_rate = (ok / total * 100) if total else 0
    print(f"🔗 URL RESOLVE DONE: target={total}, success={ok}, fail={total-ok}, success_rate={success_rate:.1f}%, elapsed={elapsed_text(elapsed)}")
    return df

# =============================
# KEYWORD LOAD
# =============================

def load_keywords():
    keywords = pd.read_excel(KEYWORD_FILE)
    keywords.columns = [str(c).strip().lower() for c in keywords.columns]

    required_cols = ["keyword", "language", "category", "importance", "active"]
    for col in required_cols:
        if col not in keywords.columns:
            raise Exception(f"❌ KEYWORD 파일 필수 컬럼 없음: {col}")

    keywords = keywords[keywords["active"].astype(str).str.upper().str.strip() == "Y"]
    keywords = keywords.dropna(subset=["keyword"])
    keywords["keyword"] = keywords["keyword"].astype(str).str.strip()
    keywords = keywords[keywords["keyword"] != ""].copy()

    return keywords

# =============================
# COLLECT
# =============================

def collect_google_rss(keywords):
    rows = []

    for _, row in keywords.iterrows():
        kw = str(row.get("keyword", "")).strip()
        lang = str(row.get("language", "EN")).strip().upper()
        category = str(row.get("category", "")).strip()
        importance = row.get("importance", "")
        importance_score = importance_to_score(importance)

        locale = get_google_locale(lang)
        query = quote(kw)

        rss_url = (
            "https://news.google.com/rss/search?"
            f"q={query}"
            f"&hl={locale['hl']}"
            f"&gl={locale['gl']}"
            f"&ceid={locale['ceid']}"
        )

        feed = feedparser.parse(rss_url)

        for entry in feed.entries:
            dt = parse_datetime(entry)
            if not is_recent(dt):
                continue

            title = clean_html(entry.get("title", ""))
            if not title:
                continue

            google_url = entry.get("link", "")
            summary = clean_html(entry.get("summary", ""))
            publisher = extract_publisher(entry, title)
            adjusted_score, score_reason = adjust_importance_score(
                kw, title, summary, importance_score
            )
            original_url_candidate = extract_original_url_candidate(entry, google_url)

            # 수집 직후에는 Google URL 저장, dedup 후 병렬 원문 URL 복구 단계에서 url/original_url_candidate 갱신
            rows.append({
                "date": dt,
                "title": title,
                "url": google_url,
                "google_url": google_url,
                "source": "Google News RSS",
                "summary": summary,
                "collected_at": datetime.now().replace(microsecond=0),
                "keyword": kw,
                "language": lang,
                "publisher": publisher,
                "category": category,
                "importance": importance,
                "importance_score": adjusted_score,
                "score_reason": score_reason,
                "url_decode_status": "PENDING_STEP2_RESOLVE" if ENABLE_ORIGINAL_URL_RESOLVE else "SKIPPED_STEP2_FAST",
                "original_url_candidate": original_url_candidate,
                "rss_url": rss_url,
            })

        time.sleep(SLEEP_SEC)

    return pd.DataFrame(rows)

# =============================
# DEDUP
# =============================

def dedup_fast(df):
    before = len(df)

    df["google_url_key"] = df["google_url"].astype(str).str.strip().str.lower()
    df["title_key"] = df["title"].apply(normalize_title)

    df = df.sort_values(["importance_score", "date"], ascending=[False, False])

    # 1차: Google URL 기준 정확 중복 제거
    df = df.drop_duplicates(subset=["google_url_key"], keep="first")

    # 2차: title_key 기준 중복 제거
    # 동일 기사가 keyword만 다르게 여러 번 잡히는 경우 제거
    before_title = len(df)
    df = df.drop_duplicates(subset=["title_key"], keep="first")

    print(f"📊 DEDUP GOOGLE_URL/TITLE: {before} -> {len(df)}")
    print(f"   - title dedup effect: {before_title} -> {len(df)}")

    df = df.drop(columns=["google_url_key", "title_key"], errors="ignore")
    return df

# =============================
# MAIN
# =============================

def main():
    total_start = time.perf_counter()

    t0 = time.perf_counter()
    keywords = load_keywords()
    print(f"🔎 active keywords: {len(keywords)}")
    print(f"⏱ keyword load elapsed: {elapsed_text(time.perf_counter() - t0)}")

    t0 = time.perf_counter()
    df = collect_google_rss(keywords)
    print(f"⏱ rss collect elapsed: {elapsed_text(time.perf_counter() - t0)}")

    if df.empty:
        print("❌ No data collected")
        saved_path, save_status = safe_to_excel(pd.DataFrame(columns=FINAL_COLS), RAW_FILE, index=False)
        print(f"💾 saved empty file: {saved_path} / status={save_status}")
        print(f"⏱ total elapsed: {elapsed_text(time.perf_counter() - total_start)}")
        return

    print(f"📊 TOTAL RAW: {len(df)}")

    before = len(df)
    df = df[df["date"].apply(is_recent)].copy()
    print(f"📊 24h FILTER: {before} -> {len(df)}")

    t0 = time.perf_counter()
    df = dedup_fast(df)
    print(f"⏱ dedup elapsed: {elapsed_text(time.perf_counter() - t0)}")

    t0 = time.perf_counter()
    df = resolve_original_urls_parallel(df)
    print(f"⏱ original url resolve stage elapsed: {elapsed_text(time.perf_counter() - t0)}")

    df = df.sort_values(["importance_score", "date"], ascending=[False, False])
    df = df[FINAL_COLS]

    t0 = time.perf_counter()
    df.to_excel(RAW_FILE, index=False)
    print(f"⏱ excel save elapsed: {elapsed_text(time.perf_counter() - t0)}")

    ok_status_count = df["url_decode_status"].astype(str).str.startswith("OK").sum() if "url_decode_status" in df.columns else 0
    ok_real_count = df["url"].apply(lambda x: bool(str(x).strip()) and not is_bad_original_url(x)).sum()
    bad_real_count = len(df) - int(ok_real_count)
    print("💾 saved:", RAW_FILE)
    print(f"✅ FINAL SAVE ROWS: {len(df)}")
    print(f"✅ ORIGINAL URL STATUS OK: {ok_status_count}/{len(df)} ({(ok_status_count/len(df)*100 if len(df) else 0):.1f}%)")
    print(f"✅ ORIGINAL URL REAL OK: {ok_real_count}/{len(df)} ({(ok_real_count/len(df)*100 if len(df) else 0):.1f}%)")
    if bad_real_count:
        print(f"⚠️ ORIGINAL URL BAD/GOOGLE/FALLBACK: {bad_real_count}")
    print(f"⏱ TOTAL STEP2-2 elapsed: {elapsed_text(time.perf_counter() - total_start)}")
    print("✅ STEP2-2 GOOGLE NEWS + ORIGINAL URL PARALLEL DONE v3.4")


if __name__ == "__main__":
    main()
