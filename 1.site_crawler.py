# -*- coding: utf-8 -*-
# GTI STEP1 FINAL - sites.xlsx 운영형
# input : C:\temp\sites.xlsx
# output: C:\temp\1.site_news_raw.xlsx
# form  : date / title / url / source / collected_at / agency

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


import re
import time
import warnings
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.error import URLError
from urllib.parse import parse_qs, urlencode, urljoin, urlparse
from urllib.request import Request, urlopen

import pandas as pd
import requests
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

BASE_DIR = Path(r"C:\temp")
SITE_FILE = BASE_DIR / "sites.xlsx"
OUT_FILE = BASE_DIR / "1.site_news_raw.xlsx"
REJECT_FILE = BASE_DIR / "1.site_news_reject_debug.xlsx"

DAYS_BACK = 7
MAX_PER_SITE = 20

results = []
rejects = []

BAD_TITLE_CONTAINS = [
    "로그인", "회원가입", "사이트맵", "skip", "menu", "home",
    "privacy", "cookie", "contact", "about us", "accessibility",
    "facebook", "twitter", "youtube", "instagram", "linkedin",
    "검색", "전체메뉴", "본문", "바로가기", "이전", "다음",
    "처음", "마지막", "다운로드", "첨부파일", "자주묻는질문", "faq",
    "네이버 블로그", "블로그", "blog.naver.com",
    "관보보기", "일자별 기간별", "마이페이지", "관심 관보",
    "발행예고보기", "내일관보", "관보분석",
    "등록·채용 신고", "관세사 · 법인 징계현황",
    "개인정보처리방침", "이메일무단수집거부"
]

BAD_TITLE_EXACT = {
    "", "-", "0", "new", "more", "보기", "상세보기", "검색",
    "공지사항", "보도자료", "고시", "공고", "훈령", "예규",
    "뉴스", "news", "home", "menu", "목록"
}

TRADE_WORDS = [
    "관세", "통관", "수입", "수출", "무역", "통상", "고시", "공고",
    "훈령", "예규", "입법예고", "행정예고", "FTA", "원산지",
    "customs", "tariff", "trade", "import", "export", "notice",
    "regulation", "announcement", "directive", "policy"
]


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", unescape(str(value))).strip()


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_active_value(value) -> bool:
    if pd.isna(value):
        return False
    if isinstance(value, (int, float)):
        return int(value) == 1
    v = str(value).strip().lower()
    return v in ["1", "1.0", "y", "yes", "true", "t"]


def normalize_date(value):
    if value is None or str(value).strip() == "":
        return None
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return None
        if hasattr(dt, "to_pydatetime"):
            dt = dt.to_pydatetime()
        if getattr(dt, "tzinfo", None):
            dt = dt.replace(tzinfo=None)
        return dt
    except Exception:
        return None


def extract_date_from_text(text):
    text = clean_text(text)

    patterns = [
        r"(20\d{2}[-/.]\s*\d{1,2}[-/.]\s*\d{1,2}\s+\d{1,2}:\d{2}(:\d{2})?)",
        r"(20\d{2}[-/.]\s*\d{1,2}[-/.]\s*\d{1,2})",
        r"(20\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.?)",
        r"(20\d{2}년\s*\d{1,2}월\s*\d{1,2}일)",
        r"([A-Z][a-z]{2,9}\s+\d{1,2},\s*20\d{2})",
        r"(\d{1,2}\s+[A-Z][a-z]{2,9}\s+20\d{2})",
        r"(\d{1,2}\s+[A-Z][a-z]{2,9},?\s+20\d{2})",
    ]

    for p in patterns:
        m = re.search(p, text)
        if not m:
            continue
        s = m.group(1)
        s = s.replace("년", "-").replace("월", "-").replace("일", "")
        s = s.replace(".", "-").replace("/", "-")
        s = re.sub(r"-+$", "", s)
        dt = normalize_date(s)
        if dt is not None:
            return dt

    return None


def extract_date_from_tag(tag):
    if tag is None:
        return None

    time_tag = tag.find("time")
    if time_tag:
        for attr in ["datetime", "content", "date"]:
            if time_tag.get(attr):
                dt = normalize_date(time_tag.get(attr))
                if dt:
                    return dt
        dt = extract_date_from_text(time_tag.get_text(" ", strip=True))
        if dt:
            return dt

    for key in [
        "article:published_time", "article:modified_time",
        "date", "dc.date", "dc:date", "pubdate",
        "publishdate", "published_time", "lastmod"
    ]:
        meta = (
            tag.find("meta", attrs={"property": key})
            or tag.find("meta", attrs={"name": key})
            or tag.find("meta", attrs={"itemprop": key})
        )
        if meta and meta.get("content"):
            dt = normalize_date(meta.get("content"))
            if dt:
                return dt

    return extract_date_from_text(tag.get_text(" ", strip=True))


def is_recent(dt):
    if dt is None:
        return False
    cutoff = datetime.now() - timedelta(days=DAYS_BACK)
    if isinstance(dt, date) and not isinstance(dt, datetime):
        dt = datetime.combine(dt, datetime.min.time())
    return dt >= cutoff


def is_valid_title(title):
    title = clean_text(title)
    low = title.lower()

    if not title:
        return False
    if low in BAD_TITLE_EXACT:
        return False
    if len(title) < 8:
        return False
    if re.fullmatch(r"\d+", title):
        return False
    if any(x.lower() in low for x in BAD_TITLE_CONTAINS):
        return False

    return True


def add_reject(reason, date_value, title, url, source, agency):
    rejects.append({
        "reason": reason,
        "date": str(date_value),
        "title": clean_text(title),
        "url": str(url),
        "source": str(source),
        "agency": str(agency),
        "checked_at": now_str()
    })


def add_result(date_value, title, url, source, agency):
    title = clean_text(title)
    url = str(url or "").strip()
    source = str(source or "").strip()
    agency = str(agency or "").strip()

    if not is_valid_title(title):
        add_reject("invalid_title", date_value, title, url, source, agency)
        return False

    if not url.startswith("http"):
        add_reject("invalid_url", date_value, title, url, source, agency)
        return False

    dt = normalize_date(date_value)
    if dt is None:
        dt = extract_date_from_text(str(date_value))

    if dt is None:
        add_reject("no_date", date_value, title, url, source, agency)
        return False

    if not is_recent(dt):
        add_reject("old_date", date_value, title, url, source, agency)
        return False

    results.append({
        "date": dt.strftime("%Y-%m-%d"),
        "title": title,
        "url": url,
        "source": source,
        "collected_at": now_str(),
        "agency": agency
    })
    return True


def fetch_html(url):
    try:
        r = requests.get(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
                "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            },
            timeout=20,
        )
        if r.status_code >= 400:
            return None
        return r
    except Exception:
        return None


def get_query_params(url):
    qs = parse_qs(urlparse(url).query)
    return {k: v[0] for k, v in qs.items()}


def find_best_anchor(container, base_url, href_keyword=None):
    anchors = container.find_all("a", href=True)
    best = None

    for a in anchors:
        title = clean_text(a.get_text(" ", strip=True))
        href = a.get("href", "")

        if href_keyword and href_keyword.lower() not in href.lower():
            continue
        if not is_valid_title(title):
            continue
        if str(href).lower().startswith(("javascript:", "#", "mailto:", "tel:")):
            continue

        link = urljoin(base_url, href)

        if "blog.naver.com" in link.lower():
            continue

        score = len(title)
        if re.search(r"[가-힣]", title):
            score += 10
        if any(k.lower() in title.lower() for k in TRADE_WORDS):
            score += 30

        item = (score, title, link)
        if best is None or item[0] > best[0]:
            best = item

    if best:
        return best[1], best[2]

    return None, None


KOREA_LIST_URL = "https://www.customs.go.kr/kcs/na/ntt/selectNttList.do"
KOREA_DETAIL_URL = "https://www.customs.go.kr/kcs/na/ntt/selectNttInfo.do"


@dataclass
class KoreaPost:
    number: str
    category: str
    title: str
    author: str
    posted_at: date
    views: str
    detail_url: str


def parse_korea_date(value: str):
    try:
        return datetime.strptime(value.strip(), "%Y.%m.%d").date()
    except Exception:
        return None


class CustomsBoardParser(HTMLParser):
    def __init__(self, base_params):
        super().__init__(convert_charrefs=True)
        self.base_params = base_params
        self.posts = []
        self._in_row = False
        self._in_cell = False
        self._cell_text = []
        self._cells = []
        self._link_attrs = {}
        self._anchor_title = ""

    def handle_starttag(self, tag, attrs):
        attr = {k: v or "" for k, v in attrs}

        if tag == "tr":
            self._in_row = True
            self._cells = []
            self._link_attrs = {}
            self._anchor_title = ""

        if self._in_row and tag == "td":
            self._in_cell = True
            self._cell_text = []

        if self._in_row and tag == "a" and "nttInfoBtn" in attr.get("class", ""):
            self._link_attrs = attr
            self._anchor_title = attr.get("title", "")

    def handle_data(self, data):
        if self._in_cell:
            self._cell_text.append(data)

    def handle_endtag(self, tag):
        if self._in_cell and tag == "td":
            self._cells.append(clean_text(" ".join(self._cell_text)))
            self._in_cell = False

        if self._in_row and tag == "tr":
            self._append_post_if_valid()
            self._in_row = False

    def _append_post_if_valid(self):
        if len(self._cells) < 6 or not self._link_attrs:
            return

        posted_at = parse_korea_date(self._cells[4])
        if posted_at is None:
            return

        title = clean_text(self._anchor_title or self._cells[2])
        ntt_sn = self._link_attrs.get("data-id", "")
        ntt_sn_url = self._link_attrs.get("data-url", "")

        detail_params = {
            **self.base_params,
            "nttSn": ntt_sn,
            "nttSnUrl": ntt_sn_url,
        }

        self.posts.append(
            KoreaPost(
                number=self._cells[0],
                category=self._cells[1],
                title=title,
                author=self._cells[3],
                posted_at=posted_at,
                views=self._cells[5],
                detail_url=f"{KOREA_DETAIL_URL}?{urlencode(detail_params)}",
            )
        )


def fetch_korea_html(source_url):
    params = get_query_params(source_url)

    if "mi" not in params or "bbsId" not in params:
        params = {"mi": "2891", "bbsId": "1362"}

    full_url = f"{KOREA_LIST_URL}?{urlencode(params)}"

    req = Request(
        full_url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
            "Connection": "close",
        },
    )

    last_error = None

    for attempt in range(1, 4):
        try:
            with urlopen(req, timeout=20) as response:
                charset = response.headers.get_content_charset() or "utf-8"
                html = response.read().decode(charset, errors="replace")
                return html, params, full_url
        except (ConnectionResetError, TimeoutError, URLError) as e:
            last_error = e
            time.sleep(attempt * 2)

    raise RuntimeError(f"Korea Customs 접속 실패: {last_error}")


def crawl_korea(source_url, agency):
    before = len(results)

    try:
        html, params, full_url = fetch_korea_html(source_url)
        parser = CustomsBoardParser(params)
        parser.feed(html)

        for post in parser.posts:
            add_result(
                date_value=post.posted_at,
                title=f"[{post.category}] {post.title}",
                url=post.detail_url,
                source=full_url,
                agency=agency
            )

    except Exception as e:
        print("   ERROR:", e)

    return len(results) - before


def crawl_rss(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "xml")
    items = soup.find_all(["item", "entry"])

    for item in items[:MAX_PER_SITE]:
        title_tag = item.find("title")
        link_tag = item.find("link")
        date_tag = (
            item.find("pubDate")
            or item.find("published")
            or item.find("updated")
            or item.find("dc:date")
        )

        title = title_tag.get_text(strip=True) if title_tag else ""
        link = ""

        if link_tag:
            link = link_tag.get("href") or link_tag.get_text(strip=True)

        dt = date_tag.get_text(strip=True) if date_tag else ""
        add_result(dt, title, urljoin(source_url, link), source_url, agency)

    return len(results) - before


def crawl_table(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    for row in soup.find_all("tr"):
        row_text = clean_text(row.get_text(" ", strip=True))
        if len(row_text) < 15:
            continue

        post_date = extract_date_from_tag(row)
        if post_date is None:
            continue

        title, link = find_best_anchor(row, source_url)
        if not title or not link:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    if count < 5:
        for tag in soup.find_all(["li", "dt", "dd", "article", "section", "div"]):
            text = clean_text(tag.get_text(" ", strip=True))
            if len(text) < 20:
                continue

            post_date = extract_date_from_tag(tag)
            if post_date is None:
                continue

            title, link = find_best_anchor(tag, source_url)
            if not title or not link:
                continue

            if add_result(post_date, title, link, source_url, agency):
                count += 1

            if count >= MAX_PER_SITE:
                break

    return len(results) - before


def crawl_card(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    for c in soup.find_all(["article", "li", "div", "section", "tr", "dt", "dd"]):
        text = clean_text(c.get_text(" ", strip=True))
        if len(text) < 20:
            continue

        post_date = extract_date_from_tag(c)
        if post_date is None:
            continue

        title, link = find_best_anchor(c, source_url)
        if not title or not link:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_ustr(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    for c in soup.find_all(["article", "li", "div", "tr"]):
        post_date = extract_date_from_tag(c)
        if post_date is None:
            continue

        title, link = find_best_anchor(c, source_url, "/press-releases/")
        if not title:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_cbp(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    for c in soup.find_all(["article", "li", "div", "tr"]):
        post_date = extract_date_from_tag(c)
        if post_date is None:
            continue

        title, link = find_best_anchor(c, source_url, "/newsroom/")
        if not title:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_eu(source_url, agency):
    if "rss" in source_url.lower() or source_url.lower().endswith(".xml"):
        return crawl_rss(source_url, agency)

    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    for c in soup.find_all(["article", "li", "div", "section"]):
        post_date = extract_date_from_tag(c)
        if post_date is None:
            continue

        title, link = find_best_anchor(c, source_url)
        if not title:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_wto(source_url, agency):
    if "rss" in source_url.lower() or source_url.lower().endswith(".xml"):
        return crawl_rss(source_url, agency)
    return crawl_table(source_url, agency)


def crawl_gwanbo(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    text = clean_text(soup.get_text("\n", strip=True))
    page_date = extract_date_from_text(text) or datetime.now()

    patterns = [
        r"(법률제\d+호\s*\([^)]+\))",
        r"(대통령령제\d+호\s*\([^)]+\))",
        r"(총리령제\d+호\s*\([^)]+\))",
        r"(부령제\d+호\s*\([^)]+\))",
        r"([가-힣]+부령제\d+호\s*\([^)]+\))",
        r"(고시제[\d\-]+호\s*\([^)]+\))",
        r"(공고제[\d\-]+호\s*\([^)]+\))",
        r"(훈령제\d+호\s*\([^)]+\))",
        r"(예규제\d+호\s*\([^)]+\))",
    ]

    seen = set()

    for p in patterns:
        for m in re.finditer(p, text):
            title = clean_text(m.group(1))
            if title in seen:
                continue
            seen.add(title)

            add_result(page_date, title, source_url, source_url, agency)

            if len(seen) >= MAX_PER_SITE:
                break

    return len(results) - before


def crawl_motie(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    category_words = [
        "훈령·예규·지침", "입법 예고", "행정 예고", "고시·공고",
        "보도자료", "예산·법령", "산업통상부 네이버 블로그"
    ]

    for row in soup.find_all(["tr", "li", "div", "article"]):
        post_date = extract_date_from_tag(row)
        if post_date is None:
            continue

        for a in row.find_all("a", href=True):
            title = clean_text(a.get_text(" ", strip=True))
            href = a.get("href", "")
            link = urljoin(source_url, href)

            if not title or len(title) < 10:
                continue
            if any(w in title for w in category_words):
                continue
            if "blog.naver.com" in link.lower():
                continue
            if "/kor/article/" not in link:
                continue

            if add_result(post_date, title, link, source_url, agency):
                count += 1
                break

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_law(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    law_keywords = [
        "법률", "시행령", "시행규칙", "고시", "훈령", "예규",
        "행정규칙", "공포", "개정", "폐지", "제정"
    ]

    for row in soup.find_all(["tr", "li", "div"]):
        row_text = clean_text(row.get_text(" ", strip=True))
        if len(row_text) < 15:
            continue

        post_date = extract_date_from_tag(row)
        if post_date is None:
            continue

        if not any(k in row_text for k in law_keywords):
            continue

        title, link = find_best_anchor(row, source_url)
        if not title or not link:
            continue

        if not any(k in title for k in law_keywords):
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_krcaa(source_url, agency):
    before = len(results)

    res = fetch_html(source_url)
    if not res:
        return 0

    soup = BeautifulSoup(res.text, "html.parser")
    count = 0

    menu_words = [
        "관세사 · 법인 징계현황",
        "등록·채용 신고",
        "로그인",
        "회원가입",
        "사이트맵",
        "개인정보처리방침",
        "오시는길",
        "조직도"
    ]

    for row in soup.find_all(["tr", "li", "div"]):
        post_date = extract_date_from_tag(row)
        if post_date is None:
            continue

        title, link = find_best_anchor(row, source_url)
        if not title or not link:
            continue

        if any(w in title for w in menu_words):
            continue

        if "Notify" not in link and "notify" not in link:
            continue

        if add_result(post_date, title, link, source_url, agency):
            count += 1

        if count >= MAX_PER_SITE:
            break

    return len(results) - before


def crawl_generic(source_url, agency):
    return crawl_card(source_url, agency)


def infer_parser(site_type, parser):
    p = str(parser or "").strip().lower()
    t = str(site_type or "").strip().lower()

    if p and p != "nan":
        return p

    if t in ["korea", "korea_press", "korea_fta", "korea_notice"]:
        return "customs_board"
    if t == "ustr":
        return "ustr_parser"
    if t == "cbp":
        return "cbp_parser"
    if t in ["eu", "eu_taxud"]:
        return "eu_parser"
    if t == "wto":
        return "wto_parser"
    if t == "rss":
        return "rss_parser"
    if t == "gwanbo":
        return "gwanbo_parser"
    if t == "motie":
        return "motie_parser"
    if t == "law":
        return "law_parser"
    if t == "krcaa":
        return "krcaa_parser"
    if t == "table":
        return "table_date"

    return "generic_html"


def get_status(count):
    if count == 0:
        return "FAIL"
    if count <= 2:
        return "CHECK"
    return "OK"


def main():
    print("🚀 GTI STEP1 SITES MODE START")

    BASE_DIR.mkdir(parents=True, exist_ok=True)

    if not SITE_FILE.exists():
        print(f"❌ sites.xlsx 없음: {SITE_FILE}")
        return

    sites = pd.read_excel(SITE_FILE)

    sites = sites.loc[:, ~sites.columns.astype(str).str.contains("^Unnamed")]
    sites = sites.loc[:, ~sites.columns.astype(str).str.endswith(".1")]

    for col in ["parser", "collected_count", "total_collected", "last_checked", "status"]:
        if col not in sites.columns:
            sites[col] = ""

    total_before = len(results)
    active_count = 0
    skipped_count = 0

    for idx, row in sites.iterrows():
        active = row.get("active", 1)

        if not is_active_value(active):
            skipped_count += 1
            print(f"[SKIP] inactive: {row.get('site_name', '')}")
            continue

        active_count += 1

        site_name = clean_text(row.get("site_name", ""))
        source_url = str(row.get("url", "")).strip()
        site_type = str(row.get("type", "generic")).strip().lower()
        parser = infer_parser(site_type, row.get("parser", ""))
        agency = site_name

        sites.at[idx, "parser"] = parser

        if not source_url.startswith("http"):
            count = 0
            sites.at[idx, "collected_count"] = count
            sites.at[idx, "last_checked"] = now_str()
            sites.at[idx, "status"] = get_status(count)
            continue

        print(f"[CRAWL] {site_name} / type={site_type} / parser={parser}")

        try:
            if parser == "customs_board":
                count = crawl_korea(source_url, agency)
            elif parser == "ustr_parser":
                count = crawl_ustr(source_url, agency)
            elif parser == "cbp_parser":
                count = crawl_cbp(source_url, agency)
            elif parser == "eu_parser":
                count = crawl_eu(source_url, agency)
            elif parser == "wto_parser":
                count = crawl_wto(source_url, agency)
            elif parser == "rss_parser":
                count = crawl_rss(source_url, agency)
            elif parser == "card_parser":
                count = crawl_card(source_url, agency)
            elif parser == "gwanbo_parser":
                count = crawl_gwanbo(source_url, agency)
            elif parser == "motie_parser":
                count = crawl_motie(source_url, agency)
            elif parser == "law_parser":
                count = crawl_law(source_url, agency)
            elif parser == "krcaa_parser":
                count = crawl_krcaa(source_url, agency)
            elif parser == "table_date":
                count = crawl_table(source_url, agency)
            else:
                count = crawl_generic(source_url, agency)

        except Exception as e:
            print("   ERROR:", e)
            count = 0

        print(f" → {count}건")

        prev_total = pd.to_numeric(sites.at[idx, "total_collected"], errors="coerce")
        if pd.isna(prev_total):
            prev_total = 0

        sites.at[idx, "collected_count"] = count
        sites.at[idx, "total_collected"] = int(prev_total) + int(count)
        sites.at[idx, "last_checked"] = now_str()
        sites.at[idx, "status"] = get_status(count)

    df = pd.DataFrame(results, columns=[
        "date", "title", "url", "source", "collected_at", "agency"
    ])

    raw_count = len(df)

    if df.empty:
        print("❌ No data")
        print(f"📌 active 대상: {active_count}")
        print(f"📌 inactive skip: {skipped_count}")
        sites.to_excel(SITE_FILE, index=False)
        if rejects:
            pd.DataFrame(rejects).to_excel(REJECT_FILE, index=False)
            print(f"🧾 제외 로그: {REJECT_FILE}")
        return

    df = df.drop_duplicates(subset=["url", "title"])
    final_count = len(df)
    dup_removed = raw_count - final_count

    df = df[["date", "title", "url", "source", "collected_at", "agency"]]

    try:
        df.to_excel(OUT_FILE, index=False)
    except PermissionError:
        alt = BASE_DIR / "1.site_news_raw_new.xlsx"
        df.to_excel(alt, index=False)
        print(f"⚠ 결과 파일 열림 → 대체 저장: {alt}")

    if rejects:
        try:
            pd.DataFrame(rejects).to_excel(REJECT_FILE, index=False)
        except PermissionError:
            pass

    try:
        sites.to_excel(SITE_FILE, index=False)
    except PermissionError:
        alt_sites = BASE_DIR / "sites_updated.xlsx"
        sites.to_excel(alt_sites, index=False)
        print(f"⚠ sites.xlsx 열림 → 대체 저장: {alt_sites}")

    print(f"📌 active 대상: {active_count}")
    print(f"📌 inactive skip: {skipped_count}")
    print(f"📊 RAW 수집: {raw_count}")
    print(f"🧹 중복 제거: {dup_removed}")
    print(f"✅ 최종 저장: {final_count}")
    print(f"📁 파일: {OUT_FILE}")
    print(f"🧾 제외 로그: {REJECT_FILE}")
    print(f"📊 RAW 증가: {len(results) - total_before}")


if __name__ == "__main__":
    main()
