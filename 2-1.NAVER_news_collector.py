import os
import sys
import logging
from pathlib import Path
import pandas as pd
import httpx
from pydantic import BaseModel
import google.genai as genai
from google.genai import types

# ---------------------------------------------------------
# 1. 환경변수 및 디렉터리 설정
# ---------------------------------------------------------
BASE_DIR_STR = os.getenv("GTI_BASE_DIR", r"C:\temp")
BASE_DIR = Path(BASE_DIR_STR)
BASE_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = BASE_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "2-1.NAVER_news_collector.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

# NAVER API Secrets
NAVER_CLIENT_ID = os.getenv("NAVER_CLIENT_ID")
NAVER_CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET")

# Gemini API Secrets
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

if GEMINI_API_KEY:
    ai_client = genai.Client(api_key=GEMINI_API_KEY)
else:
    ai_client = None

class NewsFilterResult(BaseModel):
    is_trade_news: bool      # 무역/관세/통상 관련 주요 뉴스 여부
    samsung_impact: str     # 삼성전자 관련 영향도 (High, Medium, Low 및 내용)
    priority: int           # 우선순위 점수 (1~100)
    ai_summary: str         # AI 작성 2줄 핵심 요약

def filter_news_with_gemini(title: str, description: str) -> NewsFilterResult:
    """네이버 뉴스 항목에 대해 Gemini AI 검증 수행"""
    if not ai_client:
        return NewsFilterResult(
            is_trade_news=True,
            samsung_impact="N/A",
            priority=50,
            ai_summary=description[:100]
        )
    
    prompt = f"""
    아래 네이버 뉴스 항목이 무역, 관세, 통상, 수출입 규제 및 삼성전자 등 전자/반도체/IT 산업 동향에 유의미한 정보인지 판별하세요.
    제목: {title}
    요약문: {description}
    """
    
    try:
        response = ai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=NewsFilterResult,
                temperature=0.1,
            ),
        )
        return NewsFilterResult.model_validate_json(response.text)
    except Exception as e:
        logging.error(f"Gemini 뉴스 분석 실패: {e}")
        return NewsFilterResult(
            is_trade_news=True,
            samsung_impact="Error",
            priority=50,
            ai_summary=description[:100]
        )

# ---------------------------------------------------------
# 2. 네이버 뉴스 수집 로직
# ---------------------------------------------------------
def fetch_naver_news(keyword: str, display_count: int = 20):
    if not NAVER_CLIENT_ID or not NAVER_CLIENT_SECRET:
        logging.error("NAVER_CLIENT_ID 또는 NAVER_CLIENT_SECRET이 설정되지 않았습니다.")
        return []

    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id": NAVER_CLIENT_ID,
        "X-Naver-Client-Secret": NAVER_CLIENT_SECRET
    }
    params = {
        "query": keyword,
        "display": display_count,
        "sort": "date"
    }

    try:
        with httpx.Client(timeout=10.0) as client:
            resp = client.get(url, headers=headers, params=params)
            if resp.status_code == 200:
                return resp.json().get("items", [])
            else:
                logging.error(f"네이버 API 호출 실패 Code: {resp.status_code}")
                return []
    except Exception as e:
        logging.error(f"네이버 API 요청 중 예외 발생: {e}")
        return []

def run_naver_collector():
    logging.info("=== 2-1. NAVER News Collector 시작 ===")
    
    # 키워드 목록 로드 (keyword.xlsx 또는 기본 키워드 사용)
    keywords = ["무역 관세", "삼성전자 통상", "수출 규제", "반도체 통상"]
    keyword_file = BASE_DIR / "keyword.xlsx"
    
    if keyword_file.exists():
        try:
            df_kw = pd.read_excel(keyword_file)
            if "keyword" in df_kw.columns:
                keywords = df_kw["keyword"].dropna().tolist()
        except Exception as e:
            logging.warning(f"키워드 파일 로드 실패, 기본 키워드 사용: {e}")

    collected_items = []
    
    for kw in keywords:
        logging.info(f"키워드 검색 시작: {kw}")
        raw_news = fetch_naver_news(kw, display_count=15)
        
        for item in raw_news:
            title = item.get("title", "").replace("<b>", "").replace("</b>", "").replace("&quot;", '"')
            description = item.get("description", "").replace("<b>", "").replace("</b>", "").replace("&quot;", '"')
            link = item.get("originallink") or item.get("link")
            pub_date = item.get("pubDate")
            
            # Gemini AI 검증
            ai_eval = filter_news_with_gemini(title, description)
            
            if ai_eval.is_trade_news:
                collected_items.append({
                    "keyword": kw,
                    "title": title,
                    "link": link,
                    "pub_date": pub_date,
                    "priority": ai_eval.priority,
                    "samsung_impact": ai_eval.samsung_impact,
                    "ai_summary": ai_eval.ai_summary
                })

    # 중복 뉴스 제거 (링크 기준)
    df_news = pd.DataFrame(collected_items)
    if not df_news.empty:
        df_news.drop_duplicates(subset=["link"], inplace=True)
        # 중요도 높은 순으로 정렬
        df_news.sort_values(by="priority", ascending=False, inplace=True)

    output_path = BASE_DIR / "2-1.NAVER_news_collector_output.xlsx"
    df_news.to_excel(output_path, index=False)
    logging.info(f"네이버 뉴스 수집 및 저장 완료: {output_path} (총 {len(df_news)} 건)")

if __name__ == "__main__":
    run_naver_collector()
