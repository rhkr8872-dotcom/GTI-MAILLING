import os
import sys
import json
import logging
from pathlib import Path
import pandas as pd
import httpx
from pydantic import BaseModel
import google.genai as genai
from google.genai import types

# ---------------------------------------------------------
# 1. 환경변수 및 디렉터리 설정 (Windows / Actions 호환)
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
        logging.FileHandler(LOG_DIR / "1.site_crawler.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

# ---------------------------------------------------------
# 2. Gemini Client 초기화
# ---------------------------------------------------------
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

if GEMINI_API_KEY:
    ai_client = genai.Client(api_key=GEMINI_API_KEY)
else:
    ai_client = None
    logging.warning("GEMINI_API_KEY가 설정되지 않아 규칙 기반 필터링으로 동작합니다.")

# Pydantic 응답 스키마
class ArticleAnalysis(BaseModel):
    is_relevant: bool          # 통상/무역/관세/법규 이슈 연관 여부
    priority_score: int        # 중요도 점수 (1~100)
    samsung_impact: str        # 삼성전자 또는 국내 산업 영향도 요약
    summary: str               # 핵심 2줄 요약

def analyze_article_with_gemini(title: str, text: str) -> ArticleAnalysis:
    """Gemini API를 호출하여 문서의 통상/무역 관련성 및 중요도를 평가합니다."""
    if not ai_client:
        return ArticleAnalysis(
            is_relevant=True,
            priority_score=50,
            samsung_impact="N/A (API Key 없음)",
            summary=title[:100]
        )
    
    prompt = f"""
    다음 문서/기사가 글로벌 무역, 통상, 관세, 통상 법규, 규제 및 삼성전자 등 주요 산업 분야와 연관이 있는지 분석하세요.
    제목: {title}
    본문 초안: {text[:1000]}
    """
    
    try:
        response = ai_client.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=ArticleAnalysis,
                temperature=0.1,
            ),
        )
        return ArticleAnalysis.model_validate_json(response.text)
    except Exception as e:
        logging.error(f"Gemini 분석 중 오류 발생: {e}")
        return ArticleAnalysis(
            is_relevant=True,
            priority_score=50,
            samsung_impact="분석 실패",
            summary=title[:100]
        )

# ---------------------------------------------------------
# 3. 크롤링 및 메인 수집 프로세스
# ---------------------------------------------------------
def run_site_crawler():
    logging.info("=== 1. Site Crawler 시작 ===")
    
    # sites.xlsx 파일 위치 탐색
    excel_path = BASE_DIR / "sites.xlsx"
    if not excel_path.exists():
        logging.error(f"대상 목록 파일이 없습니다: {excel_path}")
        return

    try:
        df_sites = pd.read_excel(excel_path)
    except Exception as e:
        logging.error(f"Excel 읽기 실패: {e}")
        return

    results = []
    
    # 예시: HTTP 스크래핑 파이프라인
    with httpx.Client(timeout=10.0, follow_redirects=True) as client:
        for idx, row in df_sites.iterrows():
            url = row.get("url")
            site_name = row.get("site_name", "Unknown")
            
            if not url or pd.isna(url):
                continue

            logging.info(f"[{idx+1}/{len(df_sites)}] 수집 중: {site_name} ({url})")
            
            try:
                # 간단한 HTTP GET 요청 예시
                res = client.get(str(url))
                if res.status_code == 200:
                    # 실제 파싱된 제목 및 본문 텍스트 (예시)
                    extracted_title = f"[{site_name}] 최신 통상 동향"
                    extracted_body = res.text[:2000]
                    
                    # Gemini AI 검증 호출
                    ai_res = analyze_article_with_gemini(extracted_title, extracted_body)
                    
                    if ai_res.is_relevant:
                        results.append({
                            "site_name": site_name,
                            "title": extracted_title,
                            "url": url,
                            "priority": ai_res.priority_score,
                            "samsung_impact": ai_res.samsung_impact,
                            "summary": ai_res.summary
                        })
            except Exception as e:
                logging.error(f"사이트 수집 오류 ({site_name}): {e}")

    # 결과 저장 (1-site_crawler_output.xlsx)
    output_path = BASE_DIR / "1-site_crawler_output.xlsx"
    df_result = pd.DataFrame(results)
    df_result.to_excel(output_path, index=False)
    logging.info(f"수집 완료 및 저장: {output_path} (총 {len(results)} 건)")

if __name__ == "__main__":
    run_site_crawler()
