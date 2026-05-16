# GTI 자동 메일링 구성

## 실행 순서

1. `1.collect_gti_news.py`
   - `data/sites.xlsx`, `data/keyword.xlsx`를 읽습니다.
   - 네이버 뉴스 API, 구글뉴스 RSS, RSS/정부 사이트, NewsAPI, SerpAPI에서 최근 24시간 뉴스를 수집합니다.
   - `3.news_master_raw.xlsx`, `3.news_ai_summary.xlsx`, `3.news_ai_cumulative.xlsx`를 생성합니다.

2. `4.policy_ai_analyzer.py`
   - `3.news_ai_summary.xlsx`, `3.news_master_raw.xlsx`를 읽습니다.
   - 동일 URL, 유사 제목, 관세·통상 노이즈를 제거합니다.
   - 삼성전자 생산거점과 제품군 관점으로 Summary, AI Analysis, Action Plan을 생성합니다.
   - `4.news_ai_analysis.xlsx`, `4.news_cumulative.xlsx`를 저장합니다.

3. `5.GTI_Mail_Engine.py`
   - `4.news_ai_analysis.xlsx`를 읽습니다.
   - Top3와 정책 이벤트 표 HTML을 생성합니다.
   - `GTI_Radar_YYYY-MM-DD_Top30.xlsx`, `GTI_Radar_YYYY-MM-DD_Top30_Email.html`, `mail_cumulative.xlsx`를 저장합니다.
   - `GTI_SEND_EMAIL=Y`이면 SMTP로 메일을 발송합니다.

4. `7.run_gti_pipeline.py`
   - 수집, 분석, 메일 생성을 순서대로 실행합니다.

## 로컬 실행 예시

```powershell
$env:GTI_BASE_DIR="C:\Temp"
$env:GTI_OUTPUT_DIR="C:\Temp"
$env:GTI_RUN_DATE="2026-05-12"
$env:GTI_SEND_EMAIL="N"
python 7.run_gti_pipeline.py
```

메일 발송 시에는 `GTI_SEND_EMAIL=Y`, `GTI_SMTP_USER`, `GTI_SMTP_PASS`, `GTI_MAIL_TO`를 환경변수로 설정합니다.

## GitHub Actions 설정

`.github/workflows/gti-daily-mail.yml`은 매일 06:00 KST에 실행됩니다.

GitHub 저장소의 `Settings > Secrets and variables > Actions`에 아래 Secrets를 등록하세요.

- `GTI_SEND_EMAIL`: `Y`
- `GTI_SMTP_HOST`: 예: `smtp.naver.com`
- `GTI_SMTP_PORT`: 예: `465`
- `GTI_SMTP_USER`: 발송 계정
- `GTI_SMTP_PASS`: SMTP 앱 비밀번호
- `GTI_MAIL_TO`: 수신자 이메일, 쉼표로 구분
- `NAVER_CLIENT_ID`: 네이버 개발자 센터 Client ID
- `NAVER_CLIENT_SECRET`: 네이버 개발자 센터 Client Secret
- `SERPAPI_KEY`: SerpAPI Key
- `NEWS_API`: `https://newsapi.org/v2/everything`
- `NEWS_KEY`: NewsAPI Key
- `GEMINI_API_KEY`: Gemini 사용 시 등록
- `GEMINI_MODEL`: 예: `gemini-1.5-flash`

GitHub 실행 환경에서는 `data/` 폴더에 `sites.xlsx`, `keyword.xlsx`가 있어야 합니다. 수집 스크립트가 이 두 파일을 기준으로 매일 최신 뉴스를 다시 만들고, 분석·메일 단계로 넘깁니다.

## data 폴더

저장소에 아래 파일을 넣으세요.

- `data/sites.xlsx`: 관리 대상 정부 사이트 또는 RSS URL 목록
- `data/keyword.xlsx`: 네이버/구글/NewsAPI/SerpAPI 검색 제시어 목록

컬럼명은 엄격하지 않습니다. 스크립트는 엑셀 안의 URL과 키워드 문자열을 자동으로 읽습니다.
