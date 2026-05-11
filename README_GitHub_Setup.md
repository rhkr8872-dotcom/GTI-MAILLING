# GTI GitHub Compatible Package

## 구조

```text
GTI/
├─ 1.site_crawler.py
├─ 2-1.naver_news_collector.py
├─ 2-2.google_news_collector.py
├─ 2-3.rss_news_collector.py
├─ 3.merge_news.py
├─ 5.GTI Mail Engine.py
├─ requirements.txt
├─ data/
│  ├─ site.xlsx
│  ├─ keyword.xlsx
│  └─ 00.xlsx
├─ output/
└─ .github/workflows/run-gti.yml
```

## STEP5 입력 우선순위

1. output/3.news_ai_summary.xlsx
2. output/news_ai_summary.xlsx
3. output/4.news_raw.xlsx
4. output/news_raw.xlsx

## GitHub Secrets

필수:
- GTI_SMTP_USER
- GTI_SMTP_PASS
- GTI_MAIL_TO

권장:
- GTI_SMTP_HOST = smtp.naver.com
- GTI_SMTP_PORT = 465
