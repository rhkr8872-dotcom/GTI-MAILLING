import pandas as pd

df = pd.read_excel("dedup_news.xlsx")

def classify_importance(text):

    text = str(text).lower()

    if "tariff" in text or "관세" in text:
        return "상"

    if "customs" in text:
        return "중"

    return "하"


df["importance"] = df["title"].apply(classify_importance)

df.to_excel("analyzed_news.xlsx", index=False)