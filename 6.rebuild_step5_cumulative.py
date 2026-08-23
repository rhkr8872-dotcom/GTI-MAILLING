from __future__ import annotations

from pathlib import Path
import re
import pandas as pd


BASE = Path(r"C:\Temp\12345\c_type_outputs")
OUTPUT = BASE / "gti_news_cumulative_rebuilt.xlsx"


def clean(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def event_key(row: pd.Series) -> str:
    cluster = clean(row.get("Cluster"))
    if cluster:
        return cluster.lower()
    title = re.sub(r"[^0-9a-z가-힣]+", " ", clean(row.get("Headline")).lower())
    return " ".join(title.split())


def main() -> None:
    files = sorted(BASE.glob("[[]GTI Radar[]]*.xlsx"))
    frames = []
    for path in files:
        try:
            book = pd.ExcelFile(path)
            sheet = "All" if "All" in book.sheet_names else book.sheet_names[0]
            frame = pd.read_excel(path, sheet_name=sheet)
            if "Headline" not in frame.columns:
                print(f"[SKIP] no Headline: {path.name}")
                continue
            frame["RebuildSource"] = path.name
            frames.append(frame)
            print(f"[LOAD] {path.name}: {len(frame)}")
        except Exception as exc:
            print(f"[SKIP] {path.name}: {type(exc).__name__}: {exc}")
    if not frames:
        raise RuntimeError(f"No GTI report XLSX found under {BASE}")
    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined["_event_key"] = combined.apply(event_key, axis=1)
    before = len(combined)
    combined = combined.drop_duplicates("_event_key", keep="last").drop(columns="_event_key")
    combined.to_excel(OUTPUT, index=False)
    print(f"[DONE] {before} -> {len(combined)} / {OUTPUT}")
    print("Review the rebuilt file, then rename it to gti_news_cumulative.xlsx if approved.")


if __name__ == "__main__":
    main()
