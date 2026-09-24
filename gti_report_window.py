"""Shared KST calendar-day window for the GTI daily report."""
from __future__ import annotations

import os
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd


KST = ZoneInfo("Asia/Seoul")


def kst_now() -> datetime:
    override = os.getenv("GTI_NOW", "").strip()
    if override:
        value = datetime.fromisoformat(override)
        return value.replace(tzinfo=KST) if value.tzinfo is None else value.astimezone(KST)
    return datetime.now(KST)


def previous_kst_day_window(now: datetime | None = None) -> tuple[pd.Timestamp, pd.Timestamp]:
    current = now or kst_now()
    if current.tzinfo is None:
        current = current.replace(tzinfo=KST)
    else:
        current = current.astimezone(KST)
    end = datetime.combine(current.date(), datetime.min.time(), tzinfo=KST)
    start = end - timedelta(days=1)
    return pd.Timestamp(start.replace(tzinfo=None)), pd.Timestamp(end.replace(tzinfo=None))


def normalize_kst_publication_dates(values: pd.Series) -> pd.Series:
    def one(value):
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return pd.NaT
        try:
            parsed = pd.Timestamp(value)
        except Exception:
            return pd.NaT
        if pd.isna(parsed):
            return pd.NaT
        if parsed.tzinfo is not None:
            parsed = parsed.tz_convert(KST).tz_localize(None)
        return parsed

    return values.map(one)


def previous_kst_day_mask(values: pd.Series, now: datetime | None = None) -> tuple[pd.Series, pd.Series, pd.Timestamp, pd.Timestamp]:
    published = normalize_kst_publication_dates(values)
    start, end = previous_kst_day_window(now)
    mask = published.notna() & published.ge(start) & published.lt(end)
    return mask, published, start, end
