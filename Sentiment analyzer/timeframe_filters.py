import re

import pandas as pd


ROLLING_TIMEFRAME_OPTIONS = (
    ("1 Bulan Terakhir", 1),
    ("3 Bulan Terakhir", 3),
    ("6 Bulan Terakhir", 6),
    ("12 Bulan Terakhir / 1 Tahun", 12),
)
UNKNOWN_APIDOG_TIMEFRAME = "Semua Data APIDog (tanggal tidak tersedia)"
FULL_CACHED_TIMEFRAME = "Semua Data Tersedia"
CUSTOM_TIMEFRAME_PREFIX = "custom_range:"


def _normalize_label(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()


def rolling_month_count(timeframe):
    normalized = _normalize_label(timeframe)
    for label, month_count in ROLLING_TIMEFRAME_OPTIONS:
        label_normalized = _normalize_label(label)
        if normalized == label_normalized or normalized.startswith(label_normalized):
            return month_count
    return None


def custom_timeframe_label(start_date, end_date):
    return f"{CUSTOM_TIMEFRAME_PREFIX}{start_date}..{end_date}"


def parse_custom_timeframe(timeframe):
    raw_value = str(timeframe or "").strip()
    if not raw_value.startswith(CUSTOM_TIMEFRAME_PREFIX):
        return None
    value = raw_value[len(CUSTOM_TIMEFRAME_PREFIX):]
    if ".." not in value:
        return None
    start_raw, end_raw = value.split("..", maxsplit=1)
    start = pd.to_datetime(start_raw, format="%Y-%m-%d", errors="coerce")
    end = pd.to_datetime(end_raw, format="%Y-%m-%d", errors="coerce")
    if pd.isna(start) or pd.isna(end):
        return None
    if end < start:
        start, end = end, start
    return start.normalize(), end.normalize()


def readable_timeframe_label(timeframe):
    custom_range = parse_custom_timeframe(timeframe)
    if custom_range is None:
        return str(timeframe or "")
    start_date, end_date = custom_range
    return f"{start_date.strftime('%Y-%m-%d')} sampai {end_date.strftime('%Y-%m-%d')}"


def _feedback_dates(dataframe):
    if dataframe is None or dataframe.empty or "Tanggal Feedback" not in dataframe.columns:
        return pd.Series(dtype="datetime64[ns]")
    raw_dates = dataframe["Tanggal Feedback"].fillna("").astype(str).str.strip()
    parseable = raw_dates.str.extract(r"^(\d{4}-\d{2}-\d{2})", expand=False)
    return pd.to_datetime(parseable, format="%Y-%m-%d", errors="coerce")


def has_feedback_dates(dataframe):
    return _feedback_dates(dataframe).notna().any()


def filter_by_timeframe(dataframe, timeframe):
    if dataframe is None or dataframe.empty:
        return dataframe.copy() if dataframe is not None else pd.DataFrame()

    if str(timeframe or "").strip() == FULL_CACHED_TIMEFRAME:
        dates = _feedback_dates(dataframe)
        return dataframe[dates.notna()].copy() if dates.notna().any() else dataframe.copy()

    custom_range = parse_custom_timeframe(timeframe)
    if custom_range is not None:
        dates = _feedback_dates(dataframe)
        if not dates.notna().any():
            return dataframe.iloc[0:0].copy()
        start_date, end_date = custom_range
        return dataframe[(dates >= start_date) & (dates <= end_date)].copy()

    month_count = rolling_month_count(timeframe)
    if month_count is None:
        if "Rentang Waktu" not in dataframe.columns:
            return dataframe.copy()
        return dataframe[dataframe["Rentang Waktu"].astype(str) == str(timeframe)].copy()

    dates = _feedback_dates(dataframe)
    if not dates.notna().any():
        if "Rentang Waktu" not in dataframe.columns:
            return dataframe.iloc[0:0].copy()
        return dataframe[dataframe["Rentang Waktu"].astype(str) == str(timeframe)].copy()

    anchor = dates.max()
    current_month_start = pd.Timestamp(year=anchor.year, month=anchor.month, day=1)
    start_date = current_month_start - pd.DateOffset(months=max(month_count, 1))
    return dataframe[dates >= start_date].copy()


def build_timeframe_options(dataframe):
    exact_values = []
    seen = set()
    has_dates = has_feedback_dates(dataframe)
    if dataframe is not None and not dataframe.empty and "Rentang Waktu" in dataframe.columns:
        for value in dataframe["Rentang Waktu"].dropna().astype(str).str.strip().tolist():
            if not value or value in seen:
                continue
            if has_dates and value == UNKNOWN_APIDOG_TIMEFRAME:
                continue
            seen.add(value)
            exact_values.append(value)

    rolling = [FULL_CACHED_TIMEFRAME, *[label for label, _ in ROLLING_TIMEFRAME_OPTIONS]] if has_dates else []
    return rolling if has_dates else sorted(exact_values)


def build_available_date_options(dataframe):
    dates = _feedback_dates(dataframe)
    readable = sorted(dates.dropna().dt.strftime("%Y-%m-%d").unique().tolist())
    return {
        "min": readable[0] if readable else "",
        "max": readable[-1] if readable else "",
        "dates": readable,
    }
