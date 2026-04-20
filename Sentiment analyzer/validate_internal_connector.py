import json
import os
import sys

from config import (
    APP_MODE,
    APP_PROFILE,
    CSV_PATH,
    INTERNAL_API_BASE_URL,
    INTERNAL_CONNECTOR_PATH,
    INTERNAL_API_SOURCE_URL,
)
from core import DemoCsvProvider, InternalApiProvider

REQUIRED_CANONICAL_FIELDS = (
    "Tipe Stakeholder",
    "Layanan",
    "Tanggal Feedback",
    "Rating",
    "Komentar",
)


def summarize_dataframe(dataframe):
    summary = {
        "status": "ready",
        "row_count": len(dataframe),
        "columns": list(dataframe.columns),
        "required_field_status": {},
    }
    healthy_required_fields = 0
    for field_name in REQUIRED_CANONICAL_FIELDS:
        if field_name not in dataframe.columns:
            summary["required_field_status"][field_name] = "missing_column"
            summary["status"] = "failed"
            continue
        non_empty_count = int(
            dataframe[field_name]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .sum()
        )
        summary["required_field_status"][field_name] = {
            "non_empty_rows": non_empty_count,
            "status": "ok" if non_empty_count > 0 else "empty",
        }
        if non_empty_count > 0:
            healthy_required_fields += 1
        else:
            summary["status"] = "failed"
    summary["required_fields_ok"] = healthy_required_fields
    return summary


def validate_demo_mode():
    provider = DemoCsvProvider()
    dataframe = provider.load_feedback_data()
    return {
        "profile": APP_PROFILE,
        "mode": APP_MODE,
        "source": "demo_csv",
        "csv_path": CSV_PATH,
        "csv_exists": os.path.exists(CSV_PATH),
        **summarize_dataframe(dataframe),
    }


def validate_hybrid_mode():
    provider = InternalApiProvider()
    connector = provider.connector
    connector_info = connector.describe() if connector else None
    dataframe = provider.load_feedback_data()
    summary = {
        "profile": APP_PROFILE,
        "mode": APP_MODE,
        "source": "internal_api",
        "connector_path": INTERNAL_CONNECTOR_PATH,
        "connector_exists": os.path.exists(INTERNAL_CONNECTOR_PATH),
        "source_url": INTERNAL_API_SOURCE_URL,
        "base_url": INTERNAL_API_BASE_URL,
        "connector": connector_info,
        **summarize_dataframe(dataframe),
    }
    return summary


def main():
    try:
        if APP_MODE == "demo":
            summary = validate_demo_mode()
        else:
            summary = validate_hybrid_mode()
        print(json.dumps(summary, indent=2, ensure_ascii=False, default=str))
    except Exception as exc:
        print(
            json.dumps(
                {
                    "profile": APP_PROFILE,
                    "mode": APP_MODE,
                    "status": "failed",
                    "error": str(exc),
                },
                indent=2,
                ensure_ascii=False,
            )
        )
        sys.exit(1)


if __name__ == "__main__":
    main()
