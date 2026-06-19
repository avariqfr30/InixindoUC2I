import logging
import os
import re

import pandas as pd

from config import CSV_PATH
from class_report_adapter import ClassReportAdapter
from data_contract import CANONICAL_INTERNAL_COLUMNS, COLUMN_ALIASES, DATE_COLUMN_ALIASES
from internal_api import InternalApiClient
from internal_connector import load_internal_connector

logger = logging.getLogger(__name__)

class InternalDataProvider:
    source_name = "internal"

    def __init__(self):
        self.last_ingestion_metadata = {"connector": "", "datasets": []}

    @staticmethod
    def _normalize_token(value):
        return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")

    @classmethod
    def _column_variants(cls, column_name):
        raw_name = str(column_name).strip()
        normalized = cls._normalize_token(raw_name)
        variants = {normalized}

        for part in re.split(r"[.\[\]/\\]+", raw_name):
            token = cls._normalize_token(part)
            if token:
                variants.add(token)

        normalized_parts = [part for part in normalized.split("_") if part]
        if normalized_parts:
            variants.add(normalized_parts[-1])
            if len(normalized_parts) >= 2:
                variants.add("_".join(normalized_parts[-2:]))

        return variants

    @classmethod
    def _find_matching_column(cls, columns, aliases):
        alias_tokens = [cls._normalize_token(alias) for alias in aliases if cls._normalize_token(alias)]
        best_match = None
        best_score = -1

        for column_name in columns:
            variants = cls._column_variants(column_name)
            score = 0
            for alias_token in alias_tokens:
                if alias_token in variants:
                    score += 100
                elif any(
                    variant.endswith(f"_{alias_token}")
                    or variant.startswith(f"{alias_token}_")
                    or (len(alias_token) >= 4 and alias_token in variant)
                    for variant in variants
                ):
                    score += 60
            if score > best_score:
                best_match = column_name
                best_score = score

        return best_match if best_score > 0 else None

    @classmethod
    def normalize_dataframe(cls, raw_df):
        if raw_df is None:
            return pd.DataFrame(columns=list(CANONICAL_INTERNAL_COLUMNS))

        dataframe = raw_df.copy()
        dataframe.columns = [str(column).strip() for column in dataframe.columns]

        rename_map = {}
        for canonical_name, aliases in COLUMN_ALIASES.items():
            if canonical_name in dataframe.columns:
                continue
            matched_column = cls._find_matching_column(dataframe.columns, aliases)
            if matched_column:
                rename_map[matched_column] = canonical_name

        dataframe = dataframe.rename(columns=rename_map)

        feedback_dates = pd.Series(dtype="datetime64[ns]")
        for alias in DATE_COLUMN_ALIASES:
            matched_column = cls._find_matching_column(dataframe.columns, (alias,))
            if matched_column:
                feedback_dates = pd.to_datetime(
                    dataframe[matched_column],
                    errors="coerce",
                )
                if feedback_dates.notna().any():
                    dataframe["Tanggal Feedback"] = feedback_dates.dt.strftime("%Y-%m-%d")
                    break

        if "Rentang Waktu" not in dataframe.columns:
            for alias in DATE_COLUMN_ALIASES:
                matched_column = cls._find_matching_column(dataframe.columns, (alias,))
                if not matched_column:
                    continue
                parsed_dates = pd.to_datetime(
                    dataframe[matched_column],
                    errors="coerce",
                )
                if parsed_dates.notna().any():
                    dataframe["Rentang Waktu"] = parsed_dates.dt.to_period("M").astype(str)
                    break

        for column_name in CANONICAL_INTERNAL_COLUMNS:
            if column_name not in dataframe.columns:
                dataframe[column_name] = pd.NA

        if dataframe["Record ID"].isna().all():
            dataframe["Record ID"] = [f"FB-{index + 1:05d}" for index in range(len(dataframe))]

        dataframe["Rating"] = pd.to_numeric(dataframe["Rating"], errors="coerce")
        for column_name in CANONICAL_INTERNAL_COLUMNS:
            dataframe[column_name] = dataframe[column_name].fillna("").astype(str).str.strip()

        return dataframe

    def load_feedback_data(self):
        raise NotImplementedError

class DemoCsvProvider(InternalDataProvider):
    source_name = "demo_csv"

    def __init__(self):
        super().__init__()

    def load_feedback_data(self):
        if not os.path.exists(CSV_PATH):
            raise FileNotFoundError(f"CSV file not found: {CSV_PATH}")
        raw_df = pd.read_csv(CSV_PATH)
        self.last_ingestion_metadata = {
            "connector": "demo_csv",
            "datasets": [
                {
                    "endpoint_name": "demo_csv",
                    "dataset": "db.csv",
                    "dataset_cache": "local_file",
                    "role": "feedback_source",
                    "row_count": int(len(raw_df)),
                }
            ],
        }
        return self.normalize_dataframe(raw_df)

class InternalApiProvider(InternalDataProvider):
    source_name = "company_api"

    def __init__(self):
        super().__init__()
        self.client = InternalApiClient()
        self.dataset_name = "feedback"
        self.connector = load_internal_connector()

    def _reload_connector(self):
        self.connector = load_internal_connector()

    def _apply_connector_auth_mode(self):
        auth_mode = str(getattr(self.connector, "auth_mode", "") or "").strip().lower()
        if auth_mode == "basic_env":
            self.client.auth_mode = "basic"
        elif auth_mode == "none":
            self.client.auth_mode = "none"
        elif auth_mode == "bearer_env":
            self.client.auth_mode = "api_key"
            self.client.auth_prefix = self.client.auth_prefix or "Bearer"

    @staticmethod
    def _endpoint_dataset_code(endpoint):
        request_data = getattr(endpoint, "request_data", {}) or {}
        return str(request_data.get("dataset") or request_data.get("dataset_code") or "").strip()

    @classmethod
    def _endpoint_spec_for_fetch(cls, endpoint):
        return endpoint.to_endpoint_spec()

    def _load_via_connector(self):
        if not self.connector or not self.connector.enabled:
            return None

        self._apply_connector_auth_mode()
        endpoint_payloads = []
        reference_lookup = {}
        dataset_metadata = []
        for endpoint in self.connector.active_endpoints():
            endpoint_spec = self._endpoint_spec_for_fetch(endpoint)
            interpreted = self.client.interpret_payload(endpoint_spec)
            raw_df = pd.DataFrame(interpreted["records"])
            dataset_code = self._endpoint_dataset_code(endpoint)
            dataset_event = {
                "endpoint_name": endpoint.endpoint_name,
                "dataset": dataset_code,
                "dataset_cache": str(
                    (endpoint_spec.query_params or {}).get("dataset_cache") or ""
                ),
                "raw_record_count": int(len(raw_df)),
            }
            if raw_df.empty:
                dataset_event["role"] = "empty"
                dataset_event["row_count"] = 0
                dataset_metadata.append(dataset_event)
                logger.warning(
                    "Internal connector endpoint '%s' returned no records.",
                    endpoint.endpoint_name,
                )
                continue

            if ClassReportAdapter.looks_like_reference_class_report(
                raw_df,
                endpoint_name=endpoint.endpoint_name,
                dataset_code=dataset_code,
            ):
                reference_lookup.update(ClassReportAdapter.question_lookup(raw_df))
                dataset_event["role"] = "reference_lookup"
                dataset_event["row_count"] = 0
                dataset_metadata.append(dataset_event)
                continue
            dataset_event["role"] = "feedback_source"
            dataset_metadata.append(dataset_event)
            endpoint_payloads.append((endpoint, raw_df))

        normalized_frames = []
        for endpoint, raw_df in endpoint_payloads:
            if ClassReportAdapter.looks_like_class_report(raw_df):
                mapped_df = ClassReportAdapter.normalize(
                    raw_df,
                    endpoint.endpoint_name,
                    reference_lookup=reference_lookup,
                )
            else:
                mapped_df = endpoint.apply_field_map(raw_df)
            normalized_df = self.normalize_dataframe(mapped_df)
            if "Sumber Feedback" in normalized_df.columns:
                empty_source = normalized_df["Sumber Feedback"].astype(str).str.strip() == ""
                normalized_df.loc[empty_source, "Sumber Feedback"] = endpoint.endpoint_name
            normalized_frames.append(normalized_df)
            for dataset_event in dataset_metadata:
                if dataset_event["endpoint_name"] == endpoint.endpoint_name:
                    dataset_event["row_count"] = int(len(normalized_df))
                    break

        if not normalized_frames:
            raise ValueError(
                f"Internal connector '{self.connector.name}' returned no records."
            )

        combined_df = pd.concat(normalized_frames, ignore_index=True)
        if "Record ID" in combined_df.columns:
            combined_df = combined_df.drop_duplicates(subset=["Record ID"], keep="last")
        self.last_ingestion_metadata = {
            "connector": self.connector.name,
            "datasets": dataset_metadata,
        }
        return combined_df.reset_index(drop=True)

    def load_dataset(self, dataset_name, extra_params=None):
        raw_df = pd.DataFrame(
            self.client.fetch_records(dataset_name, extra_params=extra_params)
        )
        if raw_df.empty:
            raise ValueError(
                f"Internal API returned no records for endpoint '{dataset_name}'."
            )
        normalized_df = self.normalize_dataframe(raw_df)
        self.last_ingestion_metadata = {
            "connector": "legacy_internal_api",
            "datasets": [
                {
                    "endpoint_name": dataset_name,
                    "dataset": dataset_name,
                    "dataset_cache": "",
                    "role": "feedback_source",
                    "raw_record_count": int(len(raw_df)),
                    "row_count": int(len(normalized_df)),
                }
            ],
        }
        return normalized_df

    def load_feedback_data(self):
        self._reload_connector()
        connector_df = self._load_via_connector()
        if connector_df is not None:
            return connector_df
        return self.load_dataset(self.dataset_name)
