import argparse
import json
import os

import pandas as pd

from config import INTERNAL_CONNECTOR_PATH
from core import CANONICAL_INTERNAL_COLUMNS, COLUMN_ALIASES, InternalDataProvider
from internal_api import EndpointSpec, InternalApiClient

DEFAULT_RECORD_KEYS = ["dataset_result", "items", "data", "results", "records", "feedback"]
REQUIRED_FIELDS = ("Tipe Stakeholder", "Layanan", "Tanggal Feedback", "Rating", "Komentar")


def load_json_object(raw_value):
    raw_value = (raw_value or "").strip()
    if not raw_value:
        return {}
    parsed = json.loads(raw_value)
    if not isinstance(parsed, dict):
        raise ValueError("Expected a JSON object.")
    return parsed


def load_json_file(path):
    with open(path, "r", encoding="utf-8") as handle:
        return json.load(handle)


def build_target(target, method, body_mode, params_json, headers_json, record_path, record_keys, auto_discover):
    target_ref = target or "feedback"
    if target_ref.startswith(("http://", "https://")) or method or params_json or headers_json or record_path or record_keys:
        return EndpointSpec(
            name="ad_hoc",
            path=target_ref,
            method=(method or "GET").strip().upper(),
            body_mode=body_mode,
            record_path=(record_path or "").strip(),
            query_params=load_json_object(params_json),
            headers=load_json_object(headers_json),
            record_keys=tuple(record_keys or DEFAULT_RECORD_KEYS),
            auto_discover=auto_discover,
        )
    return target_ref


def summarize_normalized_dataframe(normalized_df):
    field_summary = {}
    for field_name in REQUIRED_FIELDS:
        if field_name not in normalized_df.columns:
            field_summary[field_name] = {"status": "missing_column", "non_empty_rows": 0}
            continue
        non_empty_rows = int(normalized_df[field_name].fillna("").astype(str).str.strip().ne("").sum())
        field_summary[field_name] = {
            "status": "ok" if non_empty_rows > 0 else "empty",
            "non_empty_rows": non_empty_rows,
        }
    return {
        "normalized_columns": list(normalized_df.columns),
        "required_field_status": field_summary,
        "ready_for_knowledge_base": all(item["status"] == "ok" for item in field_summary.values()),
        "sample": normalized_df.head(3).to_dict(orient="records"),
    }


def suggest_field_map(raw_columns):
    suggestion = {}
    for target_column in CANONICAL_INTERNAL_COLUMNS:
        match = InternalDataProvider._find_matching_column(raw_columns, (target_column, *COLUMN_ALIASES.get(target_column, ())))
        if match and match != target_column:
            suggestion[match] = target_column
    return suggestion


def connector_template(target, interpreted, field_map, args):
    endpoint = target if isinstance(target, EndpointSpec) else EndpointSpec(name=str(target), path=str(target))
    return {
        "name": "apidog_feedback",
        "enabled": True,
        "endpoint_name": endpoint.name if endpoint.name != "ad_hoc" else "feedback",
        "url": endpoint.path if endpoint.path.startswith(("http://", "https://")) else "",
        "method": endpoint.method,
        "body_mode": endpoint.body_mode,
        "request_data": dict(endpoint.query_params),
        "headers": dict(endpoint.headers),
        "record_path": interpreted["record_path"],
        "record_keys": list(endpoint.record_keys),
        "auto_discover": True,
        "field_map": field_map,
        "required_fields": list(REQUIRED_FIELDS),
    }


def print_interpretation(client, target, payload=None, write_connector_path=None, args=None):
    interpreted = (
        client.interpret_payload_object(target, payload)
        if payload is not None
        else client.interpret_payload(target)
    )
    raw_df = pd.DataFrame(interpreted["records"])
    normalized_df = InternalDataProvider.normalize_dataframe(raw_df)
    field_map = suggest_field_map(list(raw_df.columns))
    output = {
        "endpoint": interpreted["endpoint"],
        "record_path": interpreted["record_path"],
        "record_count": interpreted["record_count"],
        "sample_keys": interpreted["sample_keys"],
        "field_map_suggestion": field_map,
        **summarize_normalized_dataframe(normalized_df),
    }
    if write_connector_path:
        template = connector_template(target, interpreted, field_map, args)
        target_dir = os.path.dirname(write_connector_path)
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)
        with open(write_connector_path, "w", encoding="utf-8") as handle:
            json.dump(template, handle, ensure_ascii=False, indent=2)
        output["connector_written"] = write_connector_path
    print(json.dumps(output, indent=2, ensure_ascii=False, default=str))


def main():
    parser = argparse.ArgumentParser(description="Inspect how the app interprets an internal API JSON payload.")
    parser.add_argument("target", nargs="?", help="Endpoint name or full URL, e.g. feedback or https://example.com/api/feedback")
    parser.add_argument("--file", help="Inspect a saved APIDog JSON response instead of calling the live endpoint.")
    parser.add_argument("--fetch", action="store_true", help="Fetch the endpoint and print a short interpretation summary.")
    parser.add_argument("--method", help="HTTP method for ad-hoc URLs, e.g. GET or POST")
    parser.add_argument("--body-mode", choices=("json", "form"), default="json", help="Request body mode for non-GET ad-hoc URLs")
    parser.add_argument("--params-json", help="JSON object sent as query params for GET or request body for non-GET")
    parser.add_argument("--headers-json", help="Extra per-endpoint headers as JSON")
    parser.add_argument("--record-path", help="Explicit JSON path to the record list, e.g. data.dataset_result")
    parser.add_argument("--record-key", action="append", dest="record_keys", help="Preferred JSON keys used to find the record list. Repeatable.")
    parser.add_argument("--no-auto-discover", action="store_true", help="Disable automatic list discovery and only use the provided record keys.")
    parser.add_argument("--write-connector", nargs="?", const=INTERNAL_CONNECTOR_PATH, help="Write a connector JSON template using the detected payload shape.")
    args = parser.parse_args()

    client = InternalApiClient()
    if not args.target and not args.file:
        print(json.dumps({"available_endpoints": client.available_endpoints()}, indent=2, ensure_ascii=False))
        return

    target = build_target(
        args.target,
        args.method,
        args.body_mode,
        args.params_json,
        args.headers_json,
        args.record_path,
        args.record_keys,
        auto_discover=not args.no_auto_discover,
    )

    description = client.describe_endpoint(target)
    print(json.dumps(description, indent=2, ensure_ascii=False))

    if args.file:
        print_interpretation(
            client,
            target,
            payload=load_json_file(args.file),
            write_connector_path=args.write_connector,
            args=args,
        )
        return

    if not args.fetch:
        return

    print_interpretation(
        client,
        target,
        write_connector_path=args.write_connector,
        args=args,
    )


if __name__ == "__main__":
    main()
