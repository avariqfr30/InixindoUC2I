import argparse
import json

import pandas as pd

from core import InternalDataProvider
from internal_api import EndpointSpec, InternalApiClient


def load_json_object(raw_value):
    raw_value = (raw_value or "").strip()
    if not raw_value:
        return {}
    parsed = json.loads(raw_value)
    if not isinstance(parsed, dict):
        raise ValueError("Expected a JSON object.")
    return parsed


def build_target(target, method, params_json, headers_json, record_keys, auto_discover):
    if target.startswith(("http://", "https://")) or method or params_json or headers_json or record_keys:
        return EndpointSpec(
            name="ad_hoc",
            path=target,
            method=(method or "GET").strip().upper(),
            query_params=load_json_object(params_json),
            headers=load_json_object(headers_json),
            record_keys=tuple(record_keys or ["items", "data", "results", "records", "feedback"]),
            auto_discover=auto_discover,
        )
    return target


def main():
    parser = argparse.ArgumentParser(description="Inspect how the app interprets an internal API JSON payload.")
    parser.add_argument("target", nargs="?", help="Endpoint name or full URL, e.g. feedback or https://example.com/api/feedback")
    parser.add_argument("--fetch", action="store_true", help="Fetch the endpoint and print a short interpretation summary.")
    parser.add_argument("--method", help="HTTP method for ad-hoc URLs, e.g. GET or POST")
    parser.add_argument("--params-json", help="JSON object sent as query params for GET or request body for non-GET")
    parser.add_argument("--headers-json", help="Extra per-endpoint headers as JSON")
    parser.add_argument("--record-key", action="append", dest="record_keys", help="Preferred JSON keys used to find the record list. Repeatable.")
    parser.add_argument("--no-auto-discover", action="store_true", help="Disable automatic list discovery and only use the provided record keys.")
    args = parser.parse_args()

    client = InternalApiClient()
    if not args.target:
        print(json.dumps({"available_endpoints": client.available_endpoints()}, indent=2, ensure_ascii=False))
        return

    target = build_target(
        args.target,
        args.method,
        args.params_json,
        args.headers_json,
        args.record_keys,
        auto_discover=not args.no_auto_discover,
    )

    description = client.describe_endpoint(target)
    print(json.dumps(description, indent=2, ensure_ascii=False))

    if not args.fetch:
        return

    interpreted = client.interpret_payload(target)
    raw_df = pd.DataFrame(interpreted["records"])
    normalized_df = InternalDataProvider.normalize_dataframe(raw_df)
    print(
        json.dumps(
            {
                "record_path": interpreted["record_path"],
                "record_count": interpreted["record_count"],
                "sample_keys": interpreted["sample_keys"],
                "normalized_columns": list(normalized_df.columns),
                "sample": normalized_df.head(3).to_dict(orient="records"),
            },
            indent=2,
            ensure_ascii=False,
            default=str,
        )
    )


if __name__ == "__main__":
    main()
