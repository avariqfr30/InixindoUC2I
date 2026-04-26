from dataclasses import dataclass, field
import json
import logging
import re
from typing import Any

import requests

from config import (
    INTERNAL_API_AUTH_MODE,
    INTERNAL_API_AUTH_HEADER,
    INTERNAL_API_AUTH_PREFIX,
    INTERNAL_API_BASE_URL,
    INTERNAL_API_DEFAULT_HEADERS,
    INTERNAL_API_ENDPOINTS,
    INTERNAL_API_KEY,
    INTERNAL_API_PASSWORD,
    INTERNAL_API_TIMEOUT_SECONDS,
    INTERNAL_API_USERNAME,
)

logger = logging.getLogger(__name__)

DISCOVERY_HINTS = (
    "feedback",
    "comment",
    "review",
    "rating",
    "score",
    "csat",
    "sentiment",
    "service",
    "stakeholder",
    "segment",
    "channel",
    "source",
    "date",
    "location",
    "instructor",
)


@dataclass(frozen=True)
class EndpointSpec:
    name: str
    path: str
    method: str = "GET"
    body_mode: str = "json"
    record_path: str = ""
    record_keys: tuple[str, ...] = ("items", "data", "results", "records")
    query_params: dict[str, Any] = field(default_factory=dict)
    headers: dict[str, str] = field(default_factory=dict)
    auto_discover: bool = True

    @classmethod
    def from_mapping(cls, name, mapping):
        return cls(
            name=name,
            path=str(mapping.get("path", f"/{name}")).strip() or f"/{name}",
            method=str(mapping.get("method", "GET")).strip().upper() or "GET",
            body_mode=str(mapping.get("body_mode", "json")).strip().lower() or "json",
            record_path=str(mapping.get("record_path", "")).strip(),
            record_keys=tuple(mapping.get("record_keys") or ("items", "data", "results", "records", name)),
            query_params=dict(mapping.get("query_params") or {}),
            headers=dict(mapping.get("headers") or {}),
            auto_discover=bool(mapping.get("auto_discover", True)),
        )


class InternalApiClient:
    def __init__(
        self,
        base_url=INTERNAL_API_BASE_URL,
        api_key=INTERNAL_API_KEY,
        timeout_seconds=INTERNAL_API_TIMEOUT_SECONDS,
        auth_mode=INTERNAL_API_AUTH_MODE,
        auth_header=INTERNAL_API_AUTH_HEADER,
        auth_prefix=INTERNAL_API_AUTH_PREFIX,
        username=INTERNAL_API_USERNAME,
        password=INTERNAL_API_PASSWORD,
        default_headers=None,
        endpoints=None,
    ):
        self.base_url = (base_url or "").rstrip("/")
        self.api_key = api_key or ""
        self.timeout_seconds = int(timeout_seconds)
        self.auth_mode = (auth_mode or "api_key").strip().lower()
        self.auth_header = auth_header or "Authorization"
        self.auth_prefix = auth_prefix or ""
        self.username = username or ""
        self.password = password or ""
        self.default_headers = dict(default_headers or INTERNAL_API_DEFAULT_HEADERS)
        endpoint_mappings = endpoints or INTERNAL_API_ENDPOINTS
        self.endpoints = {
            endpoint_name: EndpointSpec.from_mapping(endpoint_name, spec)
            for endpoint_name, spec in endpoint_mappings.items()
            if isinstance(spec, dict)
        }

    @staticmethod
    def _is_absolute_url(value):
        return str(value or "").startswith(("http://", "https://"))

    @staticmethod
    def _normalize_token(value):
        return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")

    @staticmethod
    def _redact_mapping(mapping):
        redacted = {}
        for key, value in dict(mapping or {}).items():
            token = InternalApiClient._normalize_token(key)
            if any(secret_word in token for secret_word in ("authorization", "token", "key", "secret", "password")):
                redacted[key] = "***redacted***"
            else:
                redacted[key] = value
        return redacted

    def available_endpoints(self):
        return sorted(self.endpoints.keys())

    def get_endpoint(self, endpoint_ref):
        if isinstance(endpoint_ref, EndpointSpec):
            return endpoint_ref

        if self._is_absolute_url(endpoint_ref):
            return EndpointSpec(name="ad_hoc", path=str(endpoint_ref).strip())

        if endpoint_ref not in self.endpoints:
            known = ", ".join(self.available_endpoints()) or "none"
            raise KeyError(f"Unknown internal API endpoint '{endpoint_ref}'. Known endpoints: {known}")
        return self.endpoints[endpoint_ref]

    def _resolve_url(self, endpoint):
        if self._is_absolute_url(endpoint.path):
            return endpoint.path
        if not self.base_url:
            raise RuntimeError("INTERNAL_API_BASE_URL is not configured.")
        path = endpoint.path if endpoint.path.startswith("/") else f"/{endpoint.path}"
        return f"{self.base_url}{path}"

    def _build_headers(self, endpoint):
        headers = {"Accept": "application/json", **self.default_headers, **endpoint.headers}
        if self.auth_mode == "basic":
            return headers
        if self.api_key:
            auth_value = self.api_key
            if self.auth_prefix:
                auth_value = f"{self.auth_prefix} {self.api_key}".strip()
            headers[self.auth_header] = auth_value
            headers.setdefault("X-API-Key", self.api_key)
        return headers

    def _build_auth(self):
        if self.auth_mode != "basic":
            return None
        if not self.username:
            raise RuntimeError("INTERNAL_API_USERNAME is required for basic auth mode.")
        return (self.username, self.password)

    @staticmethod
    def _extract_by_path(payload, record_path):
        if not record_path:
            return None

        current = payload
        tokens = []
        for chunk in record_path.split("."):
            if not chunk:
                continue
            parts = re.findall(r"([^.\\[\\]]+)|\\[(\\d+)\\]", chunk)
            for key_token, index_token in parts:
                tokens.append(key_token if key_token else int(index_token))

        for token in tokens:
            if isinstance(token, int):
                if not isinstance(current, list) or token >= len(current):
                    return None
                current = current[token]
            else:
                if not isinstance(current, dict) or token not in current:
                    return None
                current = current[token]

        if isinstance(current, list) and all(isinstance(item, dict) for item in current):
            return current
        return None

    @classmethod
    def _extract_by_keys(cls, payload, record_keys):
        if isinstance(payload, list):
            if payload and all(isinstance(item, dict) for item in payload):
                return payload, "root"
            return None, None
        if not isinstance(payload, dict):
            return None, None

        for key in record_keys:
            if key not in payload:
                continue
            records, path = cls._extract_by_keys(payload.get(key), record_keys)
            if records is not None:
                return records, key if path == "root" else f"{key}.{path}"
        return None, None

    @classmethod
    def _candidate_lists(cls, payload, path="root"):
        candidates = []
        if isinstance(payload, list):
            if payload and all(isinstance(item, dict) for item in payload):
                candidates.append((path, payload))
            for index, item in enumerate(payload[:5]):
                if isinstance(item, (dict, list)):
                    candidates.extend(cls._candidate_lists(item, f"{path}[{index}]"))
            return candidates

        if isinstance(payload, dict):
            for key, value in payload.items():
                child_path = f"{path}.{key}" if path else key
                if isinstance(value, (dict, list)):
                    candidates.extend(cls._candidate_lists(value, child_path))
        return candidates

    @classmethod
    def _score_candidate(cls, path, records, preferred_keys):
        if not records or not all(isinstance(item, dict) for item in records):
            return -1.0

        sample = records[: min(len(records), 5)]
        keys = set()
        for record in sample:
            keys.update(cls._flatten_record(record).keys())

        normalized_keys = {cls._normalize_token(key) for key in keys}
        normalized_path_tokens = {
            cls._normalize_token(token)
            for token in re.split(r"[.\[\]_/\\-]+", path)
            if token
        }
        normalized_preferred = {cls._normalize_token(key) for key in preferred_keys}

        score = 0.0
        score += min(len(records), 200) / 25.0
        score += min(len(keys), 30) / 10.0
        score += len(normalized_path_tokens.intersection(normalized_preferred)) * 5.0
        score += sum(1.0 for hint in DISCOVERY_HINTS if any(hint in key for key in normalized_keys))

        if any(key in normalized_keys for key in ("rating", "score", "comment", "feedback", "review")):
            score += 4.0
        if any(key in normalized_keys for key in ("date", "created_at", "feedback_date", "submitted_at")):
            score += 2.0
        return score

    @classmethod
    def _discover_records(cls, payload, record_keys):
        candidates = cls._candidate_lists(payload)
        if not candidates:
            return None, None

        ranked = sorted(
            candidates,
            key=lambda item: cls._score_candidate(item[0], item[1], record_keys),
            reverse=True,
        )
        best_path, best_records = ranked[0]
        return best_records, best_path

    @classmethod
    def _flatten_scalar(cls, value):
        if value is None:
            return ""
        if isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, list):
            if not value:
                return ""
            if all(not isinstance(item, (dict, list)) for item in value):
                return " | ".join(str(item) for item in value)
            return json.dumps(value, ensure_ascii=False)
        if isinstance(value, dict):
            return json.dumps(value, ensure_ascii=False)
        return str(value)

    @classmethod
    def _flatten_record(cls, record, prefix="", output=None):
        if output is None:
            output = {}
        if not isinstance(record, dict):
            output[prefix or "value"] = cls._flatten_scalar(record)
            return output

        for key, value in record.items():
            field_name = f"{prefix}.{key}" if prefix else str(key)
            if isinstance(value, dict):
                cls._flatten_record(value, field_name, output)
            elif isinstance(value, list) and value and all(isinstance(item, dict) for item in value):
                output[field_name] = json.dumps(value, ensure_ascii=False)
            else:
                output[field_name] = cls._flatten_scalar(value)
        return output

    @classmethod
    def _flatten_records(cls, records):
        return [cls._flatten_record(record) for record in records]

    def request_endpoint(self, endpoint_ref, extra_params=None):
        endpoint = self.get_endpoint(endpoint_ref)
        url = self._resolve_url(endpoint)
        params = {**endpoint.query_params, **(extra_params or {})}
        request_kwargs = {
            "headers": self._build_headers(endpoint),
            "timeout": self.timeout_seconds,
        }
        auth = self._build_auth()
        if auth is not None:
            request_kwargs["auth"] = auth
        if endpoint.method == "GET":
            request_kwargs["params"] = params
            response = requests.get(url, **request_kwargs)
        else:
            if endpoint.body_mode == "form":
                request_kwargs["data"] = params
            else:
                request_kwargs["json"] = params
            response = requests.request(endpoint.method, url, **request_kwargs)
        response.raise_for_status()
        return response.json()

    def interpret_payload(self, endpoint_ref, extra_params=None):
        endpoint = self.get_endpoint(endpoint_ref)
        payload = self.request_endpoint(endpoint, extra_params=extra_params)
        return self.interpret_payload_object(endpoint, payload)

    def interpret_payload_object(self, endpoint_ref, payload):
        endpoint = self.get_endpoint(endpoint_ref)
        records = self._extract_by_path(payload, endpoint.record_path)
        record_path = endpoint.record_path if records is not None else None

        if records is None:
            records, record_path = self._extract_by_keys(payload, endpoint.record_keys)

        if records is None and endpoint.auto_discover:
            records, record_path = self._discover_records(payload, endpoint.record_keys)

        if records is None:
            raise ValueError(
                f"Unsupported payload format for internal endpoint '{endpoint.name}'. "
                f"Expected list data under one of: {', '.join(endpoint.record_keys)}, or an auto-discoverable list of objects."
            )

        flattened_records = self._flatten_records(records)
        return {
            "endpoint": self.describe_endpoint(endpoint),
            "record_path": record_path or "root",
            "record_count": len(records),
            "records": flattened_records,
            "sample_keys": sorted(flattened_records[0].keys()) if flattened_records else [],
        }

    def fetch_records(self, endpoint_ref, extra_params=None):
        return self.interpret_payload(endpoint_ref, extra_params=extra_params)["records"]

    def describe_endpoint(self, endpoint_ref):
        endpoint = self.get_endpoint(endpoint_ref)
        return {
            "name": endpoint.name,
            "path": endpoint.path,
            "method": endpoint.method,
            "body_mode": endpoint.body_mode,
            "record_path": endpoint.record_path,
            "resolved_url": self._resolve_url(endpoint) if self.base_url or self._is_absolute_url(endpoint.path) else endpoint.path,
            "record_keys": list(endpoint.record_keys),
            "query_params": self._redact_mapping(endpoint.query_params),
            "headers": self._redact_mapping(endpoint.headers),
            "auto_discover": endpoint.auto_discover,
        }
