from dataclasses import dataclass, field
import json
import logging
import os
from typing import Any

from config import INTERNAL_CONNECTOR_PATH
from internal_api import EndpointSpec

logger = logging.getLogger(__name__)

DEFAULT_REQUIRED_FIELDS = (
    "Tipe Stakeholder",
    "Layanan",
    "Tanggal Feedback",
    "Rating",
    "Komentar",
)


@dataclass(frozen=True)
class InternalConnectorEndpoint:
    endpoint_name: str = "feedback"
    enabled: bool = True
    url: str = ""
    method: str = "GET"
    body_mode: str = "json"
    request_data: dict[str, Any] = field(default_factory=dict)
    headers: dict[str, str] = field(default_factory=dict)
    record_path: str = ""
    record_keys: tuple[str, ...] = ("items", "data", "results", "records", "feedback")
    auto_discover: bool = True
    field_map: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_mapping(cls, mapping, fallback_name="feedback"):
        return cls(
            endpoint_name=str(mapping.get("endpoint_name", fallback_name)).strip() or fallback_name,
            enabled=bool(mapping.get("enabled", True)),
            url=str(mapping.get("url", "")).strip(),
            method=str(mapping.get("method", "GET")).strip().upper() or "GET",
            body_mode=str(mapping.get("body_mode", "json")).strip().lower() or "json",
            request_data=dict(mapping.get("request_data") or {}),
            headers=dict(mapping.get("headers") or {}),
            record_path=str(mapping.get("record_path", "")).strip(),
            record_keys=tuple(mapping.get("record_keys") or ("items", "data", "results", "records", "feedback")),
            auto_discover=bool(mapping.get("auto_discover", True)),
            field_map=dict(mapping.get("field_map") or {}),
        )

    def to_endpoint_spec(self):
        path = self.url or self.endpoint_name
        return EndpointSpec(
            name=self.endpoint_name,
            path=path,
            method=self.method,
            body_mode=self.body_mode,
            record_path=self.record_path,
            record_keys=self.record_keys,
            query_params=self.request_data,
            headers=self.headers,
            auto_discover=self.auto_discover,
        )

    def apply_field_map(self, dataframe):
        if dataframe is None or dataframe.empty or not self.field_map:
            return dataframe
        rename_map = {}
        for source_field, target_field in self.field_map.items():
            source_key = str(source_field).strip()
            target_key = str(target_field).strip()
            if source_key in dataframe.columns and target_key:
                rename_map[source_key] = target_key
        if not rename_map:
            return dataframe
        return dataframe.rename(columns=rename_map)

    def describe(self):
        return {
            "endpoint_name": self.endpoint_name,
            "enabled": self.enabled,
            "url": self.url,
            "method": self.method,
            "body_mode": self.body_mode,
            "record_path": self.record_path,
            "record_keys": list(self.record_keys),
            "auto_discover": self.auto_discover,
            "request_data_keys": sorted(self.request_data.keys()),
            "headers": {key: "***redacted***" for key in self.headers.keys()},
            "field_map_keys": sorted(self.field_map.keys()),
        }


@dataclass(frozen=True)
class InternalConnectorSpec:
    name: str = "production_connector"
    enabled: bool = True
    auth_mode: str = "api_key"
    endpoints: tuple[InternalConnectorEndpoint, ...] = field(default_factory=tuple)
    required_fields: tuple[str, ...] = DEFAULT_REQUIRED_FIELDS
    context_enhancer: str = ""

    @classmethod
    def from_mapping(cls, mapping):
        raw_endpoints = mapping.get("endpoints")
        if isinstance(raw_endpoints, list) and raw_endpoints:
            endpoints = tuple(
                InternalConnectorEndpoint.from_mapping(
                    endpoint,
                    fallback_name=f"feedback_{index + 1}",
                )
                for index, endpoint in enumerate(raw_endpoints)
                if isinstance(endpoint, dict)
            )
        else:
            endpoints = (
                InternalConnectorEndpoint.from_mapping(
                    mapping,
                    fallback_name=str(mapping.get("endpoint_name", "feedback")).strip() or "feedback",
                ),
            )

        return cls(
            name=str(mapping.get("name", "production_connector")).strip() or "production_connector",
            enabled=bool(mapping.get("enabled", True)),
            auth_mode=str(mapping.get("auth_mode", "api_key")).strip().lower() or "api_key",
            endpoints=endpoints,
            required_fields=tuple(mapping.get("required_fields") or DEFAULT_REQUIRED_FIELDS),
            context_enhancer=str(mapping.get("context_enhancer", "")).strip(),
        )

    def active_endpoints(self):
        if not self.enabled:
            return ()
        return tuple(endpoint for endpoint in self.endpoints if endpoint.enabled)

    def to_endpoint_spec(self):
        active_endpoints = self.active_endpoints()
        if not active_endpoints:
            raise RuntimeError("Internal connector has no active endpoints.")
        return active_endpoints[0].to_endpoint_spec()

    def apply_field_map(self, dataframe):
        active_endpoints = self.active_endpoints()
        if not active_endpoints:
            return dataframe
        return active_endpoints[0].apply_field_map(dataframe)

    def describe(self):
        return {
            "name": self.name,
            "enabled": self.enabled,
            "auth_mode": self.auth_mode,
            "endpoint_count": len(self.endpoints),
            "active_endpoint_count": len(self.active_endpoints()),
            "endpoints": [endpoint.describe() for endpoint in self.endpoints],
            "required_fields": list(self.required_fields),
            "context_enhancer_configured": bool(self.context_enhancer),
        }


def load_internal_connector(path=INTERNAL_CONNECTOR_PATH):
    connector_path = str(path or "").strip()
    if not connector_path:
        return None
    if not os.path.exists(connector_path):
        return None

    with open(connector_path, "r", encoding="utf-8") as handle:
        payload = json.load(handle)

    if not isinstance(payload, dict):
        raise ValueError("Internal connector file must contain a JSON object.")

    spec = InternalConnectorSpec.from_mapping(payload)
    logger.info("Loaded internal connector spec from %s", connector_path)
    return spec
