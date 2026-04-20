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
class InternalConnectorSpec:
    name: str = "production_connector"
    enabled: bool = True
    endpoint_name: str = "feedback"
    url: str = ""
    method: str = "GET"
    body_mode: str = "json"
    request_data: dict[str, Any] = field(default_factory=dict)
    headers: dict[str, str] = field(default_factory=dict)
    record_path: str = ""
    record_keys: tuple[str, ...] = ("items", "data", "results", "records", "feedback")
    auto_discover: bool = False
    field_map: dict[str, str] = field(default_factory=dict)
    required_fields: tuple[str, ...] = DEFAULT_REQUIRED_FIELDS

    @classmethod
    def from_mapping(cls, mapping):
        return cls(
            name=str(mapping.get("name", "production_connector")).strip() or "production_connector",
            enabled=bool(mapping.get("enabled", True)),
            endpoint_name=str(mapping.get("endpoint_name", "feedback")).strip() or "feedback",
            url=str(mapping.get("url", "")).strip(),
            method=str(mapping.get("method", "GET")).strip().upper() or "GET",
            body_mode=str(mapping.get("body_mode", "json")).strip().lower() or "json",
            request_data=dict(mapping.get("request_data") or {}),
            headers=dict(mapping.get("headers") or {}),
            record_path=str(mapping.get("record_path", "")).strip(),
            record_keys=tuple(mapping.get("record_keys") or ("items", "data", "results", "records", "feedback")),
            auto_discover=bool(mapping.get("auto_discover", False)),
            field_map=dict(mapping.get("field_map") or {}),
            required_fields=tuple(mapping.get("required_fields") or DEFAULT_REQUIRED_FIELDS),
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
            "name": self.name,
            "enabled": self.enabled,
            "endpoint_name": self.endpoint_name,
            "url": self.url,
            "method": self.method,
            "body_mode": self.body_mode,
            "record_path": self.record_path,
            "record_keys": list(self.record_keys),
            "auto_discover": self.auto_discover,
            "request_data_keys": sorted(self.request_data.keys()),
            "field_map_keys": sorted(self.field_map.keys()),
            "required_fields": list(self.required_fields),
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
