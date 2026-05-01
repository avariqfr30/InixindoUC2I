import json
import os
import re

from config import INTERNAL_CONNECTOR_PATH
from internal_connector import DEFAULT_REQUIRED_FIELDS, InternalConnectorSpec

DEFAULT_RECORD_KEYS = ("dataset_result", "items", "data", "results", "records", "feedback")
SUPPORTED_METHODS = {"GET", "POST", "PUT", "PATCH"}
SUPPORTED_BODY_MODES = {"json", "form"}
SUPPORTED_AUTH_MODES = {"bearer_env", "basic_env", "none"}
DEFAULT_APIDOG_DATASETS = (
    ("class_report", "ClassReport"),
    ("reference_class_report", "ReferenceClassReport"),
)
DEFAULT_FIELD_MAP = {
    "id": "Record ID",
    "record_id": "Record ID",
    "feedback_id": "Record ID",
    "class_report_id": "Record ID",
    "class_id": "Record ID",
    "stakeholder_type": "Tipe Stakeholder",
    "stakeholder": "Tipe Stakeholder",
    "tipe_stakeholder": "Tipe Stakeholder",
    "participant_type": "Tipe Stakeholder",
    "customer_type": "Tipe Stakeholder",
    "company_segment": "Tipe Stakeholder",
    "service_name": "Layanan",
    "service": "Layanan",
    "layanan": "Layanan",
    "class_name": "Layanan",
    "training_name": "Layanan",
    "course_name": "Layanan",
    "product_name": "Layanan",
    "program_name": "Layanan",
    "feedback_date": "Tanggal Feedback",
    "date": "Tanggal Feedback",
    "tanggal_feedback": "Tanggal Feedback",
    "class_date": "Tanggal Feedback",
    "report_date": "Tanggal Feedback",
    "submitted_at": "Tanggal Feedback",
    "created_at": "Tanggal Feedback",
    "rating": "Rating",
    "score": "Rating",
    "nilai": "Rating",
    "satisfaction_score": "Rating",
    "class_rating": "Rating",
    "comment": "Komentar",
    "comments": "Komentar",
    "komentar": "Komentar",
    "feedback": "Komentar",
    "feedback_text": "Komentar",
    "suggestion": "Komentar",
    "instructor_name": "PIC Layanan",
    "trainer_name": "PIC Layanan",
    "location": "Lokasi",
    "venue": "Lokasi",
    "city": "Lokasi",
}


def _load_json_object(raw_value, field_name):
    if isinstance(raw_value, dict):
        return raw_value
    clean_value = str(raw_value or "").strip()
    if not clean_value:
        return {}
    try:
        parsed = json.loads(clean_value)
    except json.JSONDecodeError as exc:
        raise ValueError(f"{field_name} harus berupa JSON object valid.") from exc
    if not isinstance(parsed, dict):
        raise ValueError(f"{field_name} harus berupa JSON object.")
    return parsed


def _normalize_record_keys(raw_value):
    if isinstance(raw_value, (list, tuple)):
        values = raw_value
    else:
        values = re.split(r"[\n,]+", str(raw_value or ""))
    clean_values = [str(value).strip() for value in values if str(value).strip()]
    return clean_values or list(DEFAULT_RECORD_KEYS)


def _normalize_auth_mode(raw_value):
    auth_mode = str(raw_value or "bearer_env").strip().lower()
    return auth_mode if auth_mode in SUPPORTED_AUTH_MODES else "bearer_env"


def _existing_headers_by_endpoint(existing_payload):
    if not isinstance(existing_payload, dict):
        return {}
    connector = InternalConnectorSpec.from_mapping(existing_payload)
    return {
        endpoint.endpoint_name: dict(endpoint.headers)
        for endpoint in connector.endpoints
        if endpoint.headers
    }


def _normalize_endpoint(raw_endpoint, index, existing_headers=None):
    endpoint = raw_endpoint if isinstance(raw_endpoint, dict) else {}
    endpoint_name = str(endpoint.get("endpoint_name") or f"feedback_{index + 1}").strip()
    url = str(endpoint.get("url") or "").strip()
    if not re.match(r"^https?://", url, flags=re.IGNORECASE):
        raise ValueError(f"Endpoint {index + 1} harus memakai URL HTTP/HTTPS penuh.")

    method = str(endpoint.get("method") or "GET").strip().upper()
    if method not in SUPPORTED_METHODS:
        raise ValueError(f"Endpoint {index + 1} memakai method tidak didukung: {method}.")

    body_mode = str(endpoint.get("body_mode") or "json").strip().lower()
    if body_mode not in SUPPORTED_BODY_MODES:
        raise ValueError(f"Endpoint {index + 1} memakai body mode tidak didukung: {body_mode}.")

    raw_headers = endpoint.get("headers")
    if not str(raw_headers or "").strip() and existing_headers:
        headers = dict(existing_headers)
    else:
        headers = _load_json_object(raw_headers, f"headers endpoint {index + 1}")

    return {
        "endpoint_name": endpoint_name or f"feedback_{index + 1}",
        "enabled": bool(endpoint.get("enabled", True)),
        "url": url,
        "method": method,
        "body_mode": body_mode,
        "request_data": _load_json_object(endpoint.get("request_data"), f"request_data endpoint {index + 1}"),
        "headers": headers,
        "record_path": str(endpoint.get("record_path") or "").strip(),
        "record_keys": _normalize_record_keys(endpoint.get("record_keys")),
        "auto_discover": bool(endpoint.get("auto_discover", True)),
        "field_map": _load_json_object(endpoint.get("field_map"), f"field_map endpoint {index + 1}"),
    }


def _dataset_codes_from_payload(data):
    raw_datasets = data.get("datasets")
    if isinstance(raw_datasets, list):
        clean_codes = [str(code).strip() for code in raw_datasets if str(code).strip()]
        if clean_codes:
            return [
                (re.sub(r"[^a-z0-9]+", "_", code.lower()).strip("_") or f"dataset_{index + 1}", code)
                for index, code in enumerate(clean_codes)
            ]
    raw_dataset = str(data.get("dataset") or "").strip()
    if raw_dataset:
        return [(str(data.get("endpoint_name") or "feedback").strip() or "feedback", raw_dataset)]
    return list(DEFAULT_APIDOG_DATASETS)


def _simple_endpoints_from_payload(data):
    url = str(data.get("url") or data.get("endpoint_url") or "").strip()
    if not url:
        return []
    body_mode = str(data.get("body_mode") or "form").strip().lower()
    request_data = _load_json_object(data.get("request_data"), "request_data") if data.get("request_data") else {}
    endpoints = []
    for endpoint_name, dataset_code in _dataset_codes_from_payload(data):
        dataset_request_data = dict(request_data)
        dataset_request_data.setdefault("dataset", dataset_code)
        endpoints.append(
            {
                "endpoint_name": endpoint_name,
                "enabled": True,
                "url": url,
                "method": str(data.get("method") or "POST").strip().upper(),
                "body_mode": body_mode if body_mode in SUPPORTED_BODY_MODES else "form",
                "request_data": dataset_request_data,
                "headers": data.get("headers") or "",
                "record_path": str(data.get("record_path") or "data.dataset_result").strip(),
                "record_keys": data.get("record_keys") or list(DEFAULT_RECORD_KEYS),
                "auto_discover": True,
                "field_map": data.get("field_map") or DEFAULT_FIELD_MAP,
            }
        )
    return endpoints


def build_connector_payload(data, existing_payload=None):
    data = data if isinstance(data, dict) else {}
    raw_endpoints = data.get("endpoints")
    if not isinstance(raw_endpoints, list):
        raw_endpoints = []
    simple_endpoints = _simple_endpoints_from_payload(data)
    if simple_endpoints and not raw_endpoints:
        raw_endpoints = simple_endpoints
    existing_headers = _existing_headers_by_endpoint(existing_payload)
    endpoints = [
        _normalize_endpoint(
            endpoint,
            index,
            existing_headers=existing_headers.get(str(endpoint.get("endpoint_name") or f"feedback_{index + 1}").strip()),
        )
        for index, endpoint in enumerate(raw_endpoints)
        if isinstance(endpoint, dict) and endpoint.get("enabled", True) is not False
    ]
    if not endpoints:
        raise ValueError("Minimal satu endpoint Internal API harus diisi.")

    return {
        "name": str(data.get("name") or "ui_internal_api").strip() or "ui_internal_api",
        "enabled": bool(data.get("enabled", True)),
        "auth_mode": _normalize_auth_mode(data.get("auth_mode")),
        "endpoints": endpoints,
        "required_fields": list(DEFAULT_REQUIRED_FIELDS),
        "context_enhancer": str(data.get("context_enhancer") or "").strip(),
    }


def write_connector_payload(payload, path=INTERNAL_CONNECTOR_PATH):
    connector_path = str(path or "").strip()
    if not connector_path:
        raise ValueError("INTERNAL_CONNECTOR_PATH belum dikonfigurasi.")
    connector_dir = os.path.dirname(connector_path)
    if connector_dir:
        os.makedirs(connector_dir, exist_ok=True)
    with open(connector_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return connector_path


def load_connector_payload(path=INTERNAL_CONNECTOR_PATH):
    connector_path = str(path or "").strip()
    if not connector_path or not os.path.exists(connector_path):
        return None
    with open(connector_path, "r", encoding="utf-8") as handle:
        payload = json.load(handle)
    return payload if isinstance(payload, dict) else None


def _endpoint_to_settings(endpoint):
    return {
        "endpoint_name": endpoint.endpoint_name,
        "enabled": endpoint.enabled,
        "url": endpoint.url,
        "method": endpoint.method,
        "body_mode": endpoint.body_mode,
        "request_data": json.dumps(endpoint.request_data, ensure_ascii=False, indent=2),
        "headers": "",
        "headers_configured": bool(endpoint.headers),
        "record_path": endpoint.record_path,
        "record_keys": ", ".join(endpoint.record_keys),
        "auto_discover": endpoint.auto_discover,
        "field_map": json.dumps(endpoint.field_map, ensure_ascii=False, indent=2),
    }


def _mapping_status(endpoint, required_fields):
    mapped_targets = {str(target).strip() for target in endpoint.field_map.values() if str(target).strip()}
    missing = [field for field in required_fields if field not in mapped_targets]
    return {
        "name": endpoint.endpoint_name,
        "status": "ok" if not missing else "perlu mapping",
        "missing_fields": missing,
    }


def connector_settings_state(path=INTERNAL_CONNECTOR_PATH):
    payload = load_connector_payload(path)
    connector = InternalConnectorSpec.from_mapping(payload) if payload else None
    required_fields = list(connector.required_fields if connector else DEFAULT_REQUIRED_FIELDS)
    return {
        "connector_exists": payload is not None,
        "enabled": connector.enabled if connector else True,
        "auth_mode": _normalize_auth_mode(payload.get("auth_mode") if payload else None),
        "body_mode": connector.endpoints[0].body_mode if connector and connector.endpoints else "form",
        "project_data_source": "api" if connector and connector.enabled and connector.active_endpoints() else "local",
        "connector_path": str(path or ""),
        "context_enhancer": connector.context_enhancer if connector else "",
        "endpoints": [_endpoint_to_settings(endpoint) for endpoint in connector.endpoints] if connector else [],
        "resources": [
            _mapping_status(endpoint, required_fields)
            for endpoint in connector.endpoints
        ] if connector else [
            {"name": "feedback", "status": "perlu mapping", "missing_fields": required_fields}
        ],
    }
