from internal_api_settings import (
    build_connector_payload,
    connector_settings_state,
    load_connector_payload,
    write_connector_payload,
)

_BUSY_REFRESH_ERROR = "Sinkronisasi data sementara dikunci karena masih ada laporan yang sedang diproses."
_BUSY_SETTINGS_ERROR = "Pengaturan data internal belum bisa diubah karena masih ada laporan yang sedang diproses."


class InternalApiService:
    def __init__(
        self,
        knowledge_base,
        job_manager,
        logger=None,
        connector_settings_state_func=None,
        build_connector_payload_func=None,
        load_connector_payload_func=None,
        write_connector_payload_func=None,
        connector_settings_state=None,
        build_connector_payload=None,
        load_connector_payload=None,
        write_connector_payload=None,
    ):
        self.knowledge_base = knowledge_base
        self.job_manager = job_manager
        self.logger = logger
        self.connector_settings_state = (
            connector_settings_state_func
            or connector_settings_state
            or globals()["connector_settings_state"]
        )
        self.build_connector_payload = (
            build_connector_payload_func
            or build_connector_payload
            or globals()["build_connector_payload"]
        )
        self.load_connector_payload = (
            load_connector_payload_func
            or load_connector_payload
            or globals()["load_connector_payload"]
        )
        self.write_connector_payload = (
            write_connector_payload_func
            or write_connector_payload
            or globals()["write_connector_payload"]
        )

    def _has_active_jobs(self):
        job_stats = self.job_manager.stats()
        jobs = job_stats.get("jobs", {})
        return bool(jobs.get("queued") or jobs.get("running"))

    def ensure_mutation_allowed(self, message=_BUSY_SETTINGS_ERROR):
        if self._has_active_jobs():
            return False, {"error": message}, 409
        return True, {}, 200

    def can_mutate(self, message=_BUSY_SETTINGS_ERROR):
        allowed, body, status = self.ensure_mutation_allowed(message)
        if not allowed:
            return False, {"status": "busy", **body}, status
        return True, body, status

    def get_state(self):
        state = self.connector_settings_state()
        provider_source = getattr(self.knowledge_base.provider, "source_name", "")
        connector_configured = bool(state.get("connector_exists") and state.get("enabled"))
        project_data_source = "api" if provider_source == "company_api" else "local"
        api_connection_active = connector_configured and project_data_source == "api"
        state.update(
            {
                "project_data_source": project_data_source,
                "active_runtime_source": provider_source or "unknown",
                "api_connection_active": api_connection_active,
                "can_refresh_dataset": api_connection_active,
                "refresh_running": False,
                "connection_label": (
                    "Aktif memakai Internal API/APIDog"
                    if api_connection_active
                    else "Belum aktif untuk sesi berjalan"
                ),
                "active_record_count": int(len(self.knowledge_base.df))
                if getattr(self.knowledge_base, "df", None) is not None
                else 0,
            }
        )
        return state

    def refresh_knowledge(self):
        allowed, body, status = self.ensure_mutation_allowed(_BUSY_REFRESH_ERROR)
        if not allowed:
            return {"status": "busy", **body}, status
        success = self.knowledge_base.refresh_data()
        return {"status": "success" if success else "error"}, 200

    def refresh_dataset(self):
        allowed, body, status = self.ensure_mutation_allowed(_BUSY_REFRESH_ERROR)
        if not allowed:
            return {"status": "busy", **body}, status

        connector_state = self.connector_settings_state()
        if not (connector_state.get("connector_exists") and connector_state.get("enabled")):
            return {
                "status": "not_configured",
                "error": "Internal API belum aktif. Simpan konfigurasi endpoint sebelum refresh dataset.",
                **self.get_state(),
            }, 400

        try:
            self.knowledge_base.activate_internal_api_provider()
            refresh_success = self.knowledge_base.refresh_data()
        except Exception as exc:
            if self.logger:
                self.logger.exception("Failed to refresh Internal API dataset.")
            return {"status": "error", "error": str(exc), **self.get_state()}, 500

        return {
            "status": "refreshed" if refresh_success else "error",
            "refresh_status": "success" if refresh_success else "degraded",
            **self.get_state(),
        }, 200 if refresh_success else 503

    def save_and_refresh(self, data):
        allowed, body, status = self.ensure_mutation_allowed(_BUSY_SETTINGS_ERROR)
        if not allowed:
            return body, status

        try:
            payload = self.build_connector_payload(data, existing_payload=self.load_connector_payload())
            self.write_connector_payload(payload)
            if payload.get("enabled", True):
                self.knowledge_base.activate_internal_api_provider()
            refresh_success = self.knowledge_base.refresh_data()
        except ValueError as exc:
            return {"error": str(exc)}, 400
        except Exception as exc:
            if self.logger:
                self.logger.exception("Failed to update Internal API settings.")
            return {"error": str(exc)}, 500

        return {
            "status": "saved",
            "refresh_status": "success" if refresh_success else "degraded",
            **self.get_state(),
        }, 200
