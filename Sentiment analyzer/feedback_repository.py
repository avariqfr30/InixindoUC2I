import json
from datetime import datetime, timezone

import pandas as pd


FEEDBACK_SCHEMA_VERSION = 1
FEEDBACK_INDEX_SQL = (
    'CREATE INDEX IF NOT EXISTS idx_feedback_timeframe ON feedback ("Rentang Waktu", "Tanggal Feedback")',
    'CREATE INDEX IF NOT EXISTS idx_feedback_source ON feedback ("Sumber Feedback", "Kanal Feedback")',
    'CREATE INDEX IF NOT EXISTS idx_feedback_service ON feedback ("Layanan", "Tipe Stakeholder")',
    'CREATE INDEX IF NOT EXISTS idx_feedback_journey ON feedback ("Customer Journey Hint")',
)


class FeedbackRepository:
    def __init__(self, engine):
        self.engine = engine

    @staticmethod
    def _utc_now():
        return (
            datetime.now(timezone.utc)
            .replace(microsecond=0)
            .isoformat()
            .replace("+00:00", "Z")
        )

    def ensure_schema(self):
        with self.engine.begin() as connection:
            connection.exec_driver_sql(
                """
                CREATE TABLE IF NOT EXISTS feedback_schema_version (
                    version INTEGER NOT NULL,
                    applied_at TEXT NOT NULL
                )
                """
            )
            current = connection.exec_driver_sql(
                "SELECT version FROM feedback_schema_version ORDER BY version DESC LIMIT 1"
            ).fetchone()
            if current is None:
                connection.exec_driver_sql(
                    "INSERT INTO feedback_schema_version (version, applied_at) VALUES (?, ?)",
                    (FEEDBACK_SCHEMA_VERSION, self._utc_now()),
                )
            connection.exec_driver_sql(
                """
                CREATE TABLE IF NOT EXISTS feedback_ingestion_runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    source_name TEXT NOT NULL,
                    connector_name TEXT,
                    dataset_summary TEXT,
                    row_count INTEGER NOT NULL,
                    ingested_at TEXT NOT NULL
                )
                """
            )

    @staticmethod
    def _ensure_feedback_indexes(connection):
        for statement in FEEDBACK_INDEX_SQL:
            connection.exec_driver_sql(statement)

    def save_feedback_dataframe(self, dataframe, source_name="", metadata=None):
        metadata = metadata or {}
        self.ensure_schema()
        dataframe.to_sql("feedback", self.engine, index=False, if_exists="replace")
        with self.engine.begin() as connection:
            self._ensure_feedback_indexes(connection)
            connection.exec_driver_sql(
                """
                INSERT INTO feedback_ingestion_runs (
                    source_name,
                    connector_name,
                    dataset_summary,
                    row_count,
                    ingested_at
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    str(source_name or "unknown"),
                    str(metadata.get("connector") or ""),
                    json.dumps(
                        metadata.get("datasets") or [],
                        ensure_ascii=False,
                        sort_keys=True,
                    ),
                    int(len(dataframe) if dataframe is not None else 0),
                    self._utc_now(),
                ),
            )

    def load_feedback_dataframe(self):
        self.ensure_schema()
        return pd.read_sql("SELECT * FROM feedback", self.engine)
