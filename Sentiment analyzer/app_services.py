from dataclasses import dataclass

from config import DEFAULT_SCORE_ENGINE, SIGNUP_ALLOWED_EMAIL_DOMAIN


@dataclass(frozen=True)
class ReportRequest:
    timeframe: str
    notes: str = ""
    sentiment: str = "all"
    segment: str = "all"
    score_engine: str = DEFAULT_SCORE_ENGINE

    @classmethod
    def from_mapping(cls, data):
        data = data or {}
        return cls(
            timeframe=str(data.get("timeframe") or "").strip(),
            notes=str(data.get("notes") or "").strip(),
            sentiment=str(data.get("sentiment") or "all").strip() or "all",
            segment=str(data.get("segment") or "all").strip() or "all",
            score_engine=str(data.get("score_engine") or DEFAULT_SCORE_ENGINE).strip()
            or DEFAULT_SCORE_ENGINE,
        )

    def validate(self):
        if not self.timeframe:
            raise ValueError("Parameter 'timeframe' wajib diisi.")
        return self

    def to_job_payload(self):
        return {
            "timeframe": self.timeframe,
            "notes": self.notes,
            "sentiment": self.sentiment,
            "segment": self.segment,
            "score_engine": self.score_engine,
        }


def report_request_payload(data):
    return ReportRequest.from_mapping(data).to_job_payload()


def allowed_signup_domain():
    domain = str(SIGNUP_ALLOWED_EMAIL_DOMAIN or "").strip().lower()
    if domain and not domain.startswith("@"):
        domain = f"@{domain}"
    return domain


def is_allowed_signup_email(value):
    email = str(value or "").strip().lower()
    allowed_domain = allowed_signup_domain()
    return bool(email and allowed_domain and email.endswith(allowed_domain))
