import os
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]


class SignupPolicyTests(unittest.TestCase):
    def _run_probe(self, extra_env, username="publicuser@inixindojogja.co.id"):
        with tempfile.TemporaryDirectory() as temp_dir:
            env = os.environ.copy()
            env.update(
                {
                    "APP_PROFILE": "demo",
                    "APP_MODE": "demo",
                    "AUTH_DB_PATH": str(Path(temp_dir) / "auth.db"),
                    "DB_URI": f"sqlite:///{Path(temp_dir) / 'cx_feedback.db'}",
                    "JOB_STATE_PATH": str(Path(temp_dir) / "jobs.json"),
                    "REPORT_ARTIFACT_DIR": str(Path(temp_dir) / "reports"),
                    "INTERNAL_CONNECTOR_PATH": str(Path(temp_dir) / "internal_api_config.json"),
                }
            )
            env.pop("ALLOW_SIGNUP", None)
            env.update(extra_env)

            probe = textwrap.dedent(
                """
                from app import app, user_count
                from auth_service import get_user_by_username

                client = app.test_client()
                get_response = client.get("/signup", follow_redirects=False)
                post_response = client.post(
                    "/signup",
                    data={
                        "username": "%s",
                        "password": "Password123!",
                        "confirm_password": "Password123!",
                    },
                    follow_redirects=False,
                )
                post_body = post_response.get_data(as_text=True)
                login_response = client.post(
                    "/login",
                    data={"username": "%s", "password": "Password123!"},
                    follow_redirects=False,
                )
                user = get_user_by_username("%s")
                approved_at = user["approved_at"] if user and "approved_at" in user.keys() else None
                print(
                    {
                        "get_status": get_response.status_code,
                        "post_status": post_response.status_code,
                        "post_location": post_response.headers.get("Location", ""),
                        "login_status": login_response.status_code,
                        "login_location": login_response.headers.get("Location", ""),
                        "user_count": user_count(),
                        "approved": bool(approved_at),
                        "has_warning_popup": "signup-warning-popup" in post_body,
                        "has_error_box": "class=\\"error\\"" in post_body,
                    }
                )
                """ % (username, username, username)
            )
            result = subprocess.run(
                [sys.executable, "-c", probe],
                cwd=str(PROJECT_DIR),
                env=env,
                text=True,
                capture_output=True,
                check=True,
            )
            return result.stdout

    def test_internal_email_signup_can_login_by_default(self):
        output = self._run_probe({})

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 302", output)
        self.assertIn("'post_location': '/'", output)
        self.assertIn("'login_status': 302", output)
        self.assertIn("'login_location': '/'", output)
        self.assertIn("'user_count': 1", output)
        self.assertIn("'approved': True", output)
        self.assertIn("'has_warning_popup': False", output)

    def test_signup_rejects_external_email_domains(self):
        output = self._run_probe({}, username="outsider@example.com")

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'login_status': 200", output)
        self.assertIn("'user_count': 0", output)
        self.assertIn("'approved': False", output)
        self.assertIn("'has_warning_popup': True", output)
        self.assertIn("'has_error_box': False", output)

    def test_signup_rejects_invalid_email_format(self):
        output = self._run_probe({}, username="not-an-email")

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'login_status': 200", output)
        self.assertIn("'user_count': 0", output)
        self.assertIn("'approved': False", output)
        self.assertIn("'has_warning_popup': True", output)
        self.assertIn("'has_error_box': False", output)

    def test_signup_can_still_be_closed_explicitly(self):
        output = self._run_probe({"ALLOW_SIGNUP": "0"})

        self.assertIn("'get_status': 302", output)
        self.assertIn("'post_status': 302", output)
        self.assertIn("'post_location': '/login'", output)
        self.assertIn("'user_count': 0", output)

    def test_approval_gate_can_still_be_enabled_explicitly(self):
        output = self._run_probe({"SIGNUP_REQUIRES_APPROVAL": "1"})

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'login_status': 200", output)
        self.assertIn("'login_location': ''", output)
        self.assertIn("'user_count': 1", output)
        self.assertIn("'approved': False", output)


if __name__ == "__main__":
    unittest.main()
