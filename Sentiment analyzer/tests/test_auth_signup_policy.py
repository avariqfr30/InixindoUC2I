import os
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]


class SignupPolicyTests(unittest.TestCase):
    def _run_probe(self, extra_env, username="publicuser@company.example"):
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
                    "SIGNUP_ALLOWED_EMAIL_DOMAIN": "@company.example",
                    "REFERENCE_INTERNAL_ACCOUNT_TEST_EMAILS": "publicuser@company.example",
                    "AUTH_SIGNUP_VERIFICATION_DELIVERY_MODE": "capture",
                    "DISABLE_CSRF_FOR_TESTING": "1",
                }
            )
            env.pop("ALLOW_SIGNUP", None)
            env.update(extra_env)

            probe = textwrap.dedent(
                """
                import app as app_module
                from auth_service import get_user_by_username

                delivered = {}

                def capture_delivery(email, verification_token, initial_password, user_fullname=""):
                    delivered.update(
                        email=email,
                        verification_token=verification_token,
                        initial_password=initial_password,
                    )
                    return dict(delivered)

                app_module.send_signup_verification_email = capture_delivery

                client = app_module.app.test_client()
                get_response = client.get("/signup", follow_redirects=False)
                post_response = client.post(
                    "/signup",
                    data={"username": "%s"},
                    follow_redirects=False,
                )
                post_body = post_response.get_data(as_text=True)
                user_before_verification = get_user_by_username("%s")
                initial_password = delivered.get("initial_password", "Password123!")
                pre_verify_login = client.post(
                    "/login",
                    data={"username": "%s", "password": initial_password},
                    follow_redirects=False,
                )
                verification_response = None
                if delivered.get("verification_token"):
                    verification_response = client.post(
                        "/verify-otp",
                        data={
                            "username": "%s",
                            "verification_token": delivered["verification_token"],
                        },
                        follow_redirects=False,
                    )
                user_after_verification = get_user_by_username("%s")
                final_login = client.post(
                    "/login",
                    data={"username": "%s", "password": initial_password},
                    follow_redirects=False,
                )
                print(
                    {
                        "get_status": get_response.status_code,
                        "post_status": post_response.status_code,
                        "post_location": post_response.headers.get("Location", ""),
                        "pre_verify_login_status": pre_verify_login.status_code,
                        "verification_status": verification_response.status_code if verification_response else None,
                        "verification_location": verification_response.headers.get("Location", "") if verification_response else "",
                        "login_status": final_login.status_code,
                        "login_location": final_login.headers.get("Location", ""),
                        "user_count": app_module.user_count(),
                        "approved_before_verification": bool(
                            user_before_verification and user_before_verification["approved_at"]
                        ),
                        "approved_after_verification": bool(
                            user_after_verification and user_after_verification["approved_at"]
                        ),
                        "delivery_captured": bool(delivered),
                        "approval_gate": app_module.SIGNUP_REQUIRES_APPROVAL,
                        "has_verification_form": "verification_token" in post_body,
                        "has_warning_popup": "signup-warning-popup" in post_body,
                    }
                )
                """
                % (username, username, username, username, username, username)
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

    def test_internal_email_signup_requires_otp_then_can_login(self):
        output = self._run_probe({})

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'pre_verify_login_status': 200", output)
        self.assertIn("'verification_status': 302", output)
        self.assertIn("'verification_location': '/login'", output)
        self.assertIn("'login_status': 302", output)
        self.assertIn("'login_location': '/'", output)
        self.assertIn("'user_count': 1", output)
        self.assertIn("'approved_before_verification': False", output)
        self.assertIn("'approved_after_verification': True", output)
        self.assertIn("'delivery_captured': True", output)
        self.assertIn("'has_verification_form': True", output)
        self.assertIn("'has_warning_popup': False", output)

    def test_signup_rejects_external_email_domains(self):
        output = self._run_probe({}, username="outsider@example.com")

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'delivery_captured': False", output)
        self.assertIn("'user_count': 0", output)
        self.assertIn("'approved_after_verification': False", output)
        self.assertIn("'has_warning_popup': True", output)

    def test_signup_rejects_invalid_email_format(self):
        output = self._run_probe({}, username="not-an-email")

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'delivery_captured': False", output)
        self.assertIn("'user_count': 0", output)
        self.assertIn("'approved_after_verification': False", output)
        self.assertIn("'has_warning_popup': True", output)

    def test_signup_can_still_be_closed_explicitly(self):
        output = self._run_probe({"ALLOW_SIGNUP": "0"})

        self.assertIn("'get_status': 302", output)
        self.assertIn("'post_status': 302", output)
        self.assertIn("'post_location': '/login'", output)
        self.assertIn("'user_count': 0", output)

    def test_approval_gate_still_requires_verification_before_approval(self):
        output = self._run_probe({"SIGNUP_REQUIRES_APPROVAL": "1"})

        self.assertIn("'get_status': 200", output)
        self.assertIn("'post_status': 200", output)
        self.assertIn("'approval_gate': True", output)
        self.assertIn("'pre_verify_login_status': 200", output)
        self.assertIn("'verification_status': 302", output)
        self.assertIn("'user_count': 1", output)
        self.assertIn("'approved_before_verification': False", output)
        self.assertIn("'approved_after_verification': True", output)


if __name__ == "__main__":
    unittest.main()
