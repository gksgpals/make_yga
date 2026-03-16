import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import auth_support


class AuthSupportTest(unittest.TestCase):
    def test_parse_identifier_set_supports_commas_newlines_and_semicolons(self) -> None:
        parsed = auth_support.parse_identifier_set("a@example.com,\nb@example.com; c@example.com")
        self.assertEqual(parsed, {"a@example.com", "b@example.com", "c@example.com"})

    def test_email_is_allowed_by_exact_email(self) -> None:
        self.assertTrue(
            auth_support.is_email_allowed(
                "Teacher@Example.com",
                {"teacher@example.com"},
                set(),
            )
        )

    def test_email_is_allowed_by_domain(self) -> None:
        self.assertTrue(
            auth_support.is_email_allowed(
                "student@school.ac.kr",
                set(),
                {"school.ac.kr"},
            )
        )

    def test_email_is_rejected_when_not_in_allowlist(self) -> None:
        self.assertFalse(
            auth_support.is_email_allowed(
                "student@other.com",
                {"teacher@example.com"},
                {"school.ac.kr"},
            )
        )

    def test_get_oidc_config_from_env_requires_all_mandatory_values(self) -> None:
        with patch.dict(
            os.environ,
            {
                auth_support.REDIRECT_URI_ENV: "http://localhost:8501/oauth2callback",
                auth_support.COOKIE_SECRET_ENV: "cookie-secret",
            },
            clear=False,
        ):
            with self.assertRaises(ValueError):
                auth_support.get_oidc_config_from_env()

    def test_ensure_streamlit_auth_secrets_writes_expected_toml(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir, patch.dict(
            os.environ,
            {
                auth_support.REDIRECT_URI_ENV: "http://localhost:8501/oauth2callback",
                auth_support.COOKIE_SECRET_ENV: "cookie-secret",
                auth_support.GOOGLE_CLIENT_ID_ENV: "client-id",
                auth_support.GOOGLE_CLIENT_SECRET_ENV: "client-secret",
                auth_support.GOOGLE_SERVER_METADATA_URL_ENV: auth_support.DEFAULT_GOOGLE_SERVER_METADATA_URL,
            },
            clear=False,
        ):
            secrets_path = auth_support.ensure_streamlit_auth_secrets(Path(temp_dir))
            self.assertIsNotNone(secrets_path)
            assert secrets_path is not None
            contents = secrets_path.read_text()

        self.assertIn('[auth]', contents)
        self.assertIn('redirect_uri = "http://localhost:8501/oauth2callback"', contents)
        self.assertIn('client_id = "client-id"', contents)


if __name__ == "__main__":
    unittest.main()
