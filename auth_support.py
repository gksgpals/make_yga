"""Helpers for Streamlit OIDC authentication and authorization."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

DEFAULT_GOOGLE_SERVER_METADATA_URL = "https://accounts.google.com/.well-known/openid-configuration"
GOOGLE_PROVIDER_NAME = "google"
AUTH_REQUIRED_ENV = "YGA_REQUIRE_LOGIN"
ALLOWED_EMAILS_ENV = "YGA_ALLOWED_EMAILS"
ALLOWED_EMAIL_DOMAINS_ENV = "YGA_ALLOWED_EMAIL_DOMAINS"
REDIRECT_URI_ENV = "STREAMLIT_AUTH_REDIRECT_URI"
COOKIE_SECRET_ENV = "STREAMLIT_AUTH_COOKIE_SECRET"
GOOGLE_CLIENT_ID_ENV = "STREAMLIT_AUTH_GOOGLE_CLIENT_ID"
GOOGLE_CLIENT_SECRET_ENV = "STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET"
GOOGLE_SERVER_METADATA_URL_ENV = "STREAMLIT_AUTH_GOOGLE_SERVER_METADATA_URL"
STREAMLIT_SECRETS_RELATIVE_PATH = Path(".streamlit/secrets.toml")


@dataclass(frozen=True)
class OIDCConfig:
    redirect_uri: str
    cookie_secret: str
    client_id: str
    client_secret: str
    server_metadata_url: str


def env_flag_enabled(name: str, default: bool = False) -> bool:
    raw_value = os.environ.get(name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() in {"1", "true", "yes", "on"}


def parse_identifier_set(raw_value: str | None) -> set[str]:
    if not raw_value:
        return set()

    normalized = raw_value.replace("\n", ",").replace(";", ",")
    return {item.strip().lower() for item in normalized.split(",") if item.strip()}


def normalize_email(value: str) -> str:
    return value.strip().lower()


def normalize_domain(value: str) -> str:
    return value.strip().lower().lstrip("@")


def get_allowed_emails() -> set[str]:
    return {normalize_email(item) for item in parse_identifier_set(os.environ.get(ALLOWED_EMAILS_ENV))}


def get_allowed_domains() -> set[str]:
    return {
        normalize_domain(item)
        for item in parse_identifier_set(os.environ.get(ALLOWED_EMAIL_DOMAINS_ENV))
    }


def is_email_allowed(
    email: str,
    allowed_emails: set[str] | None = None,
    allowed_domains: set[str] | None = None,
) -> bool:
    normalized_email = normalize_email(email)
    if not normalized_email:
        return False

    resolved_allowed_emails = allowed_emails if allowed_emails is not None else get_allowed_emails()
    resolved_allowed_domains = allowed_domains if allowed_domains is not None else get_allowed_domains()
    if not resolved_allowed_emails and not resolved_allowed_domains:
        return True
    if normalized_email in resolved_allowed_emails:
        return True
    _, separator, domain = normalized_email.partition("@")
    if not separator:
        return False
    return normalize_domain(domain) in resolved_allowed_domains


def get_oidc_config_from_env() -> OIDCConfig | None:
    redirect_uri = os.environ.get(REDIRECT_URI_ENV, "").strip()
    cookie_secret = os.environ.get(COOKIE_SECRET_ENV, "").strip()
    client_id = os.environ.get(GOOGLE_CLIENT_ID_ENV, "").strip()
    client_secret = os.environ.get(GOOGLE_CLIENT_SECRET_ENV, "").strip()

    required_values = [redirect_uri, cookie_secret, client_id, client_secret]
    if not any(required_values):
        return None
    if not all(required_values):
        raise ValueError(
            "OIDC 환경변수가 불완전합니다. "
            "redirect URI, cookie secret, client id, client secret를 모두 설정하세요."
        )

    server_metadata_url = os.environ.get(
        GOOGLE_SERVER_METADATA_URL_ENV,
        DEFAULT_GOOGLE_SERVER_METADATA_URL,
    ).strip() or DEFAULT_GOOGLE_SERVER_METADATA_URL

    return OIDCConfig(
        redirect_uri=redirect_uri,
        cookie_secret=cookie_secret,
        client_id=client_id,
        client_secret=client_secret,
        server_metadata_url=server_metadata_url,
    )


def auth_is_required() -> bool:
    if env_flag_enabled(AUTH_REQUIRED_ENV, default=False):
        return True
    try:
        return get_oidc_config_from_env() is not None
    except ValueError:
        return True


def build_auth_secrets_toml(config: OIDCConfig) -> str:
    return "\n".join(
        [
            "[auth]",
            f'redirect_uri = "{config.redirect_uri}"',
            f'cookie_secret = "{config.cookie_secret}"',
            "",
            f"[auth.{GOOGLE_PROVIDER_NAME}]",
            f'client_id = "{config.client_id}"',
            f'client_secret = "{config.client_secret}"',
            f'server_metadata_url = "{config.server_metadata_url}"',
            "",
        ]
    )


def ensure_streamlit_auth_secrets(base_dir: Path | None = None) -> Path | None:
    config = get_oidc_config_from_env()
    if config is None:
        return None

    target_root = base_dir if base_dir is not None else Path.cwd()
    secrets_path = target_root / STREAMLIT_SECRETS_RELATIVE_PATH
    secrets_path.parent.mkdir(parents=True, exist_ok=True)

    rendered = build_auth_secrets_toml(config)
    existing = secrets_path.read_text() if secrets_path.exists() else None
    if existing != rendered:
        secrets_path.write_text(rendered)
    return secrets_path
