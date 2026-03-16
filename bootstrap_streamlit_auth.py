#!/usr/bin/env python3
"""Create Streamlit OIDC secrets from environment variables before startup."""

from __future__ import annotations

from pathlib import Path

from auth_support import ensure_streamlit_auth_secrets


def main() -> None:
    ensure_streamlit_auth_secrets(Path(__file__).resolve().parent)


if __name__ == "__main__":
    main()
