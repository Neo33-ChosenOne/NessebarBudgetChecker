"""Application configuration, loaded from environment variables / a local .env file.

Everything about this project is local-only: no cloud services, no telemetry,
no outbound calls other than the ones the user explicitly configures (the
Telegram bot, and whatever public municipal/government sites are scraped).
"""

from __future__ import annotations

from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Runtime settings for Nessebar Budget Monitor.

    Values are read from a local `.env` file (see `.env.example`) and/or
    real environment variables. Nothing here is sent anywhere by default.
    """

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
    )

    telegram_bot_token: str | None = None
    telegram_chat_id: str | None = None

    database_url: str = "sqlite:///data/nessebar.db"

    data_dir: Path = Path("data")

    #: Base URL the static site is published at; used to build flag deep
    #: links in Telegram notifications (`{site_base_url}/flags/#<flag-id>`)
    #: and by the site generator for absolute links/canonical URLs.
    site_base_url: str = "https://neo33-chosenone.github.io/NessebarBudgetChecker"

    #: Whether the public site and data exports show the names of natural
    #: persons (company managers/owners, officials, ЕТ sole traders). Default
    #: off: names are held back for legal review. Legal entities' company
    #: names and ЕИК are always shown; an ЕТ's name is a natural person's name.
    show_person_names: bool = False


def get_settings() -> Settings:
    """Return a freshly loaded Settings instance."""
    return Settings()
