"""
Configuration for the Word MCP SSE Server.
"""

from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings using Pydantic."""

    # MCP Server Auth
    MCP_SERVER_AUTH_KEY: Optional[str] = None

    # Azure App Registration (Word MCP)
    AZURE_TENANT_ID: Optional[str] = None
    AZURE_CLIENT_ID: Optional[str] = None
    AZURE_CLIENT_SECRET: Optional[str] = None

    # SharePoint Templates (Word MCP)
    SHAREPOINT_DRIVE_ID: Optional[str] = None          # ← changed
    SHAREPOINT_TEMPLATE_FOLDER: Optional[str] = None

    # OneDrive Output (Word MCP)
    ONEDRIVE_USER: Optional[str] = None
    ONEDRIVE_OUTPUT_FOLDER: Optional[str] = None

    # Optional settings with defaults
    LOG_LEVEL: str = "INFO"
    ENVIRONMENT: str = "development"
    FILE_LOGGING: bool = True
    LOGS_DIR: str = "logs"

    model_config = SettingsConfigDict(
        env_file=".env",
        case_sensitive=True,
        extra="allow",
    )


def load_config(dotenv_path: Optional[Path] = None) -> Settings:
    """
    Load configuration from environment variables and .env file.

    Args:
        dotenv_path: Path to .env file. Defaults to .env in project root.

    Returns:
        Settings object with loaded configuration.

    Raises:
        ValueError: If required configuration is missing.
    """
    # Load .env file
    if dotenv_path is None:
        dotenv_path = Path(__file__).parents[1] / ".env"

    load_dotenv(dotenv_path=dotenv_path)

    # Create settings
    settings = Settings()

    # Validate required keys
    required_keys = [
        "MCP_SERVER_AUTH_KEY",
        "AZURE_TENANT_ID",
        "AZURE_CLIENT_ID",
        "AZURE_CLIENT_SECRET",
        "SHAREPOINT_DRIVE_ID",              # ← changed
        "SHAREPOINT_TEMPLATE_FOLDER",
        "ONEDRIVE_USER",
        "ONEDRIVE_OUTPUT_FOLDER",
    ]
    missing_keys = []

    for key in required_keys:
        value = getattr(settings, key)
        if value is None or value == "":
            missing_keys.append(key)

    if missing_keys:
        raise ValueError(f"Missing required configuration: {', '.join(missing_keys)}")

    return settings