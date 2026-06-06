from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
IMAGE_DIR = DATA_DIR / "images"
OUTPUT_DIR = DATA_DIR / "output"
AI_INPUT_IMAGE_DIR = BASE_DIR / "input_images"
AI_OUTPUT_IMAGE_DIR = BASE_DIR / "output_images"
AI_OUTPUT_EXCEL_DIR = BASE_DIR / "output_excel"
LOG_DIR = BASE_DIR / "logs"


@dataclass
class Settings:
    client_id: str = ""
    api_key: str = ""
    base_url: str = "https://api-seller.ozon.ru"
    mock_mode: bool = True


def load_settings() -> Settings:
    load_dotenv(BASE_DIR / ".env")
    return Settings(
        client_id=os.getenv("OZON_CLIENT_ID", ""),
        api_key=os.getenv("OZON_API_KEY", ""),
        base_url=os.getenv("OZON_BASE_URL", "https://api-seller.ozon.ru"),
        mock_mode=os.getenv("OZON_MOCK_MODE", "true").lower() in {"1", "true", "yes", "y"},
    )


def mask_secret(value: str) -> str:
    if not value:
        return ""
    if len(value) <= 8:
        return "***"
    return f"{value[:4]}***{value[-4:]}"


def ensure_directories() -> None:
    for path in [INPUT_DIR, IMAGE_DIR, OUTPUT_DIR, AI_INPUT_IMAGE_DIR, AI_OUTPUT_IMAGE_DIR, AI_OUTPUT_EXCEL_DIR, LOG_DIR]:
        path.mkdir(parents=True, exist_ok=True)
