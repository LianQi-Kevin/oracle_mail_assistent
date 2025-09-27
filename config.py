from dataclasses import dataclass
from typing import Optional


@dataclass
class Config:
    # Configuration for Aconex API
    lobby_url: str = "https://constructionandengineering.oraclecloud.com"
    resource_url: str = "https://api.aconex.com"
    client_id: str = "CLIENT_ID"
    client_secret: str = "CLIENT_SECRET"
    access_token_expires: int = None
    access_token: str = None
    aconex_user_id: str = "ACONEX_USER_ID"
    aconex_instance_url: str = "https://asia1.aconex.com"
    project_id: str = "PROJECT_ID"

    # request settings
    proxies: Optional[dict[str, str]] = None  # Example: {"http": "http://127.0.0.1:8000", "https": "http://127.0.0.1:8000"}
    retry_times: int = 3
    retry_delay: int = 5  # seconds

    # fill colors
    finish_fill_color: str = "92D050"  # Green
    unSuccess_fill_color: str = "FFFF00"  # Yellow
    warning_fill_color: str = "00B0F0"  # Blue


config = Config()
