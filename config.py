from dataclasses import dataclass


@dataclass
class Config:
    lobby_url: str = "https://constructionandengineering.oraclecloud.com"
    resource_url: str = "https://api.aconex.com"
    client_id: str = "CLIENT_ID"
    client_secret: str = "CLIENT_SECRET"
    access_token_expires: int = 3600
    access_token: str = "ACCESS_TOKEN"
    aconex_user_id: str = "ACONEX_USER_ID"
    aconex_instance_url: str = "https://asia1.aconex.com"
    project_id: str = "PROJECT_ID"


config = Config()
