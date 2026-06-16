from zoneinfo import ZoneInfo

from pydantic import FilePath, SecretStr
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=".env", env_file_encoding="utf-8", extra="allow"
    )

    openai_key: SecretStr
    docx_instructions: FilePath

    process_filetypes: list[str]

    tenant_id: SecretStr

    group_id: int

    sharepoint_domain: str
    sharepoint_site_name: str
    sharepoint_site_root_dir: str = "/root/children"

    tz: ZoneInfo = "America/Los_Angeles"


settings = Settings()
