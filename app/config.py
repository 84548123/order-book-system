import os
from dataclasses import dataclass


@dataclass(frozen=True)
class Settings:
    max_upload_mb: int = int(os.getenv("MAX_UPLOAD_MB", "15"))
    max_rows: int = int(os.getenv("MAX_ROWS", "100000"))
    file_ttl_minutes: int = int(os.getenv("FILE_TTL_MINUTES", "30"))


settings = Settings()

