from dataclasses import dataclass
from pathlib import Path
import os


@dataclass
class AppConfig:
    app_name: str = "QuickPaste"
    max_profiles: int = 10
    auto_save_delay_ms: int = 1000
    log_level: str = "INFO"

    @property
    def app_data_path(self) -> Path:
        # On Windows use APPDATA; otherwise fallback to home/.config
        appdata = os.environ.get("APPDATA")
        if appdata:
            return Path(appdata) / self.app_name
        return Path.home() / ".config" / self.app_name

