from dataclasses import dataclass

from . import clipboard as clipboard_mod
from . import hotkeys as hotkeys_mod
from . import storage as storage_mod
from .validators import HotkeyValidator


@dataclass
class ServiceContainer:
    clipboard: type = clipboard_mod
    hotkeys: type = hotkeys_mod
    storage: type = storage_mod
    hotkey_validator: type = HotkeyValidator

