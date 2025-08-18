class QuickPasteError(Exception):
    """Base exception for QuickPaste."""


class ClipboardError(QuickPasteError):
    """Clipboard operation failed."""


class HotkeyError(QuickPasteError):
    """Hotkey registration/validation failed."""


class StorageError(QuickPasteError):
    """Storage operation failed."""

