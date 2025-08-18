import logging
import win32clipboard


def check_clipboard_access() -> bool:
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        logging.warning(f"Clipboard health check failed: {e}")
        return False


def check_hotkeys_registered(registered_ids: list[int]) -> bool:
    return len(registered_ids) > 0

