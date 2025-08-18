import ctypes
import logging
import time


def release_all_modifier_keys():
    """
    Lässt sicherheitshalber alle Modifiertasten los (Ctrl/Shift/Alt/Win),
    ohne die 'keyboard'-Bibliothek zu verwenden.
    """
    try:
        user32 = ctypes.windll.user32
        KEYEVENTF_KEYUP = 0x0002

        # VK-Codes: Ctrl, Shift, Alt, LWin, RWin
        modifiers = (0x11, 0x10, 0x12, 0x5B, 0x5C)

        for _ in range(3):
            for vk in modifiers:
                try:
                    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
                except Exception:
                    pass
            time.sleep(0.01)

        time.sleep(0.05)
    except Exception as e:
        logging.warning(f"Error releasing modifier keys (WinAPI): {e}")

