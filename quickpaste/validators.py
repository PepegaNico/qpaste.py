class HotkeyValidator:
    """Centralized validation and normalization for hotkeys.

    Expected canonical format: 'ctrl+shift+[char]'
    Allowed chars: digits 0-9 and a subset of letters (b, e, f, h, m, p, q, v, x, z)
    """

    ALLOWED_CHARS = tuple("1234567890befhmpqvxz")

    @staticmethod
    def normalize(hotkey: str) -> str:
        if not hotkey:
            return ""
        hotkey = hotkey.strip().lower()
        # collapse multiple '+' and spaces
        parts = [p for p in hotkey.replace(' ', '').split('+') if p]
        return "+".join(parts)

    @classmethod
    def is_valid(cls, hotkey: str) -> bool:
        hotkey = cls.normalize(hotkey)
        parts = hotkey.split("+")
        if len(parts) != 3:
            return False
        if parts[0] != "ctrl" or parts[1] != "shift":
            return False
        return parts[2] in cls.ALLOWED_CHARS

    @classmethod
    def next_available(cls, used: set[str]) -> str:
        """Return the next available canonical hotkey not in 'used'."""
        for ch in cls.ALLOWED_CHARS:
            candidate = f"ctrl+shift+{ch}"
            if candidate not in used:
                return candidate
        # Fallback: create an index-based key (should not normally happen)
        idx = 1
        while True:
            candidate = f"ctrl+shift+{idx}"
            if candidate not in used:
                return candidate
            idx += 1

