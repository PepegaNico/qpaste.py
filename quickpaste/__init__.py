"""QuickPaste modular package.

This package contains extracted modules from the monolithic QuickPaste.py
to improve structure and maintainability without changing runtime behavior.
"""

# Re-export commonly used elements for convenience (avoid importing PyQt at package import time)
from .clipboard import ClipboardManager, set_clipboard_html, html_to_rtf  # noqa: F401
from .hotkeys import release_all_modifier_keys  # noqa: F401

