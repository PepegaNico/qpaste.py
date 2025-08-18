import logging
import time
import ctypes
import re

import win32clipboard
import win32con
from PyQt5 import QtGui


class ClipboardManager:
    def __init__(self):
        self.clipboard_opened = False

    def __enter__(self):
        for _ in range(5):
            try:
                win32clipboard.OpenClipboard()
                self.clipboard_opened = True
                break
            except Exception:
                time.sleep(0.01)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.clipboard_opened:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass

    def is_open(self):
        return self.clipboard_opened

    def empty(self):
        if self.clipboard_opened:
            win32clipboard.EmptyClipboard()

    def set_text(self, text, format_type):
        if self.clipboard_opened:
            win32clipboard.SetClipboardText(text, format_type)

    def set_data(self, format_type, data):
        if self.clipboard_opened:
            win32clipboard.SetClipboardData(format_type, data)


def set_clipboard_html(html_content, plain_text_content):
    max_retries = 3
    retry_delay = 0.02

    for attempt in range(max_retries):
        try:
            with ClipboardManager() as clipboard:
                if not clipboard.is_open():
                    logging.warning("Failed to open clipboard after 5 attempts")
                    return False

                clipboard.empty()
                clipboard.set_text(plain_text_content, win32con.CF_TEXT)

                html_header = """Version:0.9
StartHTML:0000000105
EndHTML:{:010d}
StartFragment:0000000141
EndFragment:{:010d}
<html>
<body>
<!--StartFragment-->{}<!--EndFragment-->
</body>
</html>""".format(len(html_content) + 175, len(html_content) + 141, html_content)

                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                clipboard.set_data(cf_html, html_header.encode('utf-8'))

                time.sleep(0.01)
                with ClipboardManager() as verify_clipboard:
                    if verify_clipboard.is_open():
                        html_available = win32clipboard.IsClipboardFormatAvailable(cf_html)
                        text_available = win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT)
                        if html_available and text_available:
                            logging.info(f"Successfully set clipboard with HTML content (attempt {attempt + 1}): {html_content[:50]}...")
                            return True
                        else:
                            logging.warning(
                                f"Clipboard verification failed (attempt {attempt + 1}): HTML={html_available}, Text={text_available}"
                            )

        except Exception as e:
            logging.warning(f"Error setting clipboard (attempt {attempt + 1}): {e}")

        if attempt < max_retries - 1:
            time.sleep(retry_delay)

    logging.error(f"Failed to set clipboard after {max_retries} attempts")
    return False


def html_to_rtf(html_content, plain_text):
    try:
        rtf = "{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Times New Roman;}} "
        link_pattern = r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>([^<]+)</a>'

        def replace_link(match):
            url = match.group(1)
            text = match.group(2)
            return f"{{\\field{{\\*\\fldinst{{HYPERLINK \"{url}\"}}}}{{\\fldrslt{{\\ul\\cf1 {text}}}}}}}"

        converted_text = re.sub(link_pattern, replace_link, html_content, flags=re.IGNORECASE)
        converted_text = re.sub(r'<[^>]+>', '', converted_text)
        rtf += converted_text + "}"
        return rtf
    except Exception as e:
        logging.warning(f"RTF conversion failed: {e}")
        return f"{{\\rtf1\\ansi\\deff0 {{\\fonttbl {{\\f0 Times New Roman;}}}} {plain_text}}}"

