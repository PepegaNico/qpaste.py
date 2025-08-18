import gettext
import os


def setup_localization(domain: str = 'quickpaste', localedir: str = 'locale'):
    try:
        gettext.install(domain, localedir=localedir)
    except Exception:
        # Fallback to built-in no-op if locales are not present
        gettext.install(domain)

