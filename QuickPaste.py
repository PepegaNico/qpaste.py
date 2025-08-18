import sys, os, json, re, ctypes, logging
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QByteArray, QThread, pyqtSignal
import time
import pyperclip
from PyQt5.QtWidgets import QSystemTrayIcon, QAction, QMenu
import win32clipboard
import win32con
from functools import partial
import tempfile
import shutil
from quickpaste.clipboard import ClipboardManager, set_clipboard_html, html_to_rtf
from quickpaste.hotkeys import release_all_modifier_keys

APPDATA_PATH = os.path.join(os.environ["APPDATA"], "QuickPaste")
os.makedirs(APPDATA_PATH, exist_ok=True)
CONFIG_FILE = os.path.join(APPDATA_PATH, "config.json")
WINDOW_CONFIG = os.path.join(APPDATA_PATH, "window_config.json")
SDE_FILE = os.path.join(APPDATA_PATH, "sde.json")
LOG_FILE = os.path.join(APPDATA_PATH, "qp.log")
logging.basicConfig(filename=LOG_FILE, filemode="a", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8")
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
ICON_PATH = os.path.join(BASE_DIR, "assets", "H.ico")
class QuickPasteState:
    """Centralized application state management"""
    def __init__(self):
        self.dark_mode = False
        self.registered_hotkey_refs = []
        self.unsaved_changes = False
        self.dragged_index = None
        self.current_target = None
        self.profile_buttons = {}
        self.profile_lineedits = {}
        self.edit_mode = False
        self.text_entries = []
        self.title_entries = []
        self.hotkey_entries = []
        self.tray = None
        self.data = None
        self.active_profile = None
        self.profile_entries = {}
        self.last_ui_data = None
        self.registered_hotkey_ids = []
        self.id_to_index = {}
        self.hotkey_filter_instance = None