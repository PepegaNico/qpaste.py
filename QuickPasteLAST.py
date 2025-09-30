import sys, os, json, re, ctypes, logging, copy
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QFontMetrics
from PyQt5.QtCore import QByteArray
import time
import pyperclip
from PyQt5.QtWidgets import QSystemTrayIcon, QAction, QMenu
import win32clipboard
import win32con
from functools import partial
import tempfile
import shutil
import sip

APPDATA_PATH = os.path.join(os.environ["APPDATA"], "QuickPaste")
os.makedirs(APPDATA_PATH, exist_ok=True)
CONFIG_FILE = os.path.join(APPDATA_PATH, "config.json")
WINDOW_CONFIG = os.path.join(APPDATA_PATH, "window_config.json")
SDE_FILE = os.path.join(APPDATA_PATH, "sde.json")
LOG_FILE = os.path.join(APPDATA_PATH, "qp.log")
logging.basicConfig(filename=LOG_FILE, filemode="a", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8")
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
ICON_PATH = os.path.join(BASE_DIR, "assets", "H.ico")

DEFAULT_FONT_SIZE = 5 
MIN_FONT_SIZE = 5
MAX_FONT_SIZE = 5

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
        self.profile_selector = None
        self.profile_delete_button = None
        self.last_ui_data = None
        self.registered_hotkey_ids = []
        self.id_to_index = {}
        self.hotkey_filter_instance = None
        self.mini_mode = False
        self.saved_geometry = None
        self.normal_minimum_width = None
        self.zoom_level = 1.0  
        self.base_font_size = 5  

app_state = QuickPasteState()

class ComboBoxItemProxy:
    """Wrap QComboBox items to mimic QLineEdit behaviour."""
    def __init__(self, combo_box, index):
        self._combo_box = combo_box
        self._index = index
        self._pending_text = None
    def text(self):
        if self._pending_text is not None:
            return self._pending_text
        if self._combo_box.isEditable() and self._combo_box.currentIndex() == self._index:
            line_edit = self._combo_box.lineEdit()
            if line_edit is not None:
                return line_edit.text()
        return self._combo_box.itemText(self._index)
    def setText(self, value):
        self._pending_text = None
        if self._combo_box.currentIndex() == self._index and self._combo_box.isEditable():
            line_edit = self._combo_box.lineEdit()
            if line_edit is not None:
                line_edit.setText(value)
        self._combo_box.setItemText(self._index, value)
    def set_pending_text(self, value):
        normalized_new = (value or "").strip()
        original = (self._combo_box.itemData(self._index) or "").strip()
        if normalized_new == original:
            self._pending_text = None
        else:
            self._pending_text = value
    def clear_pending_text(self):
        self._pending_text = None

class ProfileComboBox(QtWidgets.QComboBox):
    """ComboBox mit intern gerendertem Pfeil-Glyph."""
    GLYPH = "▼" 
    def paintEvent(self, event):
        super().paintEvent(event)
        option = QtWidgets.QStyleOptionComboBox()
        self.initStyleOption(option)
        style = self.style()
        arrow_rect = style.subControlRect(
            QtWidgets.QStyle.CC_ComboBox,
            option,
            QtWidgets.QStyle.SC_ComboBoxArrow,
            self,)
        if not arrow_rect.isValid() or arrow_rect.width() <= 0 or arrow_rect.height() <= 0:
            drop_w = self.style().pixelMetric(QtWidgets.QStyle.PM_ComboBoxButtonWidth, option, self)
            if drop_w <= 0:
                drop_w = int(self.height() * 0.8)
            r = self.rect()
            arrow_rect = QtCore.QRect(r.right() - drop_w, r.top(), drop_w, r.height())
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.TextAntialiasing, True)
        base_palette = self.palette()
        if self.isEnabled():
            color = base_palette.color(QtGui.QPalette.ButtonText)
        else:
            color = base_palette.color(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText)
        painter.setPen(QtGui.QPen(color))
        font = painter.font()
        target_size = int(min(arrow_rect.width(), arrow_rect.height()) * 0.65)
        if target_size > 0:
            font.setPixelSize(target_size)
        painter.setFont(font)
        painter.drawText(arrow_rect, QtCore.Qt.AlignCenter, self.GLYPH)
        painter.end()

class DebouncedSaver:
    def __init__(self, delay_ms=600):
        self.timer = QtCore.QTimer()
        self.timer.setSingleShot(True)
        self.timer.setInterval(delay_ms)
        self.timer.timeout.connect(self._save)
        self.pending_data = None
    def schedule_save(self, data):
        self.pending_data = data
        self.timer.start()
    def _save(self):
        if self.pending_data is not None:
            try:
                save_data_atomic(self.pending_data, CONFIG_FILE)
            finally:
                self.pending_data = None

def save_data_atomic(data, filename):
    """Atomic file write to prevent corruption"""
    dirpath = os.path.dirname(filename)
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(dir=dirpath, prefix=".tmp_", suffix=".json")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        os.replace(tmp, filename) 
    except Exception as e:
        if tmp:
            try: os.unlink(tmp)
            except Exception: pass
        raise e
debounced_saver = DebouncedSaver(600)

#region window position 

def save_window_position():
    """Speichert Fensterposition und weitere UI-Einstellungen"""
    try:
        geo_bytes = win.saveGeometry()
        geo_hex = bytes(geo_bytes.toHex()).decode()
        cfg = {
            "geometry_hex": geo_hex,
            "dark_mode": app_state.dark_mode,
            "mini_mode": app_state.mini_mode}
        if app_state.saved_geometry is not None:
            cfg["normal_geometry_hex"] = bytes(app_state.saved_geometry.toHex()).decode()
        else:
            cfg["normal_geometry_hex"] = geo_hex
        tmp = WINDOW_CONFIG + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
        os.replace(tmp, WINDOW_CONFIG)
    except Exception as e:
        logging.exception(f"⚠ Fehler beim Speichern der Fensterposition: {e}")

def load_window_position():
    """Lädt Fensterposition und andere UI-Einstellungen"""
    try:
        with open(WINDOW_CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("dark_mode") is not None:
            app_state.dark_mode = cfg["dark_mode"]
        if cfg.get("mini_mode") is not None:
            app_state.mini_mode = cfg["mini_mode"]
        app_state.zoom_level = detect_optimal_zoom()
        hexstr = cfg.get("geometry_hex")
        if hexstr:
            ba = QByteArray.fromHex(hexstr.encode())
            win.restoreGeometry(ba)
        normal_hex = cfg.get("normal_geometry_hex")
        if normal_hex:
            app_state.saved_geometry = QByteArray.fromHex(normal_hex.encode())
        elif hexstr:
            app_state.saved_geometry = QByteArray.fromHex(hexstr.encode())
        else:
            app_state.saved_geometry = None
        return True
    except (FileNotFoundError, json.JSONDecodeError):
        app_state.zoom_level = detect_optimal_zoom()
        return False

def detect_optimal_zoom():
    """Erkennt optimalen Zoom basierend auf System-DPI"""
    try:
        screen = QtWidgets.QApplication.primaryScreen()
        if screen:
            logical_dpi = screen.logicalDotsPerInch()
            if logical_dpi <= 96:
                return 1.0
            elif logical_dpi <= 120:
                return 1.1
            elif logical_dpi <= 144:
                return 1.2
            else:
                return 1.3
    except Exception:
        pass
    return 1.0

#endregion

#region data

def load_sde_profile():
    try:
        with open(SDE_FILE, "r", encoding="utf-8") as f:
            sde = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        logging.warning("⚠ Konnte sde.json nicht laden. Setze Standard‑SDE.")
        sde = {}
    if not sde.get("titles") and not sde.get("texts") and not sde.get("hotkeys"):
        sde = {
            "titles": ["Standard Titel 1", "Standard Titel 2", "Standard Titel 3"],
            "texts":  ["Standard Text 1",  "Standard Text 2",  "Standard Text 3"],
            "hotkeys":["ctrl+shift+1",    "ctrl+shift+2",    "ctrl+shift+3"]}
    return sde

def load_data():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        if not isinstance(loaded.get("profiles"), dict):
            loaded["profiles"] = {}
        for prof, vals in loaded["profiles"].items():
            vals.setdefault("titles", [])
            vals.setdefault("texts",  [])
            vals.setdefault("hotkeys", [])
        loaded["profiles"]["SDE"] = load_sde_profile()
        ap = loaded.get("active_profile")
        if ap not in loaded["profiles"]:
            if loaded["profiles"]:
                loaded["active_profile"] = next(iter(loaded["profiles"]))
            else:
                logging.warning("⚠ Keine Profile gefunden. Erstelle Standardprofil.")
                loaded["profiles"] = {
                    "Profil 1": {"titles": [], "texts": [], "hotkeys": []}}
                loaded["active_profile"] = "Profil 1"
        return loaded
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "profiles": {
                "Profil 1": {
                    "titles":  [f"Titel {i}"        for i in range(1,6)],
                    "texts":   [f"Text {i}"         for i in range(1,6)],
                    "hotkeys": [f"ctrl+shift+{i}"   for i in range(1,6)]},
                "Profil 2": {
                    "titles":  [f"Titel {i}"        for i in range(1,6)],
                    "texts":   [f"Profil 2 Text {i}"for i in range(1,6)],
                    "hotkeys": [f"ctrl+shift+{i}"   for i in range(1,6)]},
                "SDE": load_sde_profile()},
            "active_profile": "Profil 1"}
app_state.data = load_data()
app_state.active_profile = app_state.data.get("active_profile", list(app_state.data["profiles"].keys())[0])

#endregion 

#region profiles

def _normalize_rich_text(value):
    """Normalize rich-text HTML for reliable comparisons."""
    doc = QtGui.QTextDocument()
    doc.setHtml(value or "")
    return doc.toHtml()
def _normalize_title(value):
    return (value or "").strip()
def _normalize_hotkey(value):
    return (value or "").strip().lower()
def has_field_changes(profile_to_check=None):
    if profile_to_check is None:
        profile_to_check = app_state.active_profile
    profiles = app_state.data.get("profiles", {})
    if profile_to_check not in profiles:
        return False

    rename_changed = False
    if app_state.profile_entries:
        for old_name, entry in app_state.profile_entries.items():
            if old_name == "SDE":
                continue
            try:
                new_name = _normalize_title(entry.text())
            except Exception:
                new_name = _normalize_title(getattr(entry, "text", lambda: "")())
            if new_name != _normalize_title(old_name):
                rename_changed = True
                break

    titles, texts, hks = [], [], []
    for i in range(entries_layout.count()):
        item = entries_layout.itemAt(i)
        if item is None:
            continue
        row = item.widget()
        if app_state.edit_mode and isinstance(row, QtWidgets.QWidget):
            line_edits = row.findChildren(QtWidgets.QLineEdit)
            text_edits = row.findChildren(QtWidgets.QTextEdit)
            if len(line_edits) >= 2 and len(text_edits) >= 1:
                titles.append(line_edits[0].text())
                texts.append(text_edits[0].toHtml())
                hks.append(line_edits[1].text())

    normalized_titles = [_normalize_title(t) for t in titles]
    normalized_texts = [_normalize_rich_text(t) for t in texts]
    normalized_hotkeys = [_normalize_hotkey(h) for h in hks]

    reference_profiles = app_state.last_ui_data if isinstance(app_state.last_ui_data, dict) else None
    reference_profile = None
    if reference_profiles is not None:
        reference_profile = reference_profiles.get(profile_to_check)
    if reference_profile is None:
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                persisted = json.load(f)
            reference_profile = persisted.get("profiles", {}).get(profile_to_check)
        except (FileNotFoundError, json.JSONDecodeError):
            reference_profile = None
    if reference_profile is None:
        reference_profile = {}

    stored_titles = [_normalize_title(t) for t in reference_profile.get("titles", [])]
    stored_texts = [_normalize_rich_text(t) for t in reference_profile.get("texts", [])]
    stored_hotkeys = [_normalize_hotkey(h) for h in reference_profile.get("hotkeys", [])]

    fields_changed = (
        normalized_titles != stored_titles
        or normalized_texts != stored_texts
        or normalized_hotkeys != stored_hotkeys)

    changed = rename_changed or fields_changed
    if not changed:
        app_state.unsaved_changes = False
        return False
    app_state.unsaved_changes = True
    return True

def apply_profile_renames(show_errors=True):
    """Überträgt Umbenennungen aus der UI in die Datenstruktur."""
    if not app_state.edit_mode or not app_state.profile_entries:
        return False
    proposed = {}
    proposed_lower = set()
    for old_name, entry in app_state.profile_entries.items():
        if old_name == "SDE":
            continue
        new_name = (entry.text() or "").strip()
        if not new_name:
            if show_errors:
                show_critical_message("Fehler", f"Profilname für '{old_name}' darf nicht leer sein.")
            return None
        if new_name == "SDE":
            if show_errors:
                show_critical_message("Fehler", "Der Profilname 'SDE' ist reserviert.")
            return None
        lower_name = new_name.lower()
        if lower_name in proposed_lower:
            if show_errors:
                show_critical_message("Fehler", f"Profilname '{new_name}' ist doppelt.")
            return None
        existing_lower = {
            n.lower()
            for n in app_state.data["profiles"].keys()
            if n not in (old_name, "SDE")}
        if lower_name in existing_lower:
            if show_errors:
                show_critical_message("Fehler", f"Profilname '{new_name}' existiert bereits.")
            return None
        if new_name != old_name:
            proposed[old_name] = new_name
            proposed_lower.add(lower_name)
    if not proposed:
        return False
    new_profiles = {}
    for old_name, prof_data in app_state.data["profiles"].items():
        if old_name == "SDE":
            continue
        target = proposed.get(old_name, old_name)
        new_profiles[target] = prof_data
    if len(new_profiles) > 11:
        if show_errors:
            show_critical_message("Limit erreicht", "Maximal 10 Profile erlaubt!")
        return None
    sde_profile = None
    if "SDE" in app_state.data["profiles"]:
        sde_profile = load_sde_profile()
        new_profiles["SDE"] = sde_profile
    if app_state.active_profile in proposed:
        app_state.active_profile = proposed[app_state.active_profile]
    if sde_profile is not None and app_state.active_profile not in new_profiles:
        app_state.active_profile = next((k for k in new_profiles.keys() if k != "SDE"), "SDE")
    app_state.data["profiles"] = new_profiles
    app_state.data["active_profile"] = app_state.active_profile
    return True
def _remember_profile_name_edit(text):
    if not app_state.edit_mode:
        return
    combo = getattr(app_state, "profile_selector", None)
    if combo is None or not combo.isEditable():
        return
    index = combo.currentIndex()
    if index < 0:
        return
    original = combo.itemData(index)
    if not original or original == "SDE":
        return
    entry = app_state.profile_entries.get(original)
    if entry is None:
        return
    entry.set_pending_text(text)
def save_profile_names():
    result = apply_profile_renames(show_errors=True)
    if result is None or result is False:
        return
    profiles_to_save = {k: v for k, v in app_state.data["profiles"].items() if k != "SDE"}
    debounced_saver.schedule_save({
        "profiles": profiles_to_save,
        "active_profile": app_state.active_profile})
    update_ui()
    update_profile_buttons()
    refresh_tray()
def update_profile_buttons():
    combo = getattr(app_state, "profile_selector", None)
    if combo is None:
        return
    target_index = -1
    for idx in range(combo.count()):
        if combo.itemData(idx) == app_state.active_profile:
            target_index = idx
            break
    if target_index >= 0 and combo.currentIndex() != target_index:
        with QtCore.QSignalBlocker(combo):
            combo.setCurrentIndex(target_index)
    delete_btn = getattr(app_state, "profile_delete_button", None)
    if delete_btn is not None:
        index = combo.currentIndex()
        data = combo.itemData(index) if index >= 0 else None
        can_delete = (
            index >= 0
            and combo.count() > 1
            and data not in (None, "SDE"))
        delete_btn.setEnabled(can_delete)


def switch_profile(profile_name):
    if profile_name == app_state.active_profile:
        return
    was_visible = win.isVisible()
    def restore_active_in_selector():
        combo = getattr(app_state, "profile_selector", None)
        if combo is None:
            return
        target_index = next(
            (i for i in range(combo.count()) if combo.itemData(i) == app_state.active_profile),
            -1,)
        if target_index < 0:
            target_index = next(
            (i for i in range(combo.count()) if combo.itemText(i) == app_state.active_profile),
            -1,)
        if target_index >= 0:
            with QtCore.QSignalBlocker(combo):
                combo.setCurrentIndex(target_index)
    if app_state.edit_mode and has_field_changes():
        resp = show_question_message(
            "Ungesicherte Änderungen",
            "Du hast ungespeicherte Änderungen. Jetzt speichern?",
            QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel,)
        if resp == QtWidgets.QMessageBox.Cancel:
            restore_active_in_selector()
            return
        if resp == QtWidgets.QMessageBox.Yes:
            save_data(stay_in_edit_mode=True)
        else:
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    file_data = json.load(f)
                if app_state.active_profile in file_data.get("profiles", {}):
                    app_state.data["profiles"][app_state.active_profile] = file_data["profiles"][app_state.active_profile].copy()
            except (FileNotFoundError, json.JSONDecodeError, KeyError):
                pass
    if profile_name not in app_state.data["profiles"]:
        show_critical_message("Fehler", f"Profil '{profile_name}' existiert nicht!")
        return
    app_state.active_profile = profile_name
    app_state.data["active_profile"] = profile_name
    profiles_to_save = {k: v for k, v in app_state.data["profiles"].items() if k != "SDE"}
    debounced_saver.schedule_save({"profiles": profiles_to_save, "active_profile": profile_name})
    update_profile_buttons()
    if was_visible:
        update_ui()
    register_hotkeys()
    refresh_tray()
    if not was_visible:
        win.hide()

def add_new_profile():
    non_sde_count = sum(1 for p in app_state.data["profiles"].keys() if p != "SDE")
    if non_sde_count >= 10:
        show_critical_message("Limit erreicht", "Maximal 10 eigene Profile (ohne SDE) erlaubt!")
        return
    base, cnt = "Profil", 1
    while f"{base} {cnt}" in app_state.data["profiles"]:
        cnt += 1
    name = f"{base} {cnt}"
    app_state.data["profiles"][name] = {
        "titles":  [f"Titel {i+1}" for i in range(3)],
        "texts":   [f"Text {i+1}"  for i in range(3)],
        "hotkeys": [f"ctrl+shift+{i+1}" for i in range(3)]}
    app_state.active_profile = name
    app_state.data["active_profile"] = name
    update_ui()

def delete_profile(profile_name):
    if profile_name == "SDE":
        show_critical_message("Fehler", "Das SDE‑Profil kann nicht gelöscht werden.")
        return
    if len(app_state.data["profiles"]) <= 1:
        show_critical_message("Fehler", "Mindestens ein Profil muss bestehen bleiben!")
        return
    resp = show_question_message(
        "Profil löschen",
        f"Soll Profil '{profile_name}' wirklich gelöscht werden?")
    if resp != QtWidgets.QMessageBox.Yes:
        return
    del app_state.data["profiles"][profile_name]
    if app_state.active_profile == profile_name:
        app_state.active_profile = next(iter(app_state.data["profiles"]))
        app_state.data["active_profile"] = app_state.active_profile
    update_ui()
    save_data()

#endregion 

#region insert text / hotkeys

class ClipboardManager:
    def __init__(self):
        self.clipboard_opened = False
    def __enter__(self):
        for open_attempt in range(5):
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
    """
    Legt HTML + Plaintext korrekt in die Windows-Zwischenablage:
    - Plaintext als CF_UNICODETEXT (Umlaute/Emoji sicher)
    - HTML als CF_HTML mit korrekten Byte-Offsets (CRLF, UTF-8)
    """
    max_retries = 3
    retry_delay = 0.02
    for attempt in range(max_retries):
        try:
            with ClipboardManager() as clipboard:
                if not clipboard.is_open():
                    logging.warning("Failed to open clipboard after 5 attempts")
                    return False
                clipboard.empty()
                clipboard.set_text(plain_text_content or "", win32con.CF_UNICODETEXT)
                fragment = html_content or ""
                html_body = (
                    "<!DOCTYPE html><html><body>"
                    "<!--StartFragment-->"
                    + fragment +
                    "<!--EndFragment-->"
                    "</body></html>")
                header_tmpl = (
                    "Version:0.9\r\n"
                    "StartHTML:{start_html:010d}\r\n"
                    "EndHTML:{end_html:010d}\r\n"
                    "StartFragment:{start_frag:010d}\r\n"
                    "EndFragment:{end_frag:010d}\r\n")
                placeholder = header_tmpl.format(
                    start_html=0, end_html=0, start_frag=0, end_frag=0
                ).encode("utf-8")
                body_bytes = html_body.encode("utf-8")
                start_html = len(placeholder)
                end_html   = start_html + len(body_bytes)
                sf_in_body = body_bytes.find(b"<!--StartFragment-->") + len(b"<!--StartFragment-->")
                ef_in_body = body_bytes.find(b"<!--EndFragment-->")
                start_frag = start_html + sf_in_body
                end_frag   = start_html + ef_in_body
                header_bytes = header_tmpl.format(
                    start_html=start_html,
                    end_html=end_html,
                    start_frag=start_frag,
                    end_frag=end_frag
                ).encode("utf-8")
                full_bytes = header_bytes + body_bytes
                cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
                clipboard.set_data(cf_html, full_bytes)
                time.sleep(0.01)
                with ClipboardManager() as verify_clipboard:
                    if verify_clipboard.is_open():
                        html_ok = win32clipboard.IsClipboardFormatAvailable(cf_html)
                        txt_ok  = win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT)
                        if html_ok and txt_ok:
                            logging.info(f"Clipboard set (attempt {attempt+1}).")
                            return True
                        else:
                            logging.warning(
                                f"Clipboard verify failed (attempt {attempt+1}): HTML={html_ok}, TEXT={txt_ok}")
        except Exception as e:
            logging.warning(f"Error setting clipboard (attempt {attempt+1}): {e}")
        if attempt < max_retries - 1:
            time.sleep(retry_delay)
    logging.error(f"Failed to set clipboard after {max_retries} attempts")
    return False

def release_all_modifier_keys():
    """
    Lässt sicherheitshalber alle Modifiertasten los (Ctrl/Shift/Alt/Win),
    ohne die 'keyboard'-Bibliothek zu verwenden.
    """
    try:
        user32 = ctypes.windll.user32
        KEYEVENTF_KEYUP = 0x0002
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

def insert_text(index):
    def _send_ctrl_v():
        user32 = ctypes.windll.user32
        user32.keybd_event(0x11, 0, 0, 0) 
        user32.keybd_event(0x56, 0, 0, 0)  
        user32.keybd_event(0x56, 0, 2, 0)  
        user32.keybd_event(0x11, 0, 2, 0)  
    try:
        txt = app_state.data["profiles"][app_state.active_profile]["texts"][index]
        logging.info(f"Inserting text for index {index}: {txt[:50]}...")
    except IndexError:
        logging.exception(
            f"Kein Text vorhanden für Hotkey-Index {index} im Profil '{app_state.active_profile}'")
        return
    try:
        release_all_modifier_keys()
        doc = QtGui.QTextDocument()
        doc.setHtml(txt)
        plain_text = doc.toPlainText()
        success = set_clipboard_html(txt, plain_text)
        if not success:
            logging.warning("Windows clipboard failed, falling back to pyperclip")
            pyperclip.copy(plain_text)
            logging.info(f"Fallback: Set plain text to clipboard: {plain_text[:30]}...")
        time.sleep(0.2)
        release_all_modifier_keys()
        _send_ctrl_v()
        time.sleep(0.05)
        release_all_modifier_keys()
        logging.info(f"Successfully inserted text for index {index}")
    except Exception as e:
        logging.exception(f"Error in insert_text for index {index}: {e}")
        try:
            doc = QtGui.QTextDocument()
            doc.setHtml(txt)
            plain_text = doc.toPlainText()
            pyperclip.copy(plain_text)
            logging.info(f"Final fallback: Set plain text to clipboard: {plain_text[:30]}...")
            release_all_modifier_keys()
            _send_ctrl_v()
            time.sleep(0.05)
            release_all_modifier_keys()
            logging.info(f"Successfully inserted text (final fallback) for index {index}")
        except Exception as fallback_error:
            logging.exception(f"All clipboard methods failed for index {index}: {fallback_error}")

def copy_text_to_clipboard(index):
    try:
        txt = app_state.data["profiles"][app_state.active_profile]["texts"][index]
    except IndexError:
        logging.exception(
            f"Kein Text vorhanden für Index {index} im Profil '{app_state.active_profile}'")
        return
    try:
        doc = QtGui.QTextDocument()
        doc.setHtml(txt)
        plain_text = doc.toPlainText()
        success = set_clipboard_html(txt, plain_text)
        if not success:
            logging.warning("Windows clipboard failed, falling back to pyperclip")
            pyperclip.copy(plain_text)
            logging.info(f"Fallback: Copied plain text to clipboard: {plain_text[:30]}...")
    except Exception as e:
        logging.exception(f"Error copying text for index {index}: {e}")
        try:
            doc = QtGui.QTextDocument()
            doc.setHtml(txt)
            plain_text = doc.toPlainText()
            pyperclip.copy(plain_text)
            logging.info(f"Final fallback: Copied plain text to clipboard: {plain_text[:30]}...")
        except Exception as fallback_error:
            logging.exception(f"All clipboard copy methods failed for index {index}: {fallback_error}")
    if hasattr(win, 'statusBar'):
        win.statusBar().showMessage("Text in Zwischenablage kopiert!", 2000)

def cleanup_hotkeys():
    """Properly cleanup all registered hotkeys and event filters"""
    for hotkey_id in app_state.registered_hotkey_ids:
        try:
            ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
        except Exception as e:
            logging.warning(f"Failed to unregister hotkey {hotkey_id}: {e}")
    app_state.registered_hotkey_ids.clear()
    app_state.id_to_index.clear()
    if app_state.hotkey_filter_instance is not None:
        try:
            app.removeNativeEventFilter(app_state.hotkey_filter_instance)
            app_state.hotkey_filter_instance = None
        except Exception as e:
            logging.warning(f"Failed to remove native event filter: {e}")

def register_hotkeys():
    """
    Registriert globale Hotkeys via WinAPI (RegisterHotKey) und verarbeitet sie
    über ein Qt-NativeEventFilter – ganz ohne 'keyboard'-Bibliothek.
    Erwartetes Format der Hotkeys (wie bisher im UI): 'ctrl+shift+[zeichen]'
    """
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    WM_HOTKEY = 0x0312
    MOD_ALT     = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT   = 0x0004
    MOD_WIN     = 0x0008
    def _vk_from_char(ch: str):
        """
        Liefert den Virtual-Key-Code (VK) für ein einzelnes Zeichen.
        - Ziffern/Buchstaben: schnelle Pfade
        - Sonderzeichen (§ ' ^): via VkKeyScanExW anhand des *aktuellen* Keyboard-Layouts
        """
        ch = (ch or "").strip()
        if not ch:
            return None
        if "0" <= ch <= "9":
            return ord(ch) 
        if "a" <= ch <= "z":
            return ord(ch.upper())
        try:
            user32 = ctypes.windll.user32
            user32.VkKeyScanExW.restype = ctypes.c_short
            hkl = user32.GetKeyboardLayout(0)
            res = user32.VkKeyScanExW(ch, hkl) 
            if res == -1:
                return None
            vk = res & 0xFF
            return vk or None
        except Exception:
            return None
    cleanup_hotkeys()
    erlaubte_zeichen = set("1234567890befhmpqvxz§'^")
    belegte = set()
    fehler = False
    hotkeys = app_state.data["profiles"].setdefault(app_state.active_profile, {}).setdefault("hotkeys", [])
    next_id = 1
    for i, hot in enumerate(hotkeys):
        hot = (hot or "").strip().lower()
        if not hot:
            continue
        parts = hot.split("+")
        if (
            len(parts) != 3
            or parts[0] != "ctrl"
            or parts[1] != "shift"
            or parts[2] not in erlaubte_zeichen):
            show_critical_message(
                "Fehler",
                f"Ungültiger Hotkey \"{hotkeys[i]}\" für Eintrag {i+1}.\n"
                f"Erlaubte Zeichen: {''.join(sorted(erlaubte_zeichen))}\n"
                f"Format: ctrl+shift+[zeichen]")
            fehler = True
            continue
        if hot in belegte:
            show_critical_message("Fehler", f"Hotkey \"{hotkeys[i]}\" wird bereits verwendet!")
            fehler = True
            continue
        belegte.add(hot)
        if i >= len(app_state.data["profiles"][app_state.active_profile]["texts"]):
            logging.warning(f"⚠ Hotkey '{hot}' zeigt auf Eintrag {i+1}, aber dieser existiert nicht.")
            continue
        ch = parts[2]
        vk = _vk_from_char(ch)
        if vk is None:
            show_critical_message("Fehler", f"Hotkey-Zeichen '{ch}' wird nicht unterstützt.")
            fehler = True
            continue
        mods = MOD_CONTROL | MOD_SHIFT
        if not user32.RegisterHotKey(None, next_id, mods, vk):
            logging.error(f"RegisterHotKey fehlgeschlagen für {hot} (id={next_id})")
            fehler = True
            continue
        app_state.id_to_index[next_id] = i
        app_state.registered_hotkey_ids.append(next_id)
        next_id += 1
    logging.info(f"Registered {len(app_state.registered_hotkey_ids)} hotkeys for profile '{app_state.active_profile}'")
    if app_state.hotkey_filter_instance is None:
        class _MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd",    wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam",  wintypes.WPARAM),
                ("lParam",  wintypes.LPARAM),
                ("time",    wintypes.DWORD),
                ("pt",      wintypes.POINT),]
        class _HotkeyFilter(QtCore.QAbstractNativeEventFilter):
            def nativeEventFilter(self, eventType, message):
                if eventType == "windows_generic_MSG":
                    addr = int(message)               
                    msg  = _MSG.from_address(addr)  
                    if msg.message == WM_HOTKEY:
                        try:
                            hotkey_id = int(msg.wParam)
                            idx = app_state.id_to_index.get(hotkey_id)
                            if idx is not None:
                                insert_text(idx)
                        except Exception as e:
                            logging.exception(f"Fehler im WM_HOTKEY-Handler: {e}")
                return False, 0
        app_state.hotkey_filter_instance = _HotkeyFilter()
        app.installNativeEventFilter(app_state.hotkey_filter_instance)
        app.aboutToQuit.connect(cleanup_hotkeys)
    return fehler

#endregion

#region Tray

def create_tray_icon():
    try:
        if app_state.tray:
            try:
                app_state.tray.hide()
                app_state.tray.setParent(None)
                app_state.tray.deleteLater()
            except:
                pass
            app_state.tray = None
        icon = None
        if os.path.exists(ICON_PATH):
            try:
                icon = QtGui.QIcon(ICON_PATH)
                if icon.isNull():
                    icon = None
            except:
                logging.warning(f"Konnte Icon nicht laden: {ICON_PATH}")
                icon = None
        if not icon:
            icon = win.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon)
        app_state.tray = QSystemTrayIcon(icon, win)
        app_state.tray.setToolTip(f"Aktives Profil: {app_state.active_profile}")
        menu = QMenu()
        for prof in app_state.data["profiles"]:
            label = f"✓ {prof}" if prof == app_state.active_profile else f"  {prof}"
            act = QAction(label, win)
            act.triggered.connect(partial(switch_profile, prof))
            menu.addAction(act)
        menu.addSeparator()
        act_show = QAction("Öffnen", win)
        act_show.triggered.connect(lambda: (win.show(), win.raise_(), win.activateWindow()))
        menu.addAction(act_show)
        act_quit = QAction("Beenden", win)
        act_quit.triggered.connect(lambda: (save_window_position(), app.quit()))
        menu.addAction(act_quit)
        app_state.tray.setContextMenu(menu)
        app_state.tray.activated.connect(
            lambda reason: (win.show(), win.raise_(), win.activateWindow()) 
            if reason == QSystemTrayIcon.Trigger else None)
        app_state.tray.show()
        logging.info("Tray-Icon erfolgreich erstellt")
        return True
    except Exception as e:
        logging.error(f"Tray-Icon Erstellung fehlgeschlagen: {e}")
        app_state.tray = None
        return False

def refresh_tray():
    if app_state.tray is not None:
        try:
            app_state.tray.hide()
            app_state.tray.deleteLater()
        except Exception as e:
            logging.warning(f"Failed to cleanup tray: {e}")
        app_state.tray = None
    create_tray_icon()

def minimize_to_tray():
    win.hide()
    if app_state.tray and hasattr(app_state.tray, 'showMessage'):
        app_state.tray.showMessage(
            "QuickPaste", 
            "Anwendung wurde in die Taskleiste minimiert. Hotkeys bleiben aktiv.",
            QSystemTrayIcon.Information, 
            2000)

#endregion

#region add/del/move/drag Entry

def start_drag(event, index, widget):
    app_state.dragged_index = index
    row_widget = widget.parent()
    if hasattr(row_widget, 'highlight_drop_zone'):
        original_style = row_widget.styleSheet()
        row_widget.setStyleSheet(f"""
            QWidget {{
                background-color: {'#1a1a1a' if app_state.dark_mode else '#f0f0f0'};
                opacity: 0.6;
                border: 1px dashed {'#666' if app_state.dark_mode else '#999'};
                border-radius: 6px;}}""")
    drag = QtGui.QDrag(widget)
    mime_data = QtCore.QMimeData()
    mime_data.setText(str(index))
    drag.setMimeData(mime_data)
    result = drag.exec_(QtCore.Qt.MoveAction)
    if hasattr(row_widget, 'highlight_drop_zone'):
        row_widget.setStyleSheet(original_style)
        clear_all_highlights()

def clear_all_highlights():
    try:
        for i in range(entries_layout.count()):
            item = entries_layout.itemAt(i)
            if item and item.widget():
                widget = item.widget()
                if isinstance(widget, DragDropWidget) and hasattr(widget, 'highlight_drop_zone'):
                    widget.highlight_drop_zone(False)
    except Exception:
        pass
class DragDropWidget(QtWidgets.QWidget):
    def __init__(self, index, parent=None):
        super().__init__(parent)
        self.drag_index = index
        self.setAcceptDrops(True)
        self.original_style = ""
        self.is_highlighted = False
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            self.highlight_drop_zone(True)
            event.acceptProposedAction()
    
    def dragLeaveEvent(self, event):
        self.highlight_drop_zone(False)
        super().dragLeaveEvent(event)
    
    def dropEvent(self, event):
        self.highlight_drop_zone(False)
        if event.mimeData().hasText():
            source_index = int(event.mimeData().text())
            target_index = self.drag_index
            if source_index != target_index:
                move_entry_to(source_index, target_index)
            event.acceptProposedAction()
    
    def highlight_drop_zone(self, highlight):
        if highlight and not self.is_highlighted:
            self.original_style = self.styleSheet()
            highlight_color = "#4a90e2" if app_state.dark_mode else "#87ceeb"
            border_color = "#5aa3f0" if app_state.dark_mode else "#4682b4"
            self.setStyleSheet(f"""
                QWidget {{
                    background-color: {highlight_color};
                    border: 2px solid {border_color};
                    border-radius: 8px;}}""")
            self.is_highlighted = True
        elif not highlight and self.is_highlighted:
            self.setStyleSheet(self.original_style)
            self.is_highlighted = False

def add_new_entry():
    if "titles" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["titles"] = []
    if "texts" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["texts"] = []
    if "hotkeys" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["hotkeys"] = []
    erlaubte_zeichen = "1234567890befhmpqvxz§'^"
    belegte_hotkeys = set(app_state.data["profiles"][app_state.active_profile]["hotkeys"])
    neuer_hotkey = ""
    for zeichen in erlaubte_zeichen:
        test_hotkey = f"ctrl+shift+{zeichen}"
        if test_hotkey not in belegte_hotkeys:
            neuer_hotkey = test_hotkey
            break
    if not neuer_hotkey:
        neuer_hotkey = "ctrl+shift+"
    titles_list = app_state.data["profiles"][app_state.active_profile].setdefault("titles", [])
    existing_lower = {t.strip().lower() for t in titles_list}
    base = "Neuer Eintrag"
    candidate = base
    n = 2
    while candidate.strip().lower() in existing_lower:
        candidate = f"{base} {n}"
        n += 1
    app_state.data["profiles"][app_state.active_profile]["titles"].append(candidate)
    app_state.data["profiles"][app_state.active_profile]["texts"].append("Neuer Text")
    app_state.data["profiles"][app_state.active_profile]["hotkeys"].append(neuer_hotkey)
    update_ui()

def delete_entry(index):
    if "titles" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["titles"] = []
    if "texts" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["texts"] = []
    if "hotkeys" not in app_state.data["profiles"][app_state.active_profile]:
        app_state.data["profiles"][app_state.active_profile]["hotkeys"] = []
    if index < 0 or index >= len(app_state.data["profiles"][app_state.active_profile]["titles"]):
        show_critical_message("Fehler", "Ungültiger Eintrag zum Löschen ausgewählt!")
        return
    del app_state.data["profiles"][app_state.active_profile]["titles"][index]
    del app_state.data["profiles"][app_state.active_profile]["texts"][index]
    del app_state.data["profiles"][app_state.active_profile]["hotkeys"][index]
    update_ui() 
    register_hotkeys() 

def move_entry_to(old_index, new_index):
    """Vollständige Entry-Verschiebung mit allen drei Arrays"""
    profile = app_state.data["profiles"][app_state.active_profile]
    titles = profile.get("titles", [])
    texts = profile.get("texts", [])
    hotkeys = profile.get("hotkeys", [])
    if old_index < 0 or old_index >= len(titles):
        return
    if new_index < 0 or new_index >= len(titles):
        return
    if old_index < len(titles):
        title = titles.pop(old_index)
        titles.insert(new_index, title)
    if old_index < len(texts):
        text = texts.pop(old_index)
        texts.insert(new_index, text)
    if old_index < len(hotkeys):
        hotkey = hotkeys.pop(old_index)
        hotkeys.insert(new_index, hotkey)
    max_len = max(len(titles), len(texts), len(hotkeys))
    while len(titles) < max_len:
        titles.append(f"Titel {len(titles)+1}")
    while len(texts) < max_len:
        texts.append(f"Text {len(texts)+1}")
    while len(hotkeys) < max_len:
        hotkeys.append(f"ctrl+shift+{len(hotkeys)+1}")
    update_ui()

#endregion

#region toggle edit mode

def toggle_edit_mode():
    if app_state.edit_mode:
        restored_from_snapshot = False
        if has_field_changes():
            resp = show_question_message(
                "Ungespeicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Willst du sie speichern?")
            if resp == QtWidgets.QMessageBox.Yes:
                save_data()
            else:
                if app_state.last_ui_data is not None:
                    try:
                        restored_profiles = copy.deepcopy(app_state.last_ui_data)
                    except Exception:
                        restored_profiles = None
                    if restored_profiles is not None:
                        app_state.data["profiles"] = restored_profiles
                        active_profile = app_state.data.get("active_profile")
                        if active_profile not in restored_profiles:
                            fallback = None
                            if app_state.active_profile in restored_profiles:
                                fallback = app_state.active_profile
                            if fallback is None:
                                fallback = next((name for name in restored_profiles.keys() if name != "SDE"), None)
                            if fallback is None and restored_profiles:
                                fallback = next(iter(restored_profiles.keys()))
                            if fallback is not None:
                                app_state.active_profile = fallback
                                app_state.data["active_profile"] = fallback
                        else:
                            app_state.active_profile = active_profile
                        restored_from_snapshot = True
                reset_unsaved_changes()
        else:
            reset_unsaved_changes()
        app_state.edit_mode = False
        update_ui()
        if restored_from_snapshot:
            register_hotkeys()
            refresh_tray()
        app_state.last_ui_data = None
        return
    is_sde_only = len(app_state.data["profiles"]) == 1 and "SDE" in app_state.data["profiles"]
    if app_state.active_profile == "SDE" and not is_sde_only:
        show_information_message("Nicht editierbar", "Das SDE-Profil kann nicht bearbeitet werden.")
        return
    try:
        app_state.last_ui_data = copy.deepcopy(app_state.data.get("profiles", {}))
    except Exception:
        app_state.last_ui_data = None
    app_state.unsaved_changes = False
    app_state.edit_mode = True
    update_ui()

#endregion

#region save_data

def save_data(stay_in_edit_mode=False):
    try:
        if app_state.edit_mode and app_state.profile_entries:
            rename_result = apply_profile_renames(show_errors=True)
            if rename_result is None:
                return
            app_state.data["active_profile"] = app_state.active_profile
        if app_state.active_profile != "SDE":
            titles_new = [(e.text() or "").strip() for e in app_state.title_entries]
            if any(not t for t in titles_new):
                show_critical_message("Fehler", "Es gibt leere Titel. Bitte fülle alle Titel aus.")
                return
            low = [t.lower() for t in titles_new]
            if len(low) != len(set(low)):
                show_critical_message("Fehler", "Es gibt doppelte Titel im Profil. Bitte eindeutige Titel vergeben.")
                return
            texts_new = [e.toHtml() if hasattr(e, 'toHtml') else e.text() for e in app_state.text_entries]
            hotkeys_new = [e.text() for e in app_state.hotkey_entries]
            app_state.data["profiles"][app_state.active_profile]["titles"]  = titles_new
            app_state.data["profiles"][app_state.active_profile]["texts"]   = texts_new
            app_state.data["profiles"][app_state.active_profile]["hotkeys"] = hotkeys_new
        profiles_to_save = {k: v for k, v in app_state.data["profiles"].items() if k != "SDE"}
        debounced_saver.schedule_save({"profiles": profiles_to_save, "active_profile": app_state.data["active_profile"]})
        fehlerhafte_hotkeys = register_hotkeys()
        try:
            app_state.last_ui_data = copy.deepcopy(app_state.data.get("profiles", {}))
        except Exception:
            app_state.last_ui_data = None
        reset_unsaved_changes()
        update_ui()
        if not fehlerhafte_hotkeys and not stay_in_edit_mode:
            toggle_edit_mode()
        refresh_tray()
    except Exception as e:
        show_critical_message("Fehler", f"Speichern fehlgeschlagen: {e}")
    reset_unsaved_changes()

def reset_unsaved_changes():
    app_state.unsaved_changes = False
    if app_state.profile_entries:
        for entry in app_state.profile_entries.values():
            clear_pending = getattr(entry, "clear_pending_text", None)
            if clear_pending is not None:
                clear_pending()

def confirm_and_then(action_if_yes):
    if not has_field_changes():
        action_if_yes()
        return
    if action_if_yes.__name__ == "save_data":
        action_if_yes()
        return
    resp = show_question_message(
        "Ungespeicherte Änderungen", 
        "Du hast ungespeicherte Änderungen.\nWillst du sie speichern?",
        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
    if resp == QtWidgets.QMessageBox.Yes:
        save_data(stay_in_edit_mode=True)
    else:
        reset_unsaved_changes()
    QtCore.QTimer.singleShot(100, action_if_yes)

#endregion

#region zoom

def apply_auto_dpi_scaling():
    """Passt die globale Schriftgröße anhand der erkannten DPI an."""
    app_state.zoom_level = detect_optimal_zoom()
    app = QtWidgets.QApplication.instance()
    if not app:
        return
    font_size = max(8, int(round(DEFAULT_FONT_SIZE * app_state.zoom_level)))
    font = app.font()
    font.setPointSize(font_size)
    app.setFont(font)
    app.setStyleSheet(f"* {{ font-size: {font_size}pt; }}")

def calculate_button_text(html_text, button_width):
    """Berechnet dynamisch den Text für Button-Breite"""
    try:
        doc = QtGui.QTextDocument()
        doc.setHtml(html_text)
        plain_text = doc.toPlainText().replace('\n', ' ').strip()
        if not plain_text:
            return "(Leer)"
        font = QFont()
        font.setPointSize(int(app_state.base_font_size * app_state.zoom_level))
        metrics = QFontMetrics(font)
        padding = 30  
        usable_width = max(50, button_width - padding)
        if metrics.horizontalAdvance(plain_text) <= usable_width:
            return plain_text
        ellipsis = "..."
        ellipsis_width = metrics.horizontalAdvance(ellipsis)
        target_width = usable_width - ellipsis_width
        left, right = 0, len(plain_text)
        best_length = 0
        while left <= right:
            mid = (left + right) // 2
            test_text = plain_text[:mid]
            if metrics.horizontalAdvance(test_text) <= target_width:
                best_length = mid
                left = mid + 1
            else:
                right = mid - 1
        if best_length > 0:
            return plain_text[:best_length].rstrip() + ellipsis
        return ellipsis
    except Exception as e:
        logging.warning(f"Text calculation failed: {e}")
        return plain_text[:40] + "..." if len(plain_text) > 40 else plain_text

def create_text_button(i, texts, hks, ebg, fg):
    """Erstellt Text-Button mit dynamischer Größenanpassung"""
    text_html = texts[i] if i < len(texts) else ""
    text_btn = QtWidgets.QPushButton()
    text_btn.setStyleSheet(f"""
        QPushButton {{
            background: {ebg};
            color: {fg};
            text-align: left;
            padding: 8px 12px;
            border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
            border-radius: 6px;}}
        QPushButton:hover {{background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};}}""")
    text_btn.setObjectName("qp_text_btn")   # für das Refresh-Finding
    text_btn.setFixedHeight(int(40 * app_state.zoom_level))
    text_btn.setSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Fixed)
    text_btn.setToolTip(f"Klicken zum Kopieren • Hotkey: {hks[i] if i < len(hks) else ''}")
    def update_button_text():
        if text_btn and not sip.isdeleted(text_btn):
            width = text_btn.width()
            if width > 0:
                display_text = calculate_button_text(text_html, width)
                if text_btn.text() != display_text:
                    text_btn.setText(display_text)
    original_resize = text_btn.resizeEvent
    def on_resize(event):
        if original_resize:
            original_resize(event)
        for ms in (0, 40):
            t = QtCore.QTimer(text_btn)  
            t.setSingleShot(True)
            t.timeout.connect(update_button_text)
            t.start(ms)
    text_btn.resizeEvent = on_resize
    text_btn.clicked.connect(partial(copy_text_to_clipboard, i))
    text_btn._update_text = update_button_text
    for ms in (30, 120):
        t = QtCore.QTimer(text_btn)
        t.setSingleShot(True)
        t.timeout.connect(update_button_text)
        t.start(ms)
    return text_btn

def initialize_application():
    """Initialisiert die Anwendung mit korrektem Scaling"""
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
        except:
            pass
    return app

#endregion

#region Hauptfenster

app = initialize_application()
if sys.platform.startswith("win"):
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
win = QtWidgets.QMainWindow()

def setup_window_scaling():
    QtCore.QTimer.singleShot(0, apply_auto_dpi_scaling)
win.setWindowTitle("QuickPaste")
win.setMinimumSize(399, 100)
app_state.normal_minimum_width = win.minimumWidth()
win.setWindowIcon(QtGui.QIcon(ICON_PATH) if os.path.exists(ICON_PATH) else win.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon))
QtCore.QTimer.singleShot(500, lambda: win.setWindowIcon(QtGui.QIcon(ICON_PATH) if os.path.exists(ICON_PATH) else win.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon)))
win.statusBar().showMessage("Bereit")
def close_event_handler(event):
    if not win.isVisible():
        event.ignore()
        return
    save_window_position()
    minimize_to_tray()
    event.ignore()
win.closeEvent = close_event_handler
central = QtWidgets.QWidget()
main_layout = QtWidgets.QVBoxLayout(central)
main_layout.setContentsMargins(0,0,0,0)
main_layout.setSpacing(0)
win.setCentralWidget(central)
toolbar = QtWidgets.QToolBar()
toolbar.setMovable(False)
win.addToolBar(toolbar)
scroll_area = QtWidgets.QScrollArea()
scroll_area.setWidgetResizable(True)
scroll_area.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
main_layout.addWidget(scroll_area)
container = QtWidgets.QWidget()
entries_layout = QtWidgets.QVBoxLayout(container)
entries_layout.setAlignment(QtCore.Qt.AlignTop)
entries_layout.setSpacing(6)
entries_layout.setContentsMargins(8, 8, 8, 8)
scroll_area.setWidget(container)
bottom_bar_container = QtWidgets.QWidget()
bottom_bar_layout = QtWidgets.QHBoxLayout(bottom_bar_container)
bottom_bar_layout.setContentsMargins(8, 8, 8, 8)
bottom_bar_layout.setSpacing(8)
bottom_bar_container.setVisible(False)
main_layout.addWidget(bottom_bar_container)

#endregion

#region UI

def show_text_context_menu(pos, text_widget):
    menu = QtWidgets.QMenu(text_widget)
    cursor = text_widget.textCursor()
    has_selection = cursor.hasSelection()
    selected_text = cursor.selectedText() if has_selection else ""
    undo_action = menu.addAction("Undo")
    undo_action.setShortcut("Ctrl+Z")
    undo_action.triggered.connect(text_widget.undo)
    undo_action.setEnabled(text_widget.document().isUndoAvailable())
    redo_action = menu.addAction("Redo")
    redo_action.setShortcut("Ctrl+Y")
    redo_action.triggered.connect(text_widget.redo)
    redo_action.setEnabled(text_widget.document().isRedoAvailable())
    menu.addSeparator()
    cut_action = menu.addAction("Cut")
    cut_action.setShortcut("Ctrl+X")
    cut_action.triggered.connect(text_widget.cut)
    cut_action.setEnabled(has_selection)
    copy_action = menu.addAction("Copy")
    copy_action.setShortcut("Ctrl+C")
    copy_action.triggered.connect(text_widget.copy)
    copy_action.setEnabled(has_selection)
    paste_action = menu.addAction("Paste")
    paste_action.setShortcut("Ctrl+V")
    paste_action.triggered.connect(text_widget.paste)
    paste_action.setEnabled(bool(QtWidgets.QApplication.clipboard().text().strip()))
    delete_action = menu.addAction("Delete")
    delete_action.triggered.connect(lambda: cursor.removeSelectedText() if has_selection else None)
    delete_action.setEnabled(has_selection)
    menu.addSeparator()
    select_all_action = menu.addAction("Select All")
    select_all_action.setShortcut("Ctrl+A")
    select_all_action.triggered.connect(text_widget.selectAll)
    menu.addSeparator()
    if has_selection:
        add_link_action = menu.addAction("Add Hyperlink...")
        add_link_action.triggered.connect(lambda: add_hyperlink_to_selection(text_widget, cursor))
        if cursor.charFormat().isAnchor():
            remove_link_action = menu.addAction("Remove Hyperlink")
            remove_link_action.triggered.connect(lambda: remove_hyperlink_from_selection(text_widget, cursor))
    else:
        insert_link_action = menu.addAction("Insert Hyperlink...")
        insert_link_action.triggered.connect(lambda: insert_hyperlink_at_cursor(text_widget))
    if app_state.dark_mode:
        menu.setStyleSheet("""
            QMenu {
                background-color: #2e2e2e;
                color: white;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 2px;}
            QMenu::item {
                background-color: transparent;
                padding: 6px 20px;
                border-radius: 3px;}
            QMenu::item:selected {
                background-color: #4a90e2;
                color: white;}
            QMenu::item:disabled {
                color: #888;}
            QMenu::separator {
                height: 1px;
                background-color: #555;
                margin: 2px 10px;}""")
    global_pos = text_widget.mapToGlobal(pos)
    menu.exec_(global_pos)

def add_hyperlink_to_selection(text_widget, cursor):
    selected_text = cursor.selectedText()
    url, ok = QtWidgets.QInputDialog.getText(
        text_widget, 
        "Add Hyperlink", 
        f"Enter URL for '{selected_text}':",
        text="https://")
    if ok and url.strip():
        link_format = QtGui.QTextCharFormat()
        link_format.setAnchor(True)
        link_format.setAnchorHref(url.strip())
        link_format.setForeground(QtGui.QColor("#0066cc" if not app_state.dark_mode else "#4da6ff"))
        link_format.setUnderlineStyle(QtGui.QTextCharFormat.SingleUnderline)
        cursor.mergeCharFormat(link_format)
        text_widget.setTextCursor(cursor)

def remove_hyperlink_from_selection(text_widget, cursor):
    normal_format = QtGui.QTextCharFormat()
    normal_format.setAnchor(False)
    normal_format.setAnchorHref("")
    normal_format.setForeground(QtGui.QColor("white" if app_state.dark_mode else "black"))
    normal_format.setUnderlineStyle(QtGui.QTextCharFormat.NoUnderline)
    cursor.mergeCharFormat(normal_format)
    text_widget.setTextCursor(cursor)

def insert_hyperlink_at_cursor(text_widget):
    display_text, ok1 = QtWidgets.QInputDialog.getText(
        text_widget, 
        "Insert Hyperlink", 
        "Enter display text:")
    if ok1 and display_text.strip():
        url, ok2 = QtWidgets.QInputDialog.getText(
            text_widget, 
            "Insert Hyperlink", 
            f"Enter URL for '{display_text.strip()}':",
            text="https://")
        if ok2 and url.strip():
            cursor = text_widget.textCursor()
            link_format = QtGui.QTextCharFormat()
            link_format.setAnchor(True)
            link_format.setAnchorHref(url.strip())
            link_format.setForeground(QtGui.QColor("#0066cc" if not app_state.dark_mode else "#4da6ff"))
            link_format.setUnderlineStyle(QtGui.QTextCharFormat.SingleUnderline)
            cursor.insertText(display_text.strip(), link_format)
            normal_format = QtGui.QTextCharFormat()
            normal_format.setAnchor(False)
            normal_format.setForeground(QtGui.QColor("white" if app_state.dark_mode else "black"))
            normal_format.setUnderlineStyle(QtGui.QTextCharFormat.NoUnderline)
            cursor.setCharFormat(normal_format)
            text_widget.setTextCursor(cursor)

def update_ui():
    app_state.title_entries = []
    app_state.text_entries = []
    app_state.hotkey_entries = []
    app_state.profile_entries = {}
    bg    = "#2e2e2e" if app_state.dark_mode else "#eeeeee"
    fg    = "white"   if app_state.dark_mode else "black"
    ebg   = "#3c3c3c" if app_state.dark_mode else "white"
    bbg   = "#444"    if app_state.dark_mode else "#cccccc"
    win.setStyleSheet(f"background:{bg};")
    toolbar.setStyleSheet(f"background:{bg}; border: none;")
    container.setStyleSheet(f"background:{bg};")
    win.statusBar().setStyleSheet(f"""
        QStatusBar {{
            background: {bg};
            color: {fg};
            border-top: 1px solid #666;}}""")
    entries_margin = 4 if app_state.mini_mode else 8
    entries_layout.setContentsMargins(entries_margin, entries_margin, entries_margin, entries_margin)
    entries_layout.setSpacing(4 if app_state.mini_mode else 6)
    bottom_bar_container.setStyleSheet(f"background:{bg};")
    bottom_bar_layout.setContentsMargins(entries_margin, entries_margin, entries_margin, entries_margin)
    while bottom_bar_layout.count():
        item = bottom_bar_layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()
    if app_state.edit_mode:
        bottom_bar_container.setVisible(True)
        bottom_bar_container.setEnabled(True)
        button_border_color = '#555' if app_state.dark_mode else '#ccc'
        button_hover_bg = '#4a4a4a' if app_state.dark_mode else '#f0f0f0'
        button_min_height = int(40 * app_state.zoom_level)
        save_button = QtWidgets.QPushButton("💾 Speichern")
        save_button.setMinimumHeight(button_min_height)
        save_button.setStyleSheet(
            f"""
            QPushButton {{
                background: #2e7d32;
                color: white;
                border: 1px solid {button_border_color};
                border-radius: 6px;
                padding: 8px 12px;
                min-height: 10px;}}
            QPushButton:hover {{background: #388e3c;}}"""
        )
        save_button.clicked.connect(save_data)
        bottom_bar_layout.addWidget(save_button)
        add_button = QtWidgets.QPushButton("➕ Eintrag hinzufügen")
        add_button.setMinimumHeight(button_min_height)
        add_button.setStyleSheet(
            f"""
            QPushButton {{
                background: {bbg};
                color: {fg};
                border: 1px solid {button_border_color};
                border-radius: 6px;
                padding: 8px 12px;
                min-height: 10px;}}
            QPushButton:hover {{background: {button_hover_bg};}}"""
        )
        add_button.clicked.connect(add_new_entry)
        bottom_bar_layout.addWidget(add_button)
    else:
        bottom_bar_container.setVisible(False)
        bottom_bar_container.setEnabled(False)
    toolbar.clear()
    app_state.profile_buttons = {}
    app_state.profile_selector = None
    app_state.profile_delete_button = None
    profile_names = []
    for profile_name in app_state.data["profiles"].keys():
        if profile_name == "SDE":
            continue
        profile_names.append(profile_name)
    if not app_state.edit_mode and "SDE" in app_state.data["profiles"]:
        profile_names.append("SDE")
    if profile_names:
        selector_container = QtWidgets.QWidget()
        selector_container.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        selector_layout = QtWidgets.QHBoxLayout(selector_container)
        selector_layout.setContentsMargins(0, 0, 0, 0)
        selector_layout.setSpacing(1 if app_state.mini_mode else 3)
        def scaled(value):
            return max(1, int(value * app_state.zoom_level))
        
        combo = ProfileComboBox()
        combo.setEditable(app_state.edit_mode)
        combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        combo.setSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        combo_height = scaled(24 if app_state.mini_mode else 32)
        combo.setFixedHeight(combo_height)

        if app_state.mini_mode:
            combo.setMinimumWidth(scaled(90))
            combo.setMaximumWidth(scaled(110))
            drop_width = scaled(16)
        else:
            combo.setMinimumWidth(scaled(140))
            combo.setMaximumWidth(scaled(200))
            drop_width = scaled(26)

        radius = scaled(5)  # GEÄNDERT: Einheitliche Rundung für alle Buttons
        padding_v = scaled(3 if app_state.mini_mode else 6)
        padding_h = scaled(8 if app_state.mini_mode else 14)
        border_color = "#555" if app_state.dark_mode else "#ccc"
        combo.setStyleSheet(f"""
            QComboBox {{
                background:{bbg};
                color:{fg};

                
                border: 1px solid {border_color};
                border-radius:{radius}px;
                padding:{padding_v}px {drop_width + padding_v}px {padding_v}px {padding_h}px;}}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width:{drop_width}px;
                border-left: 1px solid {border_color};
                border-top-right-radius:{radius}px;
                border-bottom-right-radius:{radius}px;
                background:{bbg};
                margin:0; padding:0;}}
            QComboBox QAbstractItemView {{
                background:{ebg};
                color:{fg};
                border: 1px solid {border_color};
                selection-background-color:#4a90e2;
                selection-color:white;}}""")
        
        selector_layout.addWidget(combo)
        app_state.profile_selector = combo
        
        delete_btn = None
        if app_state.edit_mode:
            delete_btn = QtWidgets.QPushButton("❌")
            btn_size = scaled(24 if app_state.mini_mode else 32)  # GEÄNDERT: Gleiche Höhe wie Combo
            delete_btn.setFixedSize(btn_size, btn_size)
            delete_btn.setStyleSheet(
                f"""
                QPushButton {{
                    background:{bbg}; color:{fg};
                    color:white;
                    border: 1px;
                    border-radius: {radius}px;}}  
                QPushButton:hover {{background: {'#666'};}}
                QPushButton:pressed {{background:#b71c1c;}}""")
            delete_btn.setToolTip("Ausgewähltes Profil löschen")
            selector_layout.addWidget(delete_btn)
            app_state.profile_delete_button = delete_btn

        toolbar.addWidget(selector_container)
        for name in profile_names:
            combo.addItem(name, name)
        if app_state.edit_mode:
            line_edit = combo.lineEdit()
            if line_edit is not None:
                line_edit.setPlaceholderText("Profilnamen bearbeiten")
                line_edit.setStyleSheet(
                    f"color:{fg}; background:transparent; border:none; padding:0px;")
                line_edit.textEdited.connect(_remember_profile_name_edit)
            for idx in range(combo.count()):
                original = combo.itemData(idx)
                if original and original != "SDE":
                    app_state.profile_entries[original] = ComboBoxItemProxy(combo, idx)
        def update_delete_state():
            if delete_btn is None:
                return
            index = combo.currentIndex()
            if index < 0:
                delete_btn.setEnabled(False)
                return
            original = combo.itemData(index)
            can_delete = (combo.count() > 1 and original not in (None, "SDE"))
            delete_btn.setEnabled(can_delete)
        def on_profile_changed(index):
            update_delete_state()
            if index < 0:
                return
            selected = combo.itemData(index) or combo.itemText(index)
            if selected != app_state.active_profile:
                switch_profile(selected)
        combo.currentIndexChanged.connect(on_profile_changed)
        if delete_btn is not None:
            def on_delete_clicked():
                index = combo.currentIndex()
                if index < 0:
                    return
                target = combo.itemData(index) or combo.itemText(index)
                delete_profile(target)
            delete_btn.clicked.connect(on_delete_clicked)
        current_index = next(
            (i for i in range(combo.count()) if combo.itemData(i) == app_state.active_profile),-1,)
        with QtCore.QSignalBlocker(combo):
            if current_index >= 0:
                combo.setCurrentIndex(current_index)
            elif combo.count() > 0:
                combo.setCurrentIndex(0)
        update_delete_state()

    if app_state.edit_mode:
        ap = QtWidgets.QPushButton("➕ Profil")
        ap.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border: 1px solid {border_color};
                border-radius: 5px;
                padding: 6px 16px;
                margin-left: 8px;}}  
            QPushButton:hover {{background:#666;}}""")
        ap.clicked.connect(add_new_profile)
        toolbar.addWidget(ap)
        
    spacer = QtWidgets.QWidget()
    spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    toolbar.addWidget(spacer)
    
    control_size = 26 if app_state.mini_mode else 30
    control_radius = 12 if app_state.mini_mode else 15
    control_margin = 4 if app_state.mini_mode else 6
    controls = [
        ("🌙" if not app_state.dark_mode else "🌞", toggle_dark_mode, "Dunkelmodus umschalten")]
    if not app_state.edit_mode:
        controls.append(("🗕" if not app_state.mini_mode else "🗖", toggle_mini_mode, "Mini-Ansicht umschalten"))
    if not app_state.mini_mode:
        controls.append(("🔧", toggle_edit_mode, "Bearbeitungsmodus umschalten"))
    for text, func, tooltip in controls:
        b = QtWidgets.QPushButton(text)
        b.setToolTip(tooltip)
        b.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border-radius: {control_radius}px;
                min-width: {control_size}px; min-height: {control_size}px;
                margin-left: {control_margin}px;
                border: none;
                padding: 0;}}
            QPushButton:hover {{background:#888;}}""")
        b.clicked.connect(func)
        toolbar.addWidget(b)
    help_btn = QtWidgets.QPushButton("❓")
    help_btn.setToolTip("Hilfe anzeigen")
    help_btn.setStyleSheet(
        f"""
        QPushButton {{
            background:{bbg}; color:{fg};
            border-radius: {control_radius}px;
            min-width: {control_size}px; min-height: {control_size}px;
            margin-left: {control_margin}px;
            border: none;
            padding: 0;}}
        QPushButton:hover {{background:#888;}}""")
    help_btn.clicked.connect(show_help_dialog)
    toolbar.addWidget(help_btn)
    while entries_layout.count():
        w = entries_layout.takeAt(0).widget()
        if w: w.deleteLater()
    prof_data = app_state.data["profiles"][app_state.active_profile]
    titles, texts, hks = prof_data["titles"], prof_data["texts"], prof_data["hotkeys"]
    max_t = 120
    max_h = 120
    for i, title in enumerate(titles):
        if not app_state.edit_mode and app_state.mini_mode:
            hotkey = hks[i] if i < len(hks) else ""
            title_text = title or ""
            mini_button = QtWidgets.QPushButton()
            mini_hotkey_color = '#d0d0d0' if app_state.dark_mode else '#333333'
            mini_button.setStyleSheet(f"""
                QPushButton {{
                    background: {ebg};
                    border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                    border-radius: 6px;
                    padding: 1px 4px;}}
                QPushButton:hover {{
                    background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};}}
                QPushButton QLabel {{
                    color: {fg};
                    font-weight: bold;
                    background: transparent;
                    padding: 0;}}
                QPushButton QLabel#miniHotkeyLabel {{
                    font-weight: normal;
                    padding-left: 4px;
                    padding-right: 2px;
                    font-size: 12px;
                    color: {mini_hotkey_color};}}""")
            mini_button.setFixedHeight(int(30 * app_state.zoom_level))
            mini_button.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
            mini_layout = QtWidgets.QHBoxLayout()
            mini_layout.setContentsMargins(2, 1, 2, 1)
            mini_layout.setSpacing(2)
            mini_button.setLayout(mini_layout)
            title_label = QtWidgets.QLabel(title_text)
            title_label.setAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
            title_label.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
            mini_layout.addWidget(title_label, 1)
            if hotkey:
                hotkey_label = QtWidgets.QLabel(hotkey)
                hotkey_label.setObjectName("miniHotkeyLabel")
                hotkey_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                hotkey_label.setSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Preferred)
                hotkey_label.setMinimumWidth(0)
                mini_layout.addWidget(hotkey_label)
            tooltip_hotkey = hotkey or ""
            if tooltip_hotkey:
                mini_button.setToolTip(f"Klicken zum Kopieren ➡️ Hotkey: {tooltip_hotkey}")
            else:
                mini_button.setToolTip("Klicken zum Kopieren")
            mini_button.clicked.connect(partial(copy_text_to_clipboard, i))
            entries_layout.addWidget(mini_button)
            continue
        if app_state.edit_mode:
            row = DragDropWidget(i)
        else:
            row = QtWidgets.QWidget()
        hl  = QtWidgets.QHBoxLayout(row)
        hl.setContentsMargins(8, 4, 8, 4)
        hl.setSpacing(12)
        hl.setStretch(0, 0)
        hl.setStretch(1, 1) 
        hl.setStretch(2, 0) 
        if app_state.edit_mode:
            drag_handle = QtWidgets.QLabel("☰")
            drag_handle.setFixedSize(20, 28)
            drag_handle.setStyleSheet(f"""
                color: {fg}; 
                background: {bbg}; 
                padding: 2px 4px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: 4px;
                text-align: center;""")
            drag_handle.setAlignment(QtCore.Qt.AlignCenter)
            drag_handle.setToolTip("Ziehen zum Verschieben")
            row.drag_index = i
            drag_handle.drag_index = i
            row.setAcceptDrops(True)
            drag_handle.mousePressEvent = lambda event, idx=i: start_drag(event, idx, drag_handle)
            hl.addWidget(drag_handle)
        if app_state.edit_mode:
            et = QtWidgets.QLineEdit(title)
            et.setFixedWidth(max_t)
            et.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if app_state.dark_mode else '#ccc'}; border-radius: 6px; padding: 8px;")
            def validate_and_set_title(idx, widget):
                new_title = (widget.text() or "").strip()
                profile = app_state.data["profiles"].get(app_state.active_profile, {})
                titles_list = profile.get("titles", [])
                old = titles_list[idx] if idx < len(titles_list) else ""
                if not new_title:
                    show_critical_message("Fehler", "Titel darf nicht leer sein!")
                    widget.setText(old)
                    return
                current_titles = [e.text().strip().lower() for j, e in enumerate(app_state.title_entries) if j != idx]
                if new_title.lower() in current_titles:
                    show_critical_message("Fehler", f"Titel '{new_title}' wird bereits verwendet!")
                    widget.setText(old)
                    return
                if new_title == old:
                    return
                app_state.data["profiles"][app_state.active_profile]["titles"][idx] = new_title
                app_state.unsaved_changes = True
            et.editingFinished.connect(partial(validate_and_set_title, i, et))
            hl.addWidget(et)
            app_state.title_entries.append(et)
        else:
            lt = QtWidgets.QLabel(title)
            lt.setFixedWidth(max_t)
            lt.setFixedHeight(40)
            lt.setStyleSheet(f"""
                color: {fg}; 
                background: {ebg}; 
                font-weight: bold; 
                padding: 10px 12px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: 6px;""")
            lt.setAlignment(QtCore.Qt.AlignVCenter)
            hl.addWidget(lt)
        if app_state.edit_mode:
            ex = QtWidgets.QTextEdit(texts[i])
            ex.setMaximumHeight(80)
            ex.setMinimumHeight(60)
            ex.setStyleSheet(f"background:{ebg}; color:{fg};")
            ex.setAcceptRichText(True)
            ex.setHtml(texts[i])
            ex.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
            ex.customContextMenuRequested.connect(lambda pos, w=ex: show_text_context_menu(pos, w))
            def update_text(widget, index):
                new_html = widget.toHtml()
                texts_list = app_state.data["profiles"][app_state.active_profile]["texts"]
                old_html = texts_list[index] if index < len(texts_list) else None
                if new_html == old_html:
                    return
                texts_list[index] = new_html
                app_state.unsaved_changes = True
            ex.textChanged.connect(partial(update_text, ex, i))
            hl.addWidget(ex, 1)
            app_state.text_entries.append(ex)
        else:
            text_btn = create_text_button(i, texts, hks, ebg, fg)
            hl.addWidget(text_btn, 1)
        if app_state.edit_mode:
            eh = QtWidgets.QLineEdit(hks[i])
            eh.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if app_state.dark_mode else '#ccc'}; border-radius: 6px; padding: 8px;")
            def validate_and_set_hotkey(idx, widget):
                hotkey = widget.text().strip().lower()
                erlaubte_zeichen = "1234567890befhmpqvxz§'^"
                if hotkey:
                    parts = hotkey.split("+")
                    if (len(parts) != 3 or 
                        parts[0] != "ctrl" or 
                        parts[1] != "shift" or 
                        parts[2] not in erlaubte_zeichen):
                        show_critical_message(
                            "Fehler",
                            f"Ungültiger Hotkey \"{widget.text()}\" für Eintrag {idx+1}.\n"
                            f"Erlaubte Zeichen: {erlaubte_zeichen}\n"
                            f"Format: ctrl+shift+[zeichen]")
                        widget.setText(hks[idx])
                        return
                    current_hotkeys = [entry.text().strip().lower() for j, entry in enumerate(app_state.hotkey_entries) if j != idx]
                    if hotkey in current_hotkeys:
                        show_critical_message(
                            "Fehler",
                            f"Hotkey \"{widget.text()}\" wird bereits in diesem Profil verwendet!")
                        widget.setText(hks[idx])
                        return
                current_list = app_state.data["profiles"][app_state.active_profile]["hotkeys"]
                old_hotkey = current_list[idx] if idx < len(current_list) else ""
                if widget.text() == old_hotkey:
                    return
                app_state.data["profiles"][app_state.active_profile]["hotkeys"][idx] = widget.text()
                app_state.unsaved_changes = True
            eh.editingFinished.connect(partial(validate_and_set_hotkey, idx=i, widget=eh))
            hl.addWidget(eh)
            app_state.hotkey_entries.append(eh)
            delete_btn = QtWidgets.QPushButton("❌")
            delete_size = int(38 * app_state.zoom_level)
            delete_btn.setFixedSize(delete_size, delete_size)
            delete_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {ebg};
                    color: {fg};
                    border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                    border-radius: 6px;
                    padding: 8px;}}
                QPushButton:hover {{background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};}}
                QPushButton:pressed {{background: {'#3a3a3a' if app_state.dark_mode else '#e0e0e0'};}}""")
            delete_btn.clicked.connect(lambda _, j=i: delete_entry(j))
            delete_btn.setToolTip("Eintrag löschen")
            hl.addWidget(delete_btn)
        else:
            lh = QtWidgets.QLabel(hks[i])
            lh.setFixedHeight(40)
            lh.setStyleSheet(f"""
                color: {fg}; 
                background: {ebg}; 
                padding: 10px 12px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: 6px;
                font-family: 'Consolas', 'Monaco', monospace;""")
            lh.setAlignment(QtCore.Qt.AlignCenter)
            hl.addWidget(lh)
        entries_layout.addWidget(row)

#endregion

#region help 

def show_help_dialog():
    help_text = (
        "QuickPaste Hilfe\n\n"
        "• 🌙/🌞 Dunkelmodus: Wechselt zwischen hell/dunkel.\n"
        "• 🔧 Bearbeiten: Titel, Texte und Hotkeys anpassen.\n"
        "• 🗕/🗖 Mini-Ansicht umschalten \n\n"
        "• ➕ Profil: Neues Textprofil erstellen.\n"
        "• 🖊️ Im Bearbeitungsmodus zwischen Profilen wechseln.\n"
        "• ❌ Löschen: Profil entfernen.\n\n"
        "• ☰ Verschieben: Einträge per Drag & Drop umsortieren.\n"
        "• ❌ Eintrag löschen.\n"
        "• ➕ Eintrag hinzufügen: Fügt einen neuen Eintrag hinzu.\n"
        "• 💾 Speichern: Änderungen sichern.\n\n"
        "Text markieren + Rechtsklick kann ein Hyperlink hinterlegt werden. \n\n"
        "Bei Fragen oder Problemen: nico.wagner@bit.admin.ch")
    show_information_message("QuickPaste Hilfe", help_text)

#endregion

#region darkmode/minimode/messagebox

def apply_dark_mode_to_messagebox(msg):
    if app_state.dark_mode:
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #2e2e2e;
                color: white;}
            QMessageBox QLabel {
                color: white !important;}
            QMessageBox QPushButton {
                background-color: #444 !important;
                color: white !important;
                border: 1px solid #666;
                border-radius: 5px;
                min-width: 60px;
                min-height: 24px;
                padding: 4px 8px;
                font-weight: normal;}
            QMessageBox QPushButton:hover {
                background-color: #666 !important;
                color: white !important;}
            QMessageBox QPushButton:pressed {
                background-color: #555 !important;
                color: white !important;}
            QMessageBox QPushButton:focus {
                background-color: #4a90e2 !important;
                color: white !important;
                border: 1px solid #5aa3f0;}""")
        for button in msg.findChildren(QtWidgets.QPushButton):
            button.setStyleSheet("""
                QPushButton {
                    background-color: #444;
                    color: white !important;
                    border: 1px solid #666;
                    border-radius: 5px;
                    min-width: 60px;
                    min-height: 24px;
                    padding: 4px 8px;}
                QPushButton:hover {
                    background-color: #666;
                    color: white !important;}
                QPushButton:pressed {
                    background-color: #555;
                    color: white !important;}""")

def show_critical_message(title, text, parent=None):
    if parent is None:
        parent = win
    msg = QtWidgets.QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setIcon(QtWidgets.QMessageBox.Critical)
    apply_dark_mode_to_messagebox(msg)
    return msg.exec_()

def show_question_message(title, text, buttons=QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, parent=None):
    if parent is None:
        parent = win
    msg = QtWidgets.QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setStandardButtons(buttons)
    msg.setIcon(QtWidgets.QMessageBox.Question)
    apply_dark_mode_to_messagebox(msg)
    return msg.exec_()

def show_information_message(title, text, parent=None):
    if parent is None:
        parent = win
    msg = QtWidgets.QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setIcon(QtWidgets.QMessageBox.Information)
    apply_dark_mode_to_messagebox(msg)
    msg.exec_()

def toggle_dark_mode():
    app_state.dark_mode = not app_state.dark_mode
    update_ui()
    save_window_position()

def calculate_mini_mode_size():
    base_width = 200
    base_height = 260
    central_widget = win.centralWidget()
    if central_widget is not None:
        layout = central_widget.layout()
        if layout is not None:
            layout.activate()
    hint = win.sizeHint()
    if hint.isValid():
        width_hint = hint.width()
        if app_state.normal_minimum_width:
            max_allowed = max(base_width, app_state.normal_minimum_width - 200)
        else:
            max_allowed = base_width
        width = max(base_width, min(width_hint, max_allowed))
        height = max(base_height, min(hint.height(), 360))
        return int(width), int(height)
    return base_width, base_height

def toggle_mini_mode():
    app_state.mini_mode = not app_state.mini_mode
    if app_state.mini_mode:
        try:
            app_state.saved_geometry = win.saveGeometry()
        except Exception:
            app_state.saved_geometry = None
        if app_state.normal_minimum_width is None:
            app_state.normal_minimum_width = win.minimumWidth()
        update_ui()
        mini_width, mini_height = calculate_mini_mode_size()
        win.setMinimumWidth(mini_width)
        win.resize(mini_width, mini_height)
    else:
        if app_state.normal_minimum_width is not None:
            win.setMinimumWidth(app_state.normal_minimum_width)
        else:
            win.setMinimumWidth(399)
        if app_state.saved_geometry is not None:
            win.restoreGeometry(app_state.saved_geometry)
        else:
            win.resize(700, 500)
        app_state.saved_geometry = win.saveGeometry()
        update_ui()
    save_window_position()

#endregion

app.aboutToQuit.connect(lambda: (debounced_saver.timer.stop(), debounced_saver._save()))
load_window_position()
update_ui()
create_tray_icon()
register_hotkeys()
win.show()
setup_window_scaling()
sys.exit(app.exec_())