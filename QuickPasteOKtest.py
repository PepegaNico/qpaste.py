import sys, os, json, re, ctypes, logging
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

DEFAULT_FONT_SIZE = 9  # Basis-Schriftgröße
MIN_FONT_SIZE = 9
MAX_FONT_SIZE = 9

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

        # VEREINFACHTE Font-Verwaltung
        self.zoom_level = 1.0  # 75% bis 150%
        self.base_font_size = 9  # Basis-Größe

# Global application state
app_state = QuickPasteState()

class ComboBoxItemProxy:
    """Wrap QComboBox items to mimic QLineEdit behaviour."""

    def __init__(self, combo_box, index):
        self._combo_box = combo_box
        self._index = index

    def text(self):
        return self._combo_box.itemText(self._index)

    def setText(self, value):
        self._combo_box.setItemText(self._index, value)


class ComboArrowGlyphStyle(QtWidgets.QProxyStyle):
    GLYPH = "▼"   # Alternativen: "▼", "⏷", "⌄", "🞃"

    def drawComplexControl(self, control, option, painter, widget=None):
        if control == QtWidgets.QStyle.CC_ComboBox and isinstance(option, QtWidgets.QStyleOptionComboBox):
            # Erst die Standard-ComboBox zeichnen (ohne Pfeil)
            super().drawComplexControl(control, option, painter, widget)

            # Pfeil-Bereich berechnen
            arrow_rect = self.subControlRect(
                control,
                option,
                QtWidgets.QStyle.SC_ComboBoxArrow,
                widget
            )

            # Fallback wenn arrow_rect ungültig ist
            if not arrow_rect.isValid() or arrow_rect.width() <= 0 or arrow_rect.height() <= 0:
                # Manuelle Berechnung des Pfeil-Bereichs
                drop_w = 26
                r = option.rect
                arrow_rect = QtCore.QRect(r.right() - drop_w, r.top(), drop_w, r.height())

            # Separator links vom Pfeil-Bereich zeichnen (optional)
            sep_color = widget.palette().mid().color() if widget else QtGui.QColor("#888")
            painter.save()
            painter.setPen(QtGui.QPen(sep_color))
            painter.drawLine(
                arrow_rect.left() - 1, 
                arrow_rect.top() + 2, 
                arrow_rect.left() - 1, 
                arrow_rect.bottom() - 2
            )
            painter.restore()

            # Pfeil-Symbol zeichnen (gefülltes Dreieck, damit der Kontrast stimmt)
            painter.save()
            painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

            # Farbe für den Pfeil (fällt auf hellem wie dunklem Hintergrund auf)
            base_color = option.palette.buttonText().color()
            if not option.state & QtWidgets.QStyle.State_Enabled:
                base_color = option.palette.color(QtGui.QPalette.Disabled, QtGui.QPalette.ButtonText)
            arrow_color = QtGui.QColor(base_color)
            painter.setBrush(arrow_color)
            painter.setPen(QtGui.QPen(arrow_color))

            # Dreieck proportional zur verfügbaren Fläche aufziehen
            inset = max(2, int(min(arrow_rect.width(), arrow_rect.height()) * 0.2))
            triangle_rect = arrow_rect.adjusted(inset, inset, -inset, -inset)
            size = min(triangle_rect.width(), triangle_rect.height())
            center = triangle_rect.center()
            half_width = size * 0.45
            half_height = size * 0.35
            points = [
                QtCore.QPointF(center.x() - half_width, center.y() - half_height),
                QtCore.QPointF(center.x() + half_width, center.y() - half_height),
                QtCore.QPointF(center.x(), center.y() + half_height),
            ]
            painter.drawPolygon(QtGui.QPolygonF(points))
            painter.restore()
            return

        # Für alle anderen Controls die Standard-Implementierung verwenden
        super().drawComplexControl(control, option, painter, widget)


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
                self.pending_data = None   # ← nach dem Schreiben leeren


def save_data_atomic(data, filename):
    """Atomic file write to prevent corruption"""
    dirpath = os.path.dirname(filename)
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(dir=dirpath, prefix=".tmp_", suffix=".json")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        os.replace(tmp, filename)  # atomarer Replace auf NTFS
    except Exception as e:
        if tmp:
            try: os.unlink(tmp)
            except Exception: pass
        raise e


# Initialize debounced saver
debounced_saver = DebouncedSaver(600)


#region window position 

def save_window_position():
    """Speichert Fensterposition und Zoom-Level"""
    try:
        geo_bytes = win.saveGeometry()
        geo_hex = bytes(geo_bytes.toHex()).decode()
        cfg = {
            "geometry_hex": geo_hex,
            "dark_mode": app_state.dark_mode,
            "zoom_level": app_state.zoom_level,  # Zoom speichern
            "mini_mode": app_state.mini_mode
        }
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
    """Lädt Fensterposition und Zoom-Level"""
    try:
        with open(WINDOW_CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        
        if cfg.get("dark_mode") is not None:
            app_state.dark_mode = cfg["dark_mode"]

        if cfg.get("mini_mode") is not None:
            app_state.mini_mode = cfg["mini_mode"]

        # Zoom laden (mit Limits)
        if cfg.get("zoom_level") is not None:
            app_state.zoom_level = max(0.75, min(1.5, cfg["zoom_level"]))
        else:
            # Erste Nutzung: Auto-DPI Detection
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
        # Keine Config: Auto-DPI für ersten Start
        app_state.zoom_level = detect_optimal_zoom()
        return False

def detect_optimal_zoom():
    """Erkennt optimalen Zoom basierend auf System-DPI"""
    try:
        screen = QtWidgets.QApplication.primaryScreen()
        if screen:
            logical_dpi = screen.logicalDotsPerInch()
            # Windows Standard ist 96 DPI
            # 120 DPI = 125% Skalierung
            # 144 DPI = 150% Skalierung
            if logical_dpi <= 96:
                return 1.0  # 100%
            elif logical_dpi <= 120:
                return 1.1  # 110%
            elif logical_dpi <= 144:
                return 1.2  # 120%
            else:
                return 1.3  # 130% für sehr hohe DPI
    except Exception:
        pass
    return 1.0  # Fallback




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
            "hotkeys":["ctrl+shift+1",    "ctrl+shift+2",    "ctrl+shift+3"]
        }
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
                    "Profil 1": {"titles": [], "texts": [], "hotkeys": []}
                }
                loaded["active_profile"] = "Profil 1"
        return loaded
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "profiles": {
                "Profil 1": {
                    "titles":  [f"Titel {i}"        for i in range(1,6)],
                    "texts":   [f"Text {i}"         for i in range(1,6)],
                    "hotkeys": [f"ctrl+shift+{i}"   for i in range(1,6)]
                },
                "Profil 2": {
                    "titles":  [f"Titel {i}"        for i in range(1,6)],
                    "texts":   [f"Profil 2 Text {i}"for i in range(1,6)],
                    "hotkeys": [f"ctrl+shift+{i}"   for i in range(1,6)]
                },
                "SDE": load_sde_profile()
            },
            "active_profile": "Profil 1"
        }
# Initialize application data
app_state.data = load_data()
app_state.active_profile = app_state.data.get("active_profile", list(app_state.data["profiles"].keys())[0])

#endregion 

#region profiles

def has_field_changes(profile_to_check=None):
    if profile_to_check is None:
        profile_to_check = app_state.active_profile
    prof = app_state.data["profiles"][profile_to_check]
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
    return (titles != prof["titles"]
         or texts  != prof["texts"]
         or hks    != prof["hotkeys"])

def save_profile_names():
    # Nur in Edit-Mode und wenn es Eingabefelder gibt
    if not app_state.edit_mode or not app_state.profile_entries:
        return

    # Geplante Umbenennungen einsammeln
    proposed = {}
    for old_name, line_edit in app_state.profile_entries.items():
        if old_name == "SDE":
            continue  # reserviert
        new_name = (line_edit.text() or "").strip()

        if not new_name:
            show_critical_message("Fehler", f"Profilname für '{old_name}' darf nicht leer sein.")
            return
        if new_name == "SDE":
            show_critical_message("Fehler", "Der Profilname 'SDE' ist reserviert.")
            return

        # Duplikate innerhalb der Eingaben verhindern
        if new_name in proposed.values():
            show_critical_message("Fehler", f"Profilname '{new_name}' ist doppelt.")
            return

        # Kollision mit bestehenden Profilen (außer man behält den gleichen Namen)
        if new_name in app_state.data["profiles"] and new_name != old_name:
            show_critical_message("Fehler", f"Profilname '{new_name}' existiert bereits.")
            return

        proposed[old_name] = new_name

    # Umbenennungen anwenden
    new_profiles = {}
    for old_name, prof_data in app_state.data["profiles"].items():
        if old_name == "SDE":
            new_profiles["SDE"] = prof_data
            continue
        target = proposed.get(old_name, old_name)
        new_profiles[target] = prof_data

    # Aktives Profil aktualisieren, falls umbenannt
    if app_state.active_profile in proposed:
        app_state.active_profile = proposed[app_state.active_profile]
    app_state.data["profiles"] = new_profiles
    app_state.data["active_profile"] = app_state.active_profile

    # Persistieren (debounced)
    profiles_to_save = {k: v for k, v in new_profiles.items() if k != "SDE"}
    debounced_saver.schedule_save({
        "profiles": profiles_to_save,
        "active_profile": app_state.active_profile
    })

    update_ui()
    update_profile_buttons()


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
            and data not in (None, "SDE")
        )
        delete_btn.setEnabled(can_delete)

def validate_profile_data(titles, texts, hotkeys):
    """Validate and synchronize profile data arrays"""
    # Ensure all arrays have the same length
    max_length = max(len(titles), len(texts), len(hotkeys))
    
    # Pad shorter arrays with default values
    titles = titles + [f"Titel {i+1}" for i in range(len(titles), max_length)]
    texts = texts + [f"Text {i+1}" for i in range(len(texts), max_length)]
    hotkeys = hotkeys + [f"ctrl+shift+{i+1}" for i in range(len(hotkeys), max_length)]
    
    # Validate hotkey format
    erlaubte_zeichen = set("1234567890befhmpqvxz§'^")
    validated_hotkeys = []
    
    for i, hotkey in enumerate(hotkeys):
        hotkey = hotkey.strip().lower()
        parts = hotkey.split("+")
        
        if (len(parts) != 3 or 
            parts[0] != "ctrl" or 
            parts[1] != "shift" or 
            parts[2] not in erlaubte_zeichen):
            # Generate a valid hotkey
            for char in erlaubte_zeichen:
                test_hotkey = f"ctrl+shift+{char}"
                if test_hotkey not in validated_hotkeys:
                    validated_hotkeys.append(test_hotkey)
                    break
            else:
                validated_hotkeys.append(f"ctrl+shift+{i+1}")
        else:
            validated_hotkeys.append(hotkey)
    
    return titles, texts, validated_hotkeys

def switch_profile(profile_name):
    if profile_name == app_state.active_profile:
        return
    was_visible = win.isVisible()
    
    if app_state.edit_mode and app_state.title_entries and app_state.text_entries and app_state.hotkey_entries:
        current_titles = [entry.text() for entry in app_state.title_entries]
        current_texts = []
        for entry in app_state.text_entries:
            if hasattr(entry, 'toHtml'):
                current_texts.append(entry.toHtml())
            else:
                current_texts.append(entry.text())
        current_hks = [entry.text() for entry in app_state.hotkey_entries]
        
        # Validate and synchronize data before saving
        validated_titles, validated_texts, validated_hotkeys = validate_profile_data(
            current_titles, current_texts, current_hks
        )
        
        stored_data = app_state.data["profiles"][app_state.active_profile]
        
        # VEREINFACHT: Direkter Vergleich ohne unused stored_plain_texts
        has_changes = (validated_titles != stored_data["titles"] or 
                      validated_texts != stored_data["texts"] or 
                      validated_hotkeys != stored_data["hotkeys"])
        
        if has_changes:
            resp = show_question_message(
                "Ungesicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Jetzt speichern?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel
            )
            if resp == QtWidgets.QMessageBox.Cancel:
                return
            if resp == QtWidgets.QMessageBox.Yes:
                # Save validated data
                app_state.data["profiles"][app_state.active_profile]["titles"] = validated_titles
                app_state.data["profiles"][app_state.active_profile]["texts"] = validated_texts
                app_state.data["profiles"][app_state.active_profile]["hotkeys"] = validated_hotkeys
                
                profiles_to_save = {k: v for k, v in app_state.data["profiles"].items() if k != "SDE"}
                debounced_saver.schedule_save({"profiles": profiles_to_save, "active_profile": profile_name})
            elif resp == QtWidgets.QMessageBox.No:
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
    # eigene Profile zählen (ohne SDE)
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
        "hotkeys": [f"ctrl+shift+{i+1}" for i in range(3)]
    }
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
        f"Soll Profil '{profile_name}' wirklich gelöscht werden?"
    )
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

                # 1) Alles leeren
                clipboard.empty()

                # 2) UNICODE-Plaintext setzen
                clipboard.set_text(plain_text_content or "", win32con.CF_UNICODETEXT)

                # 3) CF_HTML bauen (mit CRLF und Byte-Offsets)
                fragment = html_content or ""
                html_body = (
                    "<!DOCTYPE html><html><body>"
                    "<!--StartFragment-->"
                    + fragment +
                    "<!--EndFragment-->"
                    "</body></html>"
                )

                header_tmpl = (
                    "Version:0.9\r\n"
                    "StartHTML:{start_html:010d}\r\n"
                    "EndHTML:{end_html:010d}\r\n"
                    "StartFragment:{start_frag:010d}\r\n"
                    "EndFragment:{end_frag:010d}\r\n"
                )

                # Erst Platzhalter-Header (gleiche Länge) in Bytes berechnen
                placeholder = header_tmpl.format(
                    start_html=0, end_html=0, start_frag=0, end_frag=0
                ).encode("utf-8")
                body_bytes = html_body.encode("utf-8")

                start_html = len(placeholder)
                end_html   = start_html + len(body_bytes)

                # Fragment-Offsets relativ zum Gesamtstring (Header + Body) in BYTES
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

                # 4) Verifizieren (CF_HTML + CF_UNICODETEXT)
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
                                f"Clipboard verify failed (attempt {attempt+1}): HTML={html_ok}, TEXT={txt_ok}"
                            )
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

        # VK-Codes: Ctrl, Shift, Alt, LWin, RWin
        modifiers = (0x11, 0x10, 0x12, 0x5B, 0x5C)

        # Mehrfach versuchen, falls ein Event „hängen“ bleibt
        for _ in range(3):
            for vk in modifiers:
                try:
                    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
                except Exception:
                    pass
            time.sleep(0.01)

        time.sleep(0.05)  # kleine Zusatz-Pause für Stabilität
    except Exception as e:
        logging.warning(f"Error releasing modifier keys (WinAPI): {e}")

def insert_text(index):
    # kleine lokale Helper, damit kein weiterer Code nötig ist
    def _send_ctrl_v():
        user32 = ctypes.windll.user32
        # 0x11 = VK_CONTROL, 0x56 = VK_V
        user32.keybd_event(0x11, 0, 0, 0)   # Ctrl down
        user32.keybd_event(0x56, 0, 0, 0)   # V down
        user32.keybd_event(0x56, 0, 2, 0)   # V up
        user32.keybd_event(0x11, 0, 2, 0)   # Ctrl up

    try:
        txt = app_state.data["profiles"][app_state.active_profile]["texts"][index]
        logging.info(f"Inserting text for index {index}: {txt[:50]}...")
    except IndexError:
        logging.exception(
            f"Kein Text vorhanden für Hotkey-Index {index} im Profil '{app_state.active_profile}'"
        )
        return

    try:
        # 1) Sicherstellen, dass keine Modifier „hängen“
        release_all_modifier_keys()

        # 2) HTML + Plaintext in die Zwischenablage legen
        doc = QtGui.QTextDocument()
        doc.setHtml(txt)
        plain_text = doc.toPlainText()
        success = set_clipboard_html(txt, plain_text)
        if not success:
            logging.warning("Windows clipboard failed, falling back to pyperclip")
            pyperclip.copy(plain_text)
            logging.info(f"Fallback: Set plain text to clipboard: {plain_text[:30]}...")

        # 3) kurze Stabilisierungspause
        time.sleep(0.2)

        # 4) nochmals alle Modifier loslassen (defensiv)
        release_all_modifier_keys()

        # 5) Einfügen via WinAPI (robust, kein 'keyboard' nötig)
        _send_ctrl_v()
        time.sleep(0.05)

        # 6) final: Modifier sicher loslassen
        release_all_modifier_keys()

        logging.info(f"Successfully inserted text for index {index}")

    except Exception as e:
        logging.exception(f"Error in insert_text for index {index}: {e}")
        try:
            # Minimal-Fallback: Nur Plaintext setzen und erneut pasten
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
            f"Kein Text vorhanden für Index {index} im Profil '{app_state.active_profile}'"
        )
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
    
    # Unregister all hotkeys
    for hotkey_id in app_state.registered_hotkey_ids:
        try:
            ctypes.windll.user32.UnregisterHotKey(None, hotkey_id)
        except Exception as e:
            logging.warning(f"Failed to unregister hotkey {hotkey_id}: {e}")
    app_state.registered_hotkey_ids.clear()
    
    # Clear index mapping
    app_state.id_to_index.clear()
    
    # Remove native event filter
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

    # --- Win32 Konstanten & Helper ---
    user32 = ctypes.windll.user32
    WM_HOTKEY = 0x0312
    MOD_ALT     = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT   = 0x0004
    MOD_WIN     = 0x0008

    # VK aus Zeichen (Ziffern + Buchstaben)
    def _vk_from_char(ch: str):
        """
        Liefert den Virtual-Key-Code (VK) für ein einzelnes Zeichen.
        - Ziffern/Buchstaben: schnelle Pfade
        - Sonderzeichen (§ ' ^): via VkKeyScanExW anhand des *aktuellen* Keyboard-Layouts
        """
        ch = (ch or "").strip()
        if not ch:
            return None

        # Quick paths
        if "0" <= ch <= "9":
            return ord(ch)                # VK_0 .. VK_9
        if "a" <= ch <= "z":
            return ord(ch.upper())        # VK_A .. VK_Z

        # Fallback: layoutabhängiges Mapping (auch für § ' ^)
        try:

            user32 = ctypes.windll.user32

            # Signaturen (unverzichtbar nicht, aber sauberer)
            user32.VkKeyScanExW.restype = ctypes.c_short
            # argtypes nicht streng nötig; wir übergeben direkt Python str
            hkl = user32.GetKeyboardLayout(0)
            res = user32.VkKeyScanExW(ch, hkl)   # SHORT: low byte = VK, high byte = Shift/Ctrl/Alt-Flags
            if res == -1:
                return None
            vk = res & 0xFF
            return vk or None
        except Exception:
            return None


    # --- Always cleanup first to prevent memory leaks ---
    cleanup_hotkeys()
    
    # --- Bestehende Registrierungen aufräumen (falls schon welche da sind) ---
    # Wir benutzen eigene Strukturen statt 'registered_hotkey_refs' (keyboard)

    # --- Hotkeys aus aktivem Profil einlesen & validieren ---
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
        # Wichtig: wir bleiben vorerst bei deinem bisherigen Schema 'ctrl+shift+X'
        if (
            len(parts) != 3
            or parts[0] != "ctrl"
            or parts[1] != "shift"
            or parts[2] not in erlaubte_zeichen
        ):
            show_critical_message(
                "Fehler",
                f"Ungültiger Hotkey \"{hotkeys[i]}\" für Eintrag {i+1}.\n"
                f"Erlaubte Zeichen: {''.join(sorted(erlaubte_zeichen))}\n"
                f"Format: ctrl+shift+[zeichen]"
            )
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

    # --- NativeEventFilter einmalig installieren, um WM_HOTKEY zu empfangen ---
    # Wir definieren den Filter lokal und installieren ihn nur einmal.
    if app_state.hotkey_filter_instance is None:
        class _MSG(ctypes.Structure):
            _fields_ = [
                ("hwnd",    wintypes.HWND),
                ("message", wintypes.UINT),
                ("wParam",  wintypes.WPARAM),
                ("lParam",  wintypes.LPARAM),
                ("time",    wintypes.DWORD),
                ("pt",      wintypes.POINT),
            ]

        class _HotkeyFilter(QtCore.QAbstractNativeEventFilter):
            def nativeEventFilter(self, eventType, message):
                # Nur Windows generische Messages interessieren uns
                if eventType == "windows_generic_MSG":
                    # PyQt5 gibt 'message' als sip.voidptr; erst in int-Adresse wandeln:
                    addr = int(message)                 # <- wichtig
                    msg  = _MSG.from_address(addr)      # aus der Adresse eine MSG-Struct machen
                    if msg.message == WM_HOTKEY:
                        try:
                            hotkey_id = int(msg.wParam)
                            idx = app_state.id_to_index.get(hotkey_id)
                            if idx is not None:
                                # Wir sind im Qt-Mainthread – direkt einfügen
                                insert_text(idx)
                        except Exception as e:
                            logging.exception(f"Fehler im WM_HOTKEY-Handler: {e}")
                return False, 0

        app_state.hotkey_filter_instance = _HotkeyFilter()
        app.installNativeEventFilter(app_state.hotkey_filter_instance)
        
        # Add cleanup to application exit
        app.aboutToQuit.connect(cleanup_hotkeys)

    return fehler


#endregion

#region Tray

def create_tray_icon():
    """Erstellt das Tray-Icon mit Menu"""
    try:
        # Cleanup existing
        if app_state.tray:
            try:
                app_state.tray.hide()
                app_state.tray.setParent(None)
                app_state.tray.deleteLater()
            except:
                pass
            app_state.tray = None
        
        # Icon laden mit Fallback
        icon = None
        if os.path.exists(ICON_PATH):
            try:
                icon = QtGui.QIcon(ICON_PATH)
                if icon.isNull():  # Icon konnte nicht geladen werden
                    icon = None
            except:
                logging.warning(f"Konnte Icon nicht laden: {ICON_PATH}")
                icon = None
        
        if not icon:
            # Fallback zu System-Icon
            icon = win.style().standardIcon(QtWidgets.QStyle.SP_ComputerIcon)
        
        # Tray erstellen
        app_state.tray = QSystemTrayIcon(icon, win)
        app_state.tray.setToolTip(f"Aktives Profil: {app_state.active_profile}")
        
        # Menu erstellen
        menu = QMenu()
        
        # Profile hinzufügen
        for prof in app_state.data["profiles"]:
            label = f"✓ {prof}" if prof == app_state.active_profile else f"  {prof}"
            act = QAction(label, win)
            act.triggered.connect(partial(switch_profile, prof))
            menu.addAction(act)
        
        menu.addSeparator()
        
        # Standard-Actions
        act_show = QAction("↑ Öffnen", win)
        act_show.triggered.connect(lambda: (win.show(), win.raise_(), win.activateWindow()))
        menu.addAction(act_show)
        
        act_quit = QAction("✖ Beenden", win)
        act_quit.triggered.connect(lambda: (save_window_position(), app.quit()))
        menu.addAction(act_quit)
        
        # Menu zuweisen
        app_state.tray.setContextMenu(menu)
        
        # Click-Handler für Doppelklick
        app_state.tray.activated.connect(
            lambda reason: (win.show(), win.raise_(), win.activateWindow()) 
            if reason == QSystemTrayIcon.Trigger else None
        )
        
        # Tray anzeigen
        app_state.tray.show()
        logging.info("Tray-Icon erfolgreich erstellt")
        return True
        
    except Exception as e:
        logging.error(f"Tray-Icon Erstellung fehlgeschlagen: {e}")
        app_state.tray = None
        return False

def refresh_tray():
    """Aktualisiert das Tray-Icon"""
    if app_state.tray is not None:
        try:
            app_state.tray.hide()
            app_state.tray.deleteLater()
        except Exception as e:
            logging.warning(f"Failed to cleanup tray: {e}")
        app_state.tray = None
    
    # Neu erstellen
    create_tray_icon()


def minimize_to_tray():
    win.hide()
    if app_state.tray and hasattr(app_state.tray, 'showMessage'):
        app_state.tray.showMessage(
            "QuickPaste", 
            "Anwendung wurde in die Taskleiste minimiert. Hotkeys bleiben aktiv.",
            QSystemTrayIcon.Information, 
            2000
        )

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
                border-radius: 6px;
            }}
        """)
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
                    border-radius: 8px;
                }}
            """)
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

    # eindeutigen Hotkey finden (dein bestehender Code bleibt)
    neuer_hotkey = ""
    for zeichen in erlaubte_zeichen:
        test_hotkey = f"ctrl+shift+{zeichen}"
        if test_hotkey not in belegte_hotkeys:
            neuer_hotkey = test_hotkey
            break
    if not neuer_hotkey:
        neuer_hotkey = "ctrl+shift+"

    # NEU: eindeutigen Titel finden
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
    
    # Sicherheitsprüfungen
    if old_index < 0 or old_index >= len(titles):
        return
    if new_index < 0 or new_index >= len(titles):
        return
        
    # Alle drei Arrays synchron verschieben
    if old_index < len(titles):
        title = titles.pop(old_index)
        titles.insert(new_index, title)
    
    if old_index < len(texts):
        text = texts.pop(old_index)
        texts.insert(new_index, text)
    
    if old_index < len(hotkeys):
        hotkey = hotkeys.pop(old_index)
        hotkeys.insert(new_index, hotkey)
    
    # Sicherstellen dass Arrays synchron sind
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
        if has_field_changes():
            resp = show_question_message(
                "Ungespeicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Willst du sie speichern?"
            )
            if resp == QtWidgets.QMessageBox.Yes:
                save_data()
            else:
                reset_unsaved_changes()
        app_state.edit_mode = False
        update_ui()
        return
    is_sde_only = len(app_state.data["profiles"]) == 1 and "SDE" in app_state.data["profiles"]
    if app_state.active_profile == "SDE" and not is_sde_only:
        show_information_message("Nicht editierbar", "Das SDE-Profil kann nicht bearbeitet werden.")
        return
    app_state.edit_mode = True
    update_ui()
#endregion

#region save_data

def save_data(stay_in_edit_mode=False):
    try:
        if app_state.edit_mode and app_state.profile_entries:
            updated_profiles = {}
            new_active_profile = app_state.active_profile
            for old_name, entry in app_state.profile_entries.items():
                new_name = entry.text().strip()
                if old_name == "SDE" or new_name == "SDE":
                    continue
                if new_name and new_name != old_name:
                    existing_lower = {n.lower() for n in app_state.data["profiles"].keys()
                                    if n not in (old_name, "SDE")}
                    if new_name.lower() in existing_lower:
                        show_critical_message("Fehler", f"Profilname '{new_name}' existiert bereits!")
                        return
                    updated_profiles[new_name] = app_state.data["profiles"].pop(old_name)

                    if app_state.active_profile == old_name:
                        new_active_profile = new_name
                else:
                    updated_profiles[old_name] = app_state.data["profiles"][old_name]
            app_state.data["profiles"] = updated_profiles
            sde = load_sde_profile()
            app_state.data["profiles"]["SDE"] = sde
            available_profiles = {**updated_profiles, "SDE": sde}

            if new_active_profile in available_profiles:
                app_state.active_profile = new_active_profile
            else:
                app_state.active_profile = list(available_profiles.keys())[0]
            if len(updated_profiles) > 11:
                show_critical_message("Limit erreicht", "Maximal 10 Profile erlaubt!")
                return
            app_state.data["active_profile"] = app_state.active_profile
        if app_state.active_profile != "SDE":
            titles_new = [(e.text() or "").strip() for e in app_state.title_entries]
            # leerer Titel verboten
            if any(not t for t in titles_new):
                show_critical_message("Fehler", "Es gibt leere Titel. Bitte fülle alle Titel aus.")
                return
            # Duplikate (case-insensitive) verhindern
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
        QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
    )
    if resp == QtWidgets.QMessageBox.Yes:
        save_data(stay_in_edit_mode=True)
    else:
        reset_unsaved_changes()
    QtCore.QTimer.singleShot(100, action_if_yes)

#endregion

#region zoom


# 3. ZENTRALE ZOOM-FUNKTION
def apply_zoom(z):
    # Begrenzung: 75% .. 125%
    z = max(0.75, min(1.25, z))
    app_state.zoom_level = z

    base = DEFAULT_FONT_SIZE          # z. B. 10/11
    font_size = max(8, int(round(base * z)))

    app = QtWidgets.QApplication.instance()
    if not app:
        return

    # Globale Schrift setzen (alle Widgets übernehmen)
    f = app.font()
    f.setPointSize(font_size)
    app.setFont(f)
    app.setStyleSheet(f"* {{ font-size: {font_size}pt; }}")

    # KEIN update_ui() hier! Nur leichte Auffrischung der sichtbaren Buttons:
    refresh_visible_button_texts()

def refresh_visible_button_texts():
    # Findet alle Text-Buttons und triggert deren lokalen Update-Callback
    app = QtWidgets.QApplication.instance()
    win = next((w for w in app.topLevelWidgets() if isinstance(w, QtWidgets.QMainWindow)), None)
    if not win:
        return
    for btn in win.findChildren(QtWidgets.QPushButton, "qp_text_btn"):
        if hasattr(btn, "_update_text"):
            t = QtCore.QTimer(btn)            # parenten → sicher
            t.setSingleShot(True)
            t.timeout.connect(btn._update_text)
            t.start(0)


def update_widget_sizes():
    """Passt feste Widget-Größen an Zoom an"""
    # Basis-Größen
    base_title_width = 120
    base_hotkey_width = 120
    base_row_height = 40
    base_text_edit_height = 80
    
    # Skalierte Größen
    title_width = int(base_title_width * app_state.zoom_level)
    hotkey_width = int(base_hotkey_width * app_state.zoom_level)
    row_height = int(base_row_height * app_state.zoom_level)
    text_edit_height = int(base_text_edit_height * app_state.zoom_level)
    
    # Titel-Labels/Inputs
    for widget in win.findChildren(QtWidgets.QLabel):
        if widget.objectName() != "drag_handle" and widget.width() == 120:
            widget.setFixedWidth(title_width)
            widget.setFixedHeight(row_height)
    
    for widget in win.findChildren(QtWidgets.QLineEdit):
        if widget.width() == 120:
            widget.setFixedWidth(title_width if "titel" in widget.text().lower() else hotkey_width)
    
    # Text-Edits im Edit-Mode
    for widget in win.findChildren(QtWidgets.QTextEdit):
        widget.setMaximumHeight(text_edit_height)
        widget.setMinimumHeight(int(text_edit_height * 0.75))
    
    # Buttons
    for widget in win.findChildren(QtWidgets.QPushButton):
        if hasattr(widget, 'toolTip') and "Klicken zum Kopieren" in widget.toolTip():
            widget.setFixedHeight(row_height)

# 4. MAUSRAD-EVENT HANDLER
class WheelEventFilter(QtCore.QObject):
    def eventFilter(self, obj, event):
        if event.type() != QtCore.QEvent.Wheel:
            return False

        mods = QtWidgets.QApplication.keyboardModifiers()
        if not (mods & QtCore.Qt.ControlModifier):
            return False

        # Wenn gerade (noch) ein Zoom-Rebuild aussteht oder läuft → nur Timer neu starten, nichts weiteres
        if getattr(app_state, "is_rebuilding", False) or (
            getattr(app_state, "scale_timer", None) and app_state.scale_timer.isActive()
        ):
            if app_state.scale_timer:
                app_state.scale_timer.start(50)
            return True

        # Delta robust ermitteln
        delta = 0
        if hasattr(event, "angleDelta"):
            ad = event.angleDelta()
            if ad:
                delta = ad.y() or ad.x() or 0
        if delta == 0 and hasattr(event, "pixelDelta"):
            pd = event.pixelDelta()
            if pd:
                delta = pd.y() or pd.x() or 0
        if delta == 0:
            return False

        step = 0.05
        apply_zoom(app_state.zoom_level + (step if delta > 0 else -step))
        return True


# 5. TEXT-BUTTON FUNKTIONEN
def calculate_button_text(html_text, button_width):
    """Berechnet dynamisch den Text für Button-Breite"""
    try:
        # HTML zu Plain-Text
        doc = QtGui.QTextDocument()
        doc.setHtml(html_text)
        plain_text = doc.toPlainText().replace('\n', ' ').strip()
        
        if not plain_text:
            return "(Leer)"
        
        # Font-Metriken mit aktuellem Zoom
        font = QFont()
        font.setPointSize(int(app_state.base_font_size * app_state.zoom_level))
        metrics = QFontMetrics(font)
        
        # Verfügbare Breite (minus Padding)
        padding = 30  # Links/Rechts Padding + Border
        usable_width = max(50, button_width - padding)
        
        # Passt der ganze Text?
        if metrics.horizontalAdvance(plain_text) <= usable_width:
            return plain_text
        
        # Text kürzen mit "..."
        ellipsis = "..."
        ellipsis_width = metrics.horizontalAdvance(ellipsis)
        target_width = usable_width - ellipsis_width
        
        # Binäre Suche für optimale Länge
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
            border-radius: 6px;
        }}
        QPushButton:hover {{
            background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};
        }}
    """)
    text_btn.setObjectName("qp_text_btn")   # für das Refresh-Finding

    text_btn.setFixedHeight(int(40 * app_state.zoom_level))
    text_btn.setSizePolicy(QtWidgets.QSizePolicy.Ignored, QtWidgets.QSizePolicy.Fixed)

    text_btn.setToolTip(f"Klicken zum Kopieren • Hotkey: {hks[i] if i < len(hks) else ''}")
    
    # Button-Text aktualisieren
    def update_button_text():
        if text_btn and not sip.isdeleted(text_btn):
            width = text_btn.width()
            if width > 0:
                display_text = calculate_button_text(text_html, width)
                if text_btn.text() != display_text:
                    text_btn.setText(display_text)
    
    # Resize-Event
    original_resize = text_btn.resizeEvent
    def on_resize(event):
        if original_resize:
            original_resize(event)
        for ms in (0, 40):
            t = QtCore.QTimer(text_btn)     # parent = text_btn
            t.setSingleShot(True)
            t.timeout.connect(update_button_text)
            t.start(ms)
    text_btn.resizeEvent = on_resize
    
    # Click-Handler
    text_btn.clicked.connect(partial(copy_text_to_clipboard, i))
    text_btn._update_text = update_button_text

    # Initial update
    for ms in (30, 120):
        t = QtCore.QTimer(text_btn)
        t.setSingleShot(True)
        t.timeout.connect(update_button_text)
        t.start(ms)
    return text_btn

def update_text_buttons():
    """Aktualisiert alle Text-Buttons nach Zoom-Änderung"""
    try:
        for button in win.findChildren(QtWidgets.QPushButton):
            if hasattr(button, 'toolTip') and "Klicken zum Kopieren" in button.toolTip():
                # Trigger resize event für Text-Update
                button.resize(button.size())
    except Exception as e:
        logging.warning(f"Text button update failed: {e}")

# 6. ZOOM-SLIDER (Optional - unten rechts)
def create_zoom_controls():
    """Erstellt Zoom-Controls in der Statusbar"""
    try:
        statusbar = win.statusBar()
        
        # Container Widget
        zoom_widget = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(zoom_widget)
        layout.setContentsMargins(0, 0, 5, 0)
        layout.setSpacing(5)
        
        # Reset-Button
        reset_btn = QtWidgets.QPushButton("↺")
        reset_btn.setFixedSize(20, 20)
        reset_btn.setToolTip("Zoom auf Auto-DPI zurücksetzen")
        reset_btn.clicked.connect(lambda: apply_zoom(detect_optimal_zoom()))
        
        # Zoom-Label
        zoom_label = QtWidgets.QLabel()
        def update_label():
            zoom_label.setText(f"Zoom: {int(app_state.zoom_level * 100)}%")
        update_label()
        
        # Zoom Buttons
        zoom_out = QtWidgets.QPushButton("-")
        zoom_out.setFixedSize(20, 20)
        zoom_out.clicked.connect(lambda: apply_zoom(app_state.zoom_level - 0.05))
        
        zoom_in = QtWidgets.QPushButton("+")
        zoom_in.setFixedSize(20, 20)
        zoom_in.clicked.connect(lambda: apply_zoom(app_state.zoom_level + 0.05))
        
        # Layout
        layout.addWidget(reset_btn)
        layout.addWidget(zoom_out)
        layout.addWidget(zoom_label)
        layout.addWidget(zoom_in)
        
        # In Statusbar einfügen
        statusbar.addPermanentWidget(zoom_widget)
        
        # Update-Funktion registrieren
        zoom_widget.update_label = update_label
        win._zoom_widget = zoom_widget
        
    except Exception as e:
        logging.error(f"Failed to create zoom controls: {e}")

# 7. APP INITIALISIERUNG ANPASSEN
def initialize_application():
    """Initialisiert die Anwendung mit korrektem Scaling"""
    # High-DPI Support
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    
    # Windows App ID
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
        except:
            pass
    
    return app



def install_font_scaling_globally():
    app = QtWidgets.QApplication.instance()
    if app is None:
        return
    if not hasattr(app, "_global_wheel_filter"):
        f = WheelEventFilter()
        app.installEventFilter(f)
        app._global_wheel_filter = f


#endregion

#region Hauptfenster

app = initialize_application()
if sys.platform.startswith("win"):
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
win = QtWidgets.QMainWindow()

def setup_window_scaling():
    # beim Start auf den gespeicherten Zoom gehen – aber erst NACH UI-Aufbau
    QtCore.QTimer.singleShot(0, lambda: apply_zoom(app_state.zoom_level))

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
QtCore.QTimer.singleShot(0, install_font_scaling_globally)




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
                padding: 2px;
            }
            QMenu::item {
                background-color: transparent;
                padding: 6px 20px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #4a90e2;
                color: white;
            }
            QMenu::item:disabled {
                color: #888;
            }
            QMenu::separator {
                height: 1px;
                background-color: #555;
                margin: 2px 10px;
            }
        """)
    global_pos = text_widget.mapToGlobal(pos)
    menu.exec_(global_pos)

def add_hyperlink_to_selection(text_widget, cursor):
    selected_text = cursor.selectedText()
    url, ok = QtWidgets.QInputDialog.getText(
        text_widget, 
        "Add Hyperlink", 
        f"Enter URL for '{selected_text}':",
        text="https://"
    )
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
        "Enter display text:"
    )
    if ok1 and display_text.strip():
        url, ok2 = QtWidgets.QInputDialog.getText(
            text_widget, 
            "Insert Hyperlink", 
            f"Enter URL for '{display_text.strip()}':",
            text="https://"
        )
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
            border-top: 1px solid #666;
        }}
    """)
    entries_margin = 4 if app_state.mini_mode else 8
    entries_layout.setContentsMargins(entries_margin, entries_margin, entries_margin, entries_margin)
    entries_layout.setSpacing(4 if app_state.mini_mode else 6)
    
    
    
    
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
        selector_container.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        selector_layout = QtWidgets.QHBoxLayout(selector_container)
        selector_layout.setContentsMargins(0, 0, 0, 0)
        selector_layout.setSpacing(1 if app_state.mini_mode else 3)

        def scaled(value):
            return max(1, int(value * app_state.zoom_level))

        combo = QtWidgets.QComboBox()
        combo.setEditable(app_state.edit_mode)
        combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        combo.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        combo_height = scaled(24 if app_state.mini_mode else 32)
        combo.setFixedHeight(combo_height)
        combo.setMinimumWidth(scaled(130 if app_state.mini_mode else 200))
        radius = scaled(7 if app_state.mini_mode else 11)
        padding_v = scaled(3 if app_state.mini_mode else 6)
        padding_h = scaled(8 if app_state.mini_mode else 14)
        drop_width = scaled(20 if app_state.mini_mode else 26)
        border_color = "#555" if app_state.dark_mode else "#ccc"


        combo.setStyleSheet(f"""
            QComboBox {{
                background:{bbg};
                color:{fg};
                border: 1px solid {border_color};
                border-radius:{radius}px;
                padding:{padding_v}px {drop_width + padding_v}px {padding_v}px {padding_h}px;
            }}

            /* Rechte Subcontrol explizit runden + gleiche Farbe geben */
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width:{drop_width}px;
                border-left: 1px solid {border_color};
                border-top-right-radius:{radius}px;
                border-bottom-right-radius:{radius}px;
                background:{bbg};
                margin:0; padding:0;
            }}


            QComboBox QAbstractItemView {{
                background:{ebg};
                color:{fg};
                border: 1px solid {border_color};
                selection-background-color:#4a90e2;
                selection-color:white;
            }}
        """)

        # Den Proxy-Stil als Attribut speichern, damit er nicht vom Garbage Collector
        # eingesammelt wird und der benutzerdefinierte Pfeil sichtbar bleibt.
        base_style = combo.style()
        combo._arrow_proxy_style = ComboArrowGlyphStyle(base_style)
        combo.setStyle(combo._arrow_proxy_style)

        selector_layout.addWidget(combo)

        app_state.profile_selector = combo

        delete_btn = None
        if app_state.edit_mode:
            delete_btn = QtWidgets.QPushButton("❌")
            btn_size = scaled(24 if app_state.mini_mode else 28)
            delete_btn.setFixedSize(btn_size, btn_size)
            btn_radius = max(8, btn_size // 2)
            delete_btn.setStyleSheet(
                f"""
                QPushButton {{
                    background:#d32f2f;
                    color:white;
                    border:none;
                    border-radius:{btn_radius}px;
                }}
                QPushButton:hover {{
                    background:#f44336;
                }}
                QPushButton:pressed {{
                    background:#b71c1c;
                }}
                """
            )
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
                    f"color:{fg}; background:transparent; border:none; padding:0px;"
                )
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
            (i for i in range(combo.count()) if combo.itemData(i) == app_state.active_profile),
            -1,
        )
        with QtCore.QSignalBlocker(combo):
            if current_index >= 0:
                combo.setCurrentIndex(current_index)
            elif combo.count() > 0:
                combo.setCurrentIndex(0)
        update_delete_state()





    spacer = QtWidgets.QWidget()
    spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    toolbar.addWidget(spacer)
    if app_state.edit_mode:
        ap = QtWidgets.QPushButton("➕ Profil")
        ap.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border-radius: 5px;
                padding: 6px 16px;
                margin-right: 4px;
            }}
            QPushButton:hover {{
                background:#666;
            }}
            """
        )
        ap.clicked.connect(add_new_profile)
        toolbar.addWidget(ap)

    control_size = 26 if app_state.mini_mode else 30
    control_radius = 12 if app_state.mini_mode else 15
    control_margin = 4 if app_state.mini_mode else 6
    controls = [
        ("🌙" if not app_state.dark_mode else "🌞", toggle_dark_mode, "Dunkelmodus umschalten")
    ]
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
                padding: 0;
            }}
            QPushButton:hover {{
                background:#888;
            }}
            """
        )
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
            padding: 0;
        }}
        QPushButton:hover {{
            background:#888;
        }}
        """
    )
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
                    padding: 1px 4px;
                }}
                QPushButton:hover {{
                    background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};
                }}
                QPushButton QLabel {{
                    color: {fg};
                    font-weight: bold;
                    background: transparent;
                    padding: 0;
                }}
                QPushButton QLabel#miniHotkeyLabel {{
                    font-weight: normal;
                    padding-left: 4px;
                    padding-right: 2px;
                    font-size: 12px;
                    color: {mini_hotkey_color};
                }}

            """)
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
        hl.setStretch(0, 0)   # Titel (nimmt nur so viel wie nötig)
        hl.setStretch(1, 1)   # Text-Button (nimmt Rest)
        hl.setStretch(2, 0)   # Hotkey (nur so breit wie nötig)
        if app_state.edit_mode:
            drag_handle = QtWidgets.QLabel("☰")
            drag_handle.setFixedSize(20, 28)
            drag_handle.setStyleSheet(f"""
                color: {fg}; 
                background: {bbg}; 
                padding: 2px 4px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: 4px;

                text-align: center;
            """)
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
                if not new_title:
                    show_critical_message("Fehler", "Titel darf nicht leer sein!")
                    # auf letzten gespeicherten Wert zurück
                    old = app_state.data["profiles"][app_state.active_profile]["titles"][idx]
                    widget.setText(old)
                    return
                # Duplikate (case-insensitive) gegen alle anderen Title-Feldern prüfen
                current_titles = [e.text().strip().lower() for j, e in enumerate(app_state.title_entries) if j != idx]
                if new_title.lower() in current_titles:
                    show_critical_message("Fehler", f"Titel '{new_title}' wird bereits verwendet!")
                    old = app_state.data["profiles"][app_state.active_profile]["titles"][idx]
                    widget.setText(old)
                    return
                # ok → ins State schreiben
                app_state.data["profiles"][app_state.active_profile]["titles"][idx] = new_title
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
                border-radius: 6px;
            """)
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
                app_state.data["profiles"][app_state.active_profile]["texts"][index] = widget.toHtml()
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
                            f"Format: ctrl+shift+[zeichen]"
                        )
                        widget.setText(hks[idx])
                        return
                    current_hotkeys = [entry.text().strip().lower() for j, entry in enumerate(app_state.hotkey_entries) if j != idx]
                    if hotkey in current_hotkeys:
                        show_critical_message(
                            "Fehler", 
                            f"Hotkey \"{widget.text()}\" wird bereits in diesem Profil verwendet!"
                        )
                        widget.setText(hks[idx])
                        return
                app_state.data["profiles"][app_state.active_profile]["hotkeys"][idx] = widget.text()
            eh.editingFinished.connect(partial(validate_and_set_hotkey, idx=i, widget=eh))
            hl.addWidget(eh)
            app_state.hotkey_entries.append(eh)
            delete_btn = QtWidgets.QPushButton("❌")
            delete_btn.setFixedSize(20, 20)
            delete_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'}; 
                    color: white; 
                    border: 1px solid #b71c1c;
                    border-radius: 3px;
                }}
                QPushButton:hover {{ background: #f44336; }}
            """)
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
                font-family: 'Consolas', 'Monaco', monospace;
            """)
            lh.setAlignment(QtCore.Qt.AlignCenter)
            hl.addWidget(lh)

        entries_layout.addWidget(row)

    if app_state.edit_mode:
        bw = QtWidgets.QWidget()
        bl = QtWidgets.QHBoxLayout(bw)
        bl.setContentsMargins(0,0,0,0)
        bs = QtWidgets.QPushButton("💾 Speichern")
        bs.setStyleSheet("background:green;color:white;")
        bs.clicked.connect(save_data)
        bl.addWidget(bs)
        ba = QtWidgets.QPushButton("➕ Eintrag hinzufügen")
        ba.setStyleSheet(f"background:{bbg}; color:{fg};")
        ba.clicked.connect(add_new_entry)
        bl.addWidget(ba)
        entries_layout.addWidget(bw)



#endregion

#region help 

def show_help_dialog():
    help_text = (
        "QuickPaste Hilfe\n\n"
        "• 🌙/🌞 Dunkelmodus: Wechselt zwischen hell/dunkel.\n\n"
        "• 🔧 Bearbeiten: Titel, Texte und Hotkeys anpassen.\n\n"
        "• ➕ Profil: Neues Textprofil erstellen.\n"
        "• 🖊️ Im Bearbeitungsmodus zwischen Profilen wechseln.\n"
        "• ❌ Löschen: Profil entfernen.\n\n"
        "• ☰ Verschieben: Einträge per Drag & Drop umsortieren.\n"
        "• ❌ Eintrag löschen.\n"
        "• ➕ Eintrag hinzufügen: Fügt einen neuen Eintrag hinzu.\n"
        "• 💾 Speichern: Änderungen sichern.\n\n"
        "Text markieren + Rechtsklick kann ein Hyperlink hinterlegt werden. \n"
        "CTRL+Mausrad kann die Grösse angepasst werden. \n\n"
        "Bei Fragen oder Problemen: nico.wagner@bit.admin.ch"
    )
    show_information_message("QuickPaste Hilfe", help_text)

#endregion

#region darkmode/minimode/messagebox

def apply_dark_mode_to_messagebox(msg):
    if app_state.dark_mode:
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #2e2e2e;
                color: white;
            }
            QMessageBox QLabel {
                color: white !important;
            }
            QMessageBox QPushButton {
                background-color: #444 !important;
                color: white !important;
                border: 1px solid #666;
                border-radius: 5px;
                min-width: 60px;
                min-height: 24px;
                padding: 4px 8px;
                font-weight: normal;
            }
            QMessageBox QPushButton:hover {
                background-color: #666 !important;
                color: white !important;
            }
            QMessageBox QPushButton:pressed {
                background-color: #555 !important;
                color: white !important;
            }
            QMessageBox QPushButton:focus {
                background-color: #4a90e2 !important;
                color: white !important;
                border: 1px solid #5aa3f0;
            }
        """)
        
        for button in msg.findChildren(QtWidgets.QPushButton):
            button.setStyleSheet("""
                QPushButton {
                    background-color: #444;
                    color: white !important;
                    border: 1px solid #666;
                    border-radius: 5px;
                    min-width: 60px;
                    min-height: 24px;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    background-color: #666;
                    color: white !important;
                }
                QPushButton:pressed {
                    background-color: #555;
                    color: white !important;
                }
            """)

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
    """Berechnet dynamisch die Fenstergröße für den Mini-Mode."""
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
QtCore.QTimer.singleShot(0, install_font_scaling_globally)

sys.exit(app.exec_())
