import sys, os, json, re, ctypes, logging
from PyQt5 import QtWidgets, QtGui, QtCore
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
        self.ui_scale = 1.0
        self.ui_scale_source = "auto" 
        self.ui_scale_min = 0.7  
        self.ui_scale_max = 1.3   


# Global application state
app_state = QuickPasteState()

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


#region save window position & scale

def save_window_position():
    try:
        # Geometrie sichern
        geo_bytes = win.saveGeometry()
        geo_hex   = bytes(geo_bytes.toHex()).decode()

        # Bisherige Config laden (falls vorhanden), damit "ui" nicht verloren geht
        cfg_old = {}
        try:
            with open(WINDOW_CONFIG, "r", encoding="utf-8") as f:
                cfg_old = json.load(f)
        except Exception:
            cfg_old = {}

        # Basis-Konfig neu aufbauen
        cfg = {
            "geometry_hex": geo_hex,
            "dark_mode": app_state.dark_mode
        }

        # --- UI-Scale persistieren (pro Screen) ---
        try:
            rec, dpi, scr = get_recommended_scale_and_screen()
        except Exception:
            rec, dpi, scr = 1.0, 96.0, "primary"

        ui_cfg = cfg_old.get("ui", {})
        scales = ui_cfg.get("scales", {})

        scales[scr] = {
            "scale": float(getattr(app_state, "ui_scale", 1.0)),
            "source": getattr(app_state, "ui_scale_source", "auto"),
            "last_dpi": float(dpi)
        }
        ui_cfg["scales"] = scales
        ui_cfg["active_screen"] = scr
        cfg["ui"] = ui_cfg
        # -----------------------------------------

        # Atomar schreiben
        tmp = WINDOW_CONFIG + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4)
        os.replace(tmp, WINDOW_CONFIG)

    except Exception as e:
        logging.exception(f"⚠ Fehler beim Speichern der Fensterposition/UI-Scale: {e}")

def load_window_position():
    try:
        with open(WINDOW_CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)

        # Dark-Mode übernehmen (falls vorhanden)
        if cfg.get("dark_mode") is not None:
            app_state.dark_mode = cfg["dark_mode"]

        # Fenstergeometrie wiederherstellen (falls vorhanden)
        hexstr = cfg.get("geometry_hex")
        if hexstr:
            ba = QByteArray.fromHex(hexstr.encode())
            win.restoreGeometry(ba)

        # --- UI-Scale laden oder automatisch setzen ---
        try:
            rec, dpi, scr = get_recommended_scale_and_screen()
        except Exception:
            rec, dpi, scr = 1.0, 96.0, "primary"

        ui_cfg = cfg.get("ui", {})
        scales = ui_cfg.get("scales", {})
        entry  = scales.get(scr)

        if entry:
            app_state.ui_scale = float(entry.get("scale", rec))
            app_state.ui_scale_source = entry.get("source", "auto")
        else:
            # Nichts gespeichert -> Auto-Wert übernehmen und direkt persistieren
            app_state.ui_scale = float(rec)
            app_state.ui_scale_source = "auto"

            # cfg live erweitern und zurückschreiben
            scales[scr] = {
                "scale": app_state.ui_scale,
                "source": "auto",
                "last_dpi": float(dpi)
            }
            ui_cfg["scales"] = scales
            ui_cfg["active_screen"] = scr
            cfg["ui"] = ui_cfg
            tmp = WINDOW_CONFIG + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=4)
            os.replace(tmp, WINDOW_CONFIG)
        # ---------------------------------------------

        return True

    except (FileNotFoundError, json.JSONDecodeError) as e:
        logging.warning(f"⚠ Fensterposition/UI-Scale konnten nicht geladen werden: {e}")
        # Falls keine Datei vorhanden ist: Auto-Werte setzen
        try:
            rec, _, _ = get_recommended_scale_and_screen()
            app_state.ui_scale = float(rec)
            app_state.ui_scale_source = "auto"
        except Exception:
            app_state.ui_scale = 1.0
            app_state.ui_scale_source = "auto"
        return None

BASE_DPI = 96.0

def get_recommended_scale_and_screen():
    screen = QtGui.QGuiApplication.primaryScreen()
    try:
        dpi = float(screen.logicalDotsPerInch()) if screen else 96.0
        name = screen.name() if screen else "primary"
    except Exception:
        dpi, name = 96.0, "primary"
    # auf 10%-Schritte runden, min/max kappen
    rec = max(0.75, min(3.0, round((dpi / BASE_DPI) * 10) / 10.0))
    return rec, dpi, name

def PX(px_value: float) -> int:
    """Pixelwert proportional zur UI-Skalierung liefern."""
    return int(round(px_value * getattr(app_state, "ui_scale", 1.0)))



#endregion

#region test

def apply_ui_scale(new_scale=None, source="manual"):
    # Delegiere immer auf die sichere Variante
    apply_ui_scale_safe(new_scale, source)



def reset_to_auto_scale():
    """Setzt auf automatische DPI-basierte Skalierung zurück"""
    rec, _, _ = get_recommended_scale_and_screen()
    apply_ui_scale(rec, "auto")

# =============== 3. CTRL+MAUSRAD EVENT FILTER ===============

class WheelZoomFilter(QtCore.QObject):
    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.Wheel:
            modifiers = QtWidgets.QApplication.keyboardModifiers()
            if modifiers == QtCore.Qt.ControlModifier:
                delta = event.angleDelta().y()
                step = 0.05  # 5% Schritte
                if delta > 0:
                    new_scale = app_state.ui_scale + step
                else:
                    new_scale = app_state.ui_scale - step
                # WICHTIG: kompletter UI-Rebuild statt Heuristik
                apply_ui_scale_safe(new_scale, "manual")
                return True
        return False


# =============== 4. TEXT-BUTTON MIT DYNAMISCHEM TEXT ===============

def create_dynamic_text_button(i, texts, hks, ebg, fg):
    """Erstellt Text-Button mit dynamisch angepasstem Text"""
    text_html = texts[i] if i < len(texts) else ""
    
    text_btn = QtWidgets.QPushButton()
    text_btn.setStyleSheet(f"""
        QPushButton {{
            background: {ebg}; 
            color: {fg}; 
            text-align: left; 
            padding: {PX(10)}px {PX(12)}px;
            border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
            border-radius: {PX(6)}px;
            font-size: {PX(13)}px;
        }}
        QPushButton:hover {{
            background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'};
        }}
    """)

    fm = QtGui.QFontMetrics(text_btn.font())
    min_h = fm.height() + PX(16)          # Text-Höhe + Padding
    text_btn.setFixedHeight(max(PX(36), min_h))

    text_btn.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
    text_btn.setToolTip(f"Klicken zum Kopieren • Hotkey: {hks[i] if i < len(hks) else ''}")
    
    # Text dynamisch anpassen


    def update_button_text():
        try:
            # HARTER GUARD: Button schon weg? Dann sofort raus.
            if text_btn is None or sip.isdeleted(text_btn):
                return

            # Sichtbarkeit/Größe nur prüfen, wenn das Objekt lebt
            if not text_btn.isVisible():
                return

            width = text_btn.width()
            if width <= 0:
                return

            # HTML -> Plaintext
            doc = QtGui.QTextDocument()
            doc.setHtml(text_html)
            plain_text = doc.toPlainText().replace('\n', ' ').strip()

            if not plain_text:
                text_btn.setText("(Leer)")
                return

            # Metriken & verfügbare Breite
            font = text_btn.font()
            metrics = QtGui.QFontMetrics(font)
            padding = PX(30)
            available = max(0, width - padding)

            if metrics.horizontalAdvance(plain_text) <= available:
                text_btn.setText(plain_text)
                return

            # Kürzen mit "."
            ellipsis = "."
            target = max(0, available - metrics.horizontalAdvance(ellipsis))
            left, right, best = 0, len(plain_text), 0
            while left <= right:
                mid = (left + right) // 2
                if metrics.horizontalAdvance(plain_text[:mid]) <= target:
                    best = mid
                    left = mid + 1
                else:
                    right = mid - 1

            text_btn.setText((plain_text[:best].rstrip() if best > 0 else "") + ellipsis)

        except RuntimeError:
            # Wird geworfen, wenn das C++-Objekt mitten drin stirbt – einfach abbrechen.
            pass
        except Exception as e:
            logging.warning(f"Button text update failed: {e}")
            # Fallback nur setzen, wenn der Button noch existiert
            try:
                if text_btn is not None and not sip.isdeleted(text_btn):
                    text_btn.setText(".")
            except Exception:
                pass

    # Resize-Event sicher überschreiben
    original_resize = text_btn.resizeEvent
    def on_resize(event):
        try:
            if original_resize:
                original_resize(event)
            QtCore.QTimer.singleShot(0, update_button_text)  # 0ms reicht, Guard schützt
        except RuntimeError:
            pass
    text_btn.resizeEvent = on_resize

    # Initiales Update sicher auslösen
    QtCore.QTimer.singleShot(0, update_button_text)

    return text_btn

# =============== 5. ZOOM-KONTROLLEN IN STATUSBAR ===============

def create_zoom_controls():
    """Erstellt Zoom-Kontrollen in der Statusbar"""
    # Container
    zoom_widget = QtWidgets.QWidget()
    layout = QtWidgets.QHBoxLayout(zoom_widget)
    layout.setContentsMargins(5, 0, 5, 0)
    layout.setSpacing(5)
    
    # Zoom Out Button
    zoom_out_btn = QtWidgets.QPushButton("-")
    zoom_out_btn.setFixedSize(PX(24), PX(24))
    zoom_out_btn.setToolTip("Verkleinern (Strg+Mausrad)")
    zoom_out_btn.clicked.connect(lambda: apply_ui_scale_safe(app_state.ui_scale - 0.05))
    
    # Zoom Label
    zoom_label = QtWidgets.QLabel()
    zoom_label.setAlignment(QtCore.Qt.AlignCenter)
    zoom_label.setMinimumWidth(PX(48))
    
    def update_label():
        percentage = int(app_state.ui_scale * 100)
        zoom_label.setText(f"{percentage}%")
        # Buttons aktivieren/deaktivieren bei Limits
        zoom_out_btn.setEnabled(app_state.ui_scale > app_state.ui_scale_min)
        zoom_in_btn.setEnabled(app_state.ui_scale < app_state.ui_scale_max)
    
    # Zoom In Button
    zoom_in_btn = QtWidgets.QPushButton("+")
    zoom_in_btn.setFixedSize(PX(24), PX(24))
    zoom_in_btn.setToolTip("Vergrößern (Strg+Mausrad)")
    zoom_in_btn.clicked.connect(lambda: apply_ui_scale_safe(app_state.ui_scale + 0.05))
    
    # Reset Button
    reset_btn = QtWidgets.QPushButton("↺")
    reset_btn.setFixedSize(PX(24), PX(24))
    reset_btn.setToolTip("Auf Auto-DPI zurücksetzen")
    reset_btn.clicked.connect(reset_to_auto_scale)
    
    # Layout zusammenbauen
    layout.addWidget(zoom_out_btn)
    layout.addWidget(zoom_label)
    layout.addWidget(zoom_in_btn)
    layout.addWidget(reset_btn)
    
    # Initial update
    update_label()
    
    # Update-Funktion speichern
    zoom_widget.update_display = update_label
    
    # In Statusbar einfügen
    win.statusBar().addPermanentWidget(zoom_widget)
    
    return zoom_widget

def apply_ui_scale_safe(new_scale=None, source="manual"):
    try:
        if new_scale is not None:
            lo = getattr(app_state, "ui_scale_min", 0.9)
            hi = getattr(app_state, "ui_scale_max", 2.0)
            app_state.ui_scale = max(lo, min(hi, float(new_scale)))
            app_state.ui_scale_source = source

        app_font = QtGui.QFont()
        app_font.setPointSizeF(10.0 * app_state.ui_scale)
        app.setFont(app_font)

        def _rebuild():
            try:
                update_ui()
                if hasattr(win, "_zoom_widget") and hasattr(win._zoom_widget, "update_display"):
                    win._zoom_widget.update_display()
                save_window_position()
            except Exception as ee:
                logging.warning(f"UI rebuild after scale failed: {ee}")

        QtCore.QTimer.singleShot(0, _rebuild)

        percentage = int(app_state.ui_scale * 100)
        win.statusBar().showMessage(f"UI-Skalierung: {percentage}%", 2000)

    except Exception as e:
        logging.exception(f"apply_ui_scale_safe failed: {e}")


def update_widget_sizes():
    """Aktualisiert nur die Größen ohne Widgets neu zu erstellen"""
    return

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
    for prof, btn in app_state.profile_buttons.items():
        if prof == app_state.active_profile:
            btn.setStyleSheet("background: lightblue; font-weight: bold;")
        else:
            btn.setStyleSheet("")

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
        act_show = QAction("→ Öffnen", win)
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
                border-radius: {PX(6)}px;
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
                    border-radius: {PX(8)}px;
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

#region Hauptfenster

QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
app = QtWidgets.QApplication(sys.argv)
# UI-Scale laden (falls load_window_position erst später kommt)
try:
    # wenn load_window_position() bereits aufgerufen wurde, ist ui_scale gesetzt
    if getattr(app_state, "ui_scale", None) in (None, 0):
        rec, _, _ = get_recommended_scale_and_screen()
        app_state.ui_scale = rec
except Exception:
    app_state.ui_scale = 1.0

# App-Font skalieren (Points)
app_font = app.font()
base_pt = app_font.pointSizeF() if app_font.pointSizeF() > 0 else 10.0
app_font.setPointSizeF(base_pt * app_state.ui_scale)
app.setFont(app_font)

if sys.platform.startswith("win"):
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
win = QtWidgets.QMainWindow()
win.setWindowTitle("QuickPaste")
win.setMinimumSize(399, 100)
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
main_layout.addWidget(scroll_area)
container = QtWidgets.QWidget()
entries_layout = QtWidgets.QVBoxLayout(container)
entries_layout.setAlignment(QtCore.Qt.AlignTop)
entries_layout.setSpacing(PX(6))
entries_layout.setContentsMargins(PX(8), PX(8), PX(8), PX(8))
scroll_area.setWidget(container)

wheel_filter = WheelZoomFilter()
win.installEventFilter(wheel_filter)

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
    toolbar.clear()
    profs = [p for p in app_state.data["profiles"] if p!="SDE"]
    if not app_state.edit_mode and "SDE" in app_state.data["profiles"]:
        profs.append("SDE")
    for prof in profs:
        frame = QtWidgets.QWidget()
        frame.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        layout = QtWidgets.QHBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        if app_state.edit_mode and prof != "SDE":
            entry = QtWidgets.QLineEdit(prof)
            entry.setFixedWidth(PX(80))
            entry.setStyleSheet(
                f"background:{ebg}; color:{fg}; border-radius:{PX(5)}px; padding:{PX(4)}px;"
            )
            app_state.profile_entries[prof] = entry
            layout.addWidget(entry)
            switch_btn = QtWidgets.QPushButton("🖊️" if prof == app_state.active_profile else "🖊️")
            switch_btn.setFixedWidth(PX(28))
            switch_btn.setStyleSheet(
                f"""
                background:{'#4a90e2' if prof == app_state.active_profile else bbg}; 
                color:{fg}; 
                border-radius:{PX(12)}px;
                font-size: {PX(14)}px;
                """
            )
            switch_btn.setToolTip(f"Zu Profil '{prof}' wechseln")
            switch_btn.clicked.connect(partial(switch_profile, prof))
            layout.addWidget(switch_btn)
            delete_btn = QtWidgets.QPushButton("❌")
            delete_btn.setFixedWidth(PX(28))
            delete_btn.setStyleSheet(
                f"background:{bbg}; color:{fg}; border-radius:{PX(12)}px;"
            )
            delete_btn.clicked.connect(partial(delete_profile, prof))
            layout.addWidget(delete_btn)
        else:
            btn = QtWidgets.QPushButton(prof)
            btn.setStyleSheet(
                f"""
                QPushButton {{
                    background:{bbg}; color:{fg};
                    font-weight:{'bold' if prof==app_state.active_profile else 'normal'};
                    border-radius: {PX(5)}px;
                    padding: {PX(6)}px {PX(16)}px;
                    margin-right: 4px;
                }}
                QPushButton:hover {{
                    background:#666;
                }}
                """
            )
            btn.clicked.connect(partial(switch_profile, prof))
            layout.addWidget(btn)
        toolbar.addWidget(frame)
    spacer = QtWidgets.QWidget()
    spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    toolbar.addWidget(spacer)
    if app_state.edit_mode:
        ap = QtWidgets.QPushButton("➕ Profil")
        ap.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border-radius: {PX(5)}px;
                padding: {PX(6)}px {PX(16)}px;
                margin-right: {PX(4)}px;
            }}
            QPushButton:hover {{
                background:#666;
            }}
            """
        )
        ap.clicked.connect(add_new_profile)
        toolbar.addWidget(ap)
    for text, func, tooltip in [
        ("🌙" if not app_state.dark_mode else "🌞", toggle_dark_mode, "Dunkelmodus umschalten"),
        ("🔧", toggle_edit_mode, "Bearbeitungsmodus umschalten")
    ]:
        b = QtWidgets.QPushButton(text)
        b.setToolTip(tooltip)
        b.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border-radius: {PX(15)}px;
                min-width: {PX(30)}px; min-height: {PX(30)}px;
                font-size: {PX(18)}px;
                margin-left: {PX(6)}px;
                border: none;
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
            border-radius: {PX(15)}px;
            min-width: {PX(30)}px; min-height: {PX(30)}px;
            font-size: {PX(18)}px;
            margin-left: {PX(6)}px;
            border: none;
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
    max_h = PX(120)
    for i, title in enumerate(titles):
        if app_state.edit_mode:
            row = DragDropWidget(i)
        else:
            row = QtWidgets.QWidget()
        hl  = QtWidgets.QHBoxLayout(row)
        hl.setContentsMargins(PX(8), PX(4), PX(8), PX(4))
        hl.setSpacing(12)
        if app_state.edit_mode:
            drag_handle = QtWidgets.QLabel("☰")
            drag_handle.setFixedSize(PX(20), PX(28))
            drag_handle.setStyleSheet(f"""
                color: {fg}; 
                background: {bbg}; 
                padding: 2px 4px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: {PX(4)}px;
                font-size: {PX(14)}px;
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
            et.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if app_state.dark_mode else '#ccc'}; border-radius: {PX(6)}px; padding: {PX(8)}px;")
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
                padding: {PX(10)}px {PX(12)}px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: {PX(6)}px;
                font-size: {PX(13)}px;
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
            text_btn = create_dynamic_text_button(i, texts, hks, ebg, fg)
            hl.addWidget(text_btn, 1)
        if app_state.edit_mode:
            eh = QtWidgets.QLineEdit(hks[i])
            eh.setFixedWidth(max_h)
            eh.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if app_state.dark_mode else '#ccc'}; border-radius: {PX(6)}px; padding: {PX(8)}px;")
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
            delete_btn.setFixedSize(PX(20), PX(20))
            delete_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {'#4a4a4a' if app_state.dark_mode else '#f0f0f0'}; 
                    color: white; 
                    border: 1px solid #b71c1c;
                    border-radius: {PX(3)}px;
                    font-size: {PX(10)}px;
                }}
                QPushButton:hover {{ background: #f44336; }}
            """)
            delete_btn.clicked.connect(lambda _, j=i: delete_entry(j))
            delete_btn.setToolTip("Eintrag löschen")
            hl.addWidget(delete_btn)
        else:
            hk_text = hks[i] if i < len(hks) else ""
            lh = QtWidgets.QLabel(hk_text)
            lh.setAlignment(QtCore.Qt.AlignCenter)
            lh.setStyleSheet(f"""
                color: {fg};
                background: {ebg};
                padding: {PX(6)}px {PX(10)}px;
                border: 1px solid {'#555' if app_state.dark_mode else '#ccc'};
                border-radius: {PX(6)}px;
                font-size: {PX(13)}px;
                font-family: 'Consolas','Monaco',monospace;
            """)
            # Breite dynamisch aus Text + Padding
            fm = QtGui.QFontMetrics(lh.font())
            pad_x = PX(20)  # links+rechts (entspricht padding oben)
            text_w = fm.horizontalAdvance(hk_text)
            want_w = max(PX(110), text_w + pad_x)  # min 60px, sonst Textbreite
            lh.setMinimumWidth(want_w)
            lh.setMaximumWidth(want_w)  # exakt passend, keine wackelnden Layouts

            # Höhe passend zur Font
            want_h = max(PX(36), fm.height() + PX(12))
            lh.setFixedHeight(want_h)

            lh.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
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
        "Bei Fragen oder Problemen: nico.wagner@bit.admin.ch"
    )
    show_information_message("QuickPaste Hilfe", help_text)

#endregion

#region darkmode

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

#endregion

app.aboutToQuit.connect(lambda: (debounced_saver.timer.stop(), debounced_saver._save()))
load_window_position()
win._zoom_widget = create_zoom_controls()

update_ui()
create_tray_icon()
register_hotkeys()
win.show()
sys.exit(app.exec_())
