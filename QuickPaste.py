import sys, os, json, re, ctypes, logging
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QByteArray, QMimeData
import time
import pyperclip
import keyboard
from PyQt5.QtWidgets import QSystemTrayIcon, QAction, QMenu

from pystray import Menu, Icon, MenuItem
from PIL import Image
import threading

APPDATA_PATH = os.path.join(os.environ["APPDATA"], "QuickPaste")
os.makedirs(APPDATA_PATH, exist_ok=True)
CONFIG_FILE = os.path.join(APPDATA_PATH, "config.json")
WINDOW_CONFIG = os.path.join(APPDATA_PATH, "window_config.json")
SDE_FILE = os.path.join(APPDATA_PATH, "sde.json")
LOG_FILE = os.path.join(APPDATA_PATH, "qp.log")
logging.basicConfig(filename=LOG_FILE, filemode="a", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8")
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
ICON_PATH = os.path.join(BASE_DIR, "assets", "H.ico")
dark_mode = False
registered_hotkey_refs = []
unsaved_changes = False
dragged_index = None
current_target = None
profile_buttons = {}     # Profil → QPushButton im Read‑Mode
profile_lineedits = {}   # Profil → QLineEdit im Edit‑Mode
edit_mode = False
text_entries = []
title_entries = []
hotkey_entries = []
tray = None
registered_hotkey_refs = []

#region window position 

def save_window_position():
    """
    Speichert Geometrie und Dark‑Mode in WINDOW_CONFIG (JSON).
    """
    try:
        # Qt speichert Geometrie als QByteArray
        geo_bytes = win.saveGeometry()
        geo_hex   = bytes(geo_bytes.toHex()).decode()
        cfg = {"geometry_hex": geo_hex, "dark_mode": dark_mode}
        # Atomar schreiben
        tmp = WINDOW_CONFIG + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
        os.replace(tmp, WINDOW_CONFIG)
    except Exception as e:
        logging.exception(f"⚠ Fehler beim Speichern der Fensterposition: {e}")

def load_window_position():
    """
    Liest WINDOW_CONFIG und stellt Geometrie & Dark‑Mode wieder her.
    Liefert None oder True.
    """
    global dark_mode
    try:
        with open(WINDOW_CONFIG, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        # Dark‑Mode wiederherstellen
        if cfg.get("dark_mode") is not None:
            dark_mode = cfg["dark_mode"]
        # Geometrie wiederherstellen
        hexstr = cfg.get("geometry_hex")
        if hexstr:
            ba = QByteArray.fromHex(hexstr.encode())
            win.restoreGeometry(ba)
        return True
    except (FileNotFoundError, json.JSONDecodeError) as e:
        logging.warning(f"⚠ Fensterposition konnte nicht geladen werden: {e}")
        return None

#endregion

def refresh_tray():
    global tray
    if tray is not None:
        try:
            tray.hide()
            tray.deleteLater()  # Properly delete the old tray icon
        except:
            pass
        tray = None
    create_tray_icon()


#region data

def load_sde_profile():
    """Lädt sde.json oder liefert Standard‑SDE zurück."""
    try:
        with open(SDE_FILE, "r", encoding="utf-8") as f:
            sde = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        logging.warning("⚠ Konnte sde.json nicht laden. Setze Standard‑SDE.")
        sde = {}

    # Wenn nichts drinsteht, fülle mit Defaults
    if not sde.get("titles") and not sde.get("texts") and not sde.get("hotkeys"):
        sde = {
            "titles": ["Standard Titel 1", "Standard Titel 2", "Standard Titel 3"],
            "texts":  ["Standard Text 1",  "Standard Text 2",  "Standard Text 3"],
            "hotkeys":["ctrl+shift+1",    "ctrl+shift+2",    "ctrl+shift+3"]
        }
    return sde

def load_data():
    """Lädt config.json, stellt Struktur sicher und hängt SDE an."""
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        # Sicherstellen, dass profiles ein Dict ist
        if not isinstance(loaded.get("profiles"), dict):
            loaded["profiles"] = {}
        # Für jedes Profil die drei Listen garantieren
        for prof, vals in loaded["profiles"].items():
            vals.setdefault("titles", [])
            vals.setdefault("texts",  [])
            vals.setdefault("hotkeys", [])
        # SDE‑Profil anhängen (überschreibt vorhandenes SDE)
        loaded["profiles"]["SDE"] = load_sde_profile()
        # aktives Profil prüfen
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
        # Fallback‑Defaults, falls Datei fehlt oder ungültig ist
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
data = load_data()
active_profile = data.get("active_profile")

#endregion 

#region profiles

def has_field_changes(profile_to_check=None):
    """Gibt True, wenn sich die sichtbaren Edit‑Felder von den Daten unterscheiden."""
    if profile_to_check is None:
        profile_to_check = active_profile
    prof = data["profiles"][profile_to_check]
    titles, texts, hks = [], [], []
    # Sammle Edit‑Zeilen aus entries_layout
    for i in range(entries_layout.count()):
        item = entries_layout.itemAt(i)
        if item is None:
            continue
        row = item.widget()
        if edit_mode and isinstance(row, QtWidgets.QWidget):
            # Find QLineEdit widgets for titles and hotkeys
            line_edits = row.findChildren(QtWidgets.QLineEdit)
            # Find QTextEdit widgets for texts
            text_edits = row.findChildren(QtWidgets.QTextEdit)
            
            if len(line_edits) >= 2 and len(text_edits) >= 1:
                titles.append(line_edits[0].text())
                texts.append(text_edits[0].toHtml())
                hks.append(line_edits[1].text())
    return (titles != prof["titles"]
         or texts  != prof["texts"]
         or hks    != prof["hotkeys"])

def save_profile_names():
    """Liest alle QLineEdit in profile_lineedits, benennt data['profiles'] um."""
    global data, active_profile
    if not edit_mode:
        return
    new_profiles = {}
    for old, le in profile_lineedits.items():
        new = le.text().strip()
        if new and new not in new_profiles:
            new_profiles[new] = data["profiles"].pop(old)
        else:
            show_critical_message("Fehler", f"Profilname '{new}' ungültig oder doppelt!")
            return
    data["profiles"] = new_profiles
    if active_profile not in data["profiles"]:
        active_profile = next(iter(data["profiles"]))
    data["active_profile"] = active_profile
    save_data()
    update_ui()

def update_profile_buttons():
    """Hebt in profile_buttons den aktiven Profil‑Button hervor."""
    for prof, btn in profile_buttons.items():
        if prof == active_profile:
            btn.setStyleSheet("background: lightblue; font-weight: bold;")
        else:
            btn.setStyleSheet("")

def switch_profile(profile_name):
    """Wechselt das Profil, fragt bei ungespeicherten Änderungen."""
    global active_profile
    
    # Don't check for changes if we're switching to the same profile
    if profile_name == active_profile:
        return
    
    # If we're in edit mode, check for changes in the current profile BEFORE switching
    if edit_mode and title_entries and text_entries and hotkey_entries:
        # Use the existing entry arrays that are maintained by update_ui()
        current_titles = [entry.text() for entry in title_entries]
        
        # For text comparison, use plain text to avoid HTML formatting differences
        current_texts = []
        for entry in text_entries:
            if hasattr(entry, 'toPlainText'):
                current_texts.append(entry.toPlainText())
            else:
                current_texts.append(entry.text())
        
        current_hks = [entry.text() for entry in hotkey_entries]
        
        # Compare with stored data for current profile
        stored_data = data["profiles"][active_profile]
        
        # Convert stored HTML texts to plain text for comparison
        stored_plain_texts = []
        for stored_text in stored_data["texts"]:
            if '<' in stored_text and '>' in stored_text:  # HTML content
                doc = QtGui.QTextDocument()
                doc.setHtml(stored_text)
                stored_plain_texts.append(doc.toPlainText())
            else:
                stored_plain_texts.append(stored_text)
        
        has_changes = (current_titles != stored_data["titles"] or 
                      current_texts != stored_plain_texts or 
                      current_hks != stored_data["hotkeys"])
        
        if has_changes:
            resp = show_question_message(
                "Ungesicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Jetzt speichern?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel
            )
            if resp == QtWidgets.QMessageBox.Cancel:
                return
            if resp == QtWidgets.QMessageBox.Yes:
                # Save current changes to current profile before switching
                data["profiles"][active_profile]["titles"] = current_titles
                data["profiles"][active_profile]["texts"] = [entry.toHtml() if hasattr(entry, 'toHtml') else entry.text() for entry in text_entries]
                data["profiles"][active_profile]["hotkeys"] = current_hks
                # Save to file
                profiles_to_save = {k: v for k, v in data["profiles"].items() if k != "SDE"}
                with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                    json.dump({"profiles": profiles_to_save, "active_profile": profile_name}, f, indent=4)
            elif resp == QtWidgets.QMessageBox.No:
                # User clicked "No" - reload original data from file to discard changes
                # This ensures we revert to the saved state before switching profiles
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        file_data = json.load(f)
                    if active_profile in file_data.get("profiles", {}):
                        # Restore the original data for current profile
                        data["profiles"][active_profile] = file_data["profiles"][active_profile].copy()
                except (FileNotFoundError, json.JSONDecodeError, KeyError):
                    # If file read fails, keep current data as fallback
                    pass
    
    if profile_name not in data["profiles"]:
        show_critical_message("Fehler", f"Profil '{profile_name}' existiert nicht!")
        return
    
    # Switch to the new profile
    active_profile = profile_name
    data["active_profile"] = profile_name
    
    # Save the active profile change to file
    profiles_to_save = {k: v for k, v in data["profiles"].items() if k != "SDE"}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump({"profiles": profiles_to_save, "active_profile": profile_name}, f, indent=4)
    
    # Refresh the UI to show the correct profile's data
    update_profile_buttons()
    update_ui()

def add_new_profile():
    """Legt ein neues Profil mit drei Default‑Einträgen an."""
    global active_profile
    if len(data["profiles"]) >= 11:
        show_critical_message("Limit erreicht", "Maximal 10 Profile erlaubt!")
        return
    base, cnt = "Profil", 1
    while f"{base} {cnt}" in data["profiles"]:
        cnt += 1
    name = f"{base} {cnt}"
    data["profiles"][name] = {
        "titles":  [f"Titel {i+1}" for i in range(3)],
        "texts":   [f"Text {i+1}"  for i in range(3)],
        "hotkeys": [f"ctrl+shift+{i+1}" for i in range(3)]
    }
    active_profile = name
    data["active_profile"] = name
    update_ui()

def delete_profile(profile_name):
    """Löscht ein Profil mit Bestätigung (SDE ausgeschlossen)."""
    global active_profile
    if profile_name == "SDE":
        show_critical_message("Fehler", "Das SDE‑Profil kann nicht gelöscht werden.")
        return
    if len(data["profiles"]) <= 1:
        show_critical_message("Fehler", "Mindestens ein Profil muss bestehen bleiben!")
        return
    resp = show_question_message(
        "Profil löschen",
        f"Soll Profil '{profile_name}' wirklich gelöscht werden?"
    )
    if resp != QtWidgets.QMessageBox.Yes:
        return
    del data["profiles"][profile_name]
    if active_profile == profile_name:
        active_profile = next(iter(data["profiles"]))
        data["active_profile"] = active_profile
    update_ui()
    save_data()

def make_profile_switcher(profile_name):
    """Für Tray‑Menu: Callback‑Wrapper für switch_profile."""
    return lambda icon, item: switch_profile(profile_name)

#endregion 

#region insert text / hotkeys

def insert_text(index):
    """
    Kopiert den gespeicherten Text in die Zwischenablage
    und führt Strg+V aus.
    """
    try:
        txt = data["profiles"][active_profile]["texts"][index]
    except IndexError:
        logging.exception(
            f"Kein Text vorhanden für Hotkey-Index {index} im Profil '{active_profile}'"
        )
        return
    # For rich text, copy as HTML to preserve formatting and hyperlinks
    if '<' in txt and '>' in txt:  # Simple check for HTML content
        # Use QClipboard to set both HTML and plain text
        clipboard = QtWidgets.QApplication.clipboard()
        mime_data = QtCore.QMimeData()
        mime_data.setHtml(txt)
        mime_data.setText(txt)  # Fallback plain text
        clipboard.setMimeData(mime_data)
    else:
        pyperclip.copy(txt)
    time.sleep(0.1)
    keyboard.send("ctrl+v")

def copy_text_to_clipboard(index):
    """
    Kopiert den gespeicherten Text nur in die Zwischenablage ohne Strg+V.
    Wird für Button-Klicks verwendet.
    """
    try:
        txt = data["profiles"][active_profile]["texts"][index]
    except IndexError:
        logging.exception(
            f"Kein Text vorhanden für Index {index} im Profil '{active_profile}'"
        )
        return
    
    # For rich text, copy as HTML to preserve formatting and hyperlinks
    if '<' in txt and '>' in txt:  # Simple check for HTML content
        # Use QClipboard to set both HTML and plain text
        clipboard = QtWidgets.QApplication.clipboard()
        mime_data = QtCore.QMimeData()
        mime_data.setHtml(txt)
        mime_data.setText(txt)  # Fallback plain text
        clipboard.setMimeData(mime_data)
    else:
        pyperclip.copy(txt)
    
    # Optional: Show a brief visual feedback
    if hasattr(win, 'statusBar'):
        win.statusBar().showMessage("Text in Zwischenablage kopiert!", 2000)

def register_hotkeys():
    """
    Entfernt alte Hotkeys und registriert neue für
    data['profiles'][active_profile]['hotkeys'].
    Nutzt QMessageBox für Fehleranzeigen.
    """
    global registered_hotkey_refs

    # alte Hotkeys löschen
    for ref in registered_hotkey_refs:
        try:
            keyboard.remove_hotkey(ref)
        except Exception:
            pass
    registered_hotkey_refs.clear()

    erlaubte_zeichen = set("1234567890befhmpqvxz§'^")
    belegte = set()
    fehler = False

    hotkeys = data["profiles"].setdefault(active_profile, {}).setdefault("hotkeys", [])

    def hotkey_handler(idx):
        # Falls Pfeiltasten gedrückt, nicht einfügen
        pressed = keyboard._pressed_events.keys()
        if any(k in {72,80,75,77} for k in pressed):
            logging.warning("⚠ Windows-Funktion erkannt, kein Text eingefügt.")
            return
        insert_text(idx)
        keyboard.release("ctrl")
        keyboard.release("shift")

    for i, hot in enumerate(hotkeys):
        hot = hot.strip().lower()
        if not hot:
            continue
        parts = hot.split("+")
        if (
            len(parts) != 3
            or parts[0] != "ctrl"
            or parts[1] != "shift"
            or parts[2] not in erlaubte_zeichen
        ):
            show_critical_message(
                "Fehler",
                f"Ungültiger Hotkey \"{hotkeys[i]}\" für Eintrag {i+1}.\n"
                f"Erlaubte Zeichen: {''.join(sorted(erlaubte_zeichen))}"
            )
            fehler = True
            continue
        if hot in belegte:
            show_critical_message("Fehler", f"Hotkey \"{hotkeys[i]}\" wird bereits verwendet!")
            fehler = True
            continue
        belegte.add(hot)
        if i >= len(data["profiles"][active_profile]["texts"]):
            logging.warning(
                f"⚠ Hotkey '{hot}' zeigt auf Eintrag {i+1}, aber dieser existiert nicht."
            )
            continue

        # registrieren
        try:
            ref = keyboard.add_hotkey(hot, lambda idx=i: hotkey_handler(idx), suppress=True)
            registered_hotkey_refs.append(ref)
        except Exception as e:
            logging.exception(f"Fehler beim Registrieren des Hotkeys '{hot}': {e}")
            fehler = True

    return fehler

#endregion

#region Tray

def create_tray_icon():
    global tray
    # Clean up any existing tray icon first
    if tray:
        try:
            tray.hide()
            tray.deleteLater()
        except:
            pass
        tray = None

    tray = QSystemTrayIcon(QtGui.QIcon(ICON_PATH), win)
    menu = QMenu()

    # Profil‑Einträge im Menü
    for prof in data["profiles"]:
        label = f"✓ {prof}" if prof == active_profile else f"  {prof}"
        act = QAction(label, win)
        act.triggered.connect(lambda _, p=prof: switch_profile(p))
        menu.addAction(act)

    menu.addSeparator()

    # Öffnen
    act_show = QAction("↑ Öffnen", win)
    act_show.triggered.connect(lambda: (win.show(), win.raise_(), win.activateWindow()))
    menu.addAction(act_show)

    # Beenden
    act_quit = QAction("✖ Beenden", win)
    act_quit.triggered.connect(lambda: (save_window_position(), app.quit()))
    menu.addAction(act_quit)

    tray.setContextMenu(menu)
    tray.activated.connect(
        lambda reason:
            (win.show(), win.raise_(), win.activateWindow()) 
            if reason == QSystemTrayIcon.Trigger 
            else None
    )
    tray.show()


def minimize_to_tray():
    """Hide the window to system tray"""
    win.hide()
    # Only show notification if tray exists and has the showMessage method
    if tray and hasattr(tray, 'showMessage'):
        tray.showMessage(
            "QuickPaste", 
            "Anwendung wurde in die Taskleiste minimiert. Hotkeys bleiben aktiv.",
            QSystemTrayIcon.Information, 
            2000
        )

#endregion

#region add/del/move/drag Entry

def start_drag(event, index, widget):
    """Start drag operation"""
    global dragged_index, dark_mode
    dragged_index = index
    
    # Get the parent row widget to highlight it during drag
    row_widget = widget.parent()
    if hasattr(row_widget, 'highlight_drop_zone'):
        # Store original style and dim the dragged item
        original_style = row_widget.styleSheet()
        row_widget.setStyleSheet(f"""
            QWidget {{
                background-color: {'#1a1a1a' if dark_mode else '#f0f0f0'};
                opacity: 0.6;
                border: 1px dashed {'#666' if dark_mode else '#999'};
                border-radius: 6px;
            }}
        """)
    
    # Create drag object
    drag = QtGui.QDrag(widget)
    mime_data = QtCore.QMimeData()
    mime_data.setText(str(index))
    drag.setMimeData(mime_data)
    
    # Execute drag
    result = drag.exec_(QtCore.Qt.MoveAction)
    
    # Restore original appearance after drag is complete
    if hasattr(row_widget, 'highlight_drop_zone'):
        row_widget.setStyleSheet("")
        # Clear any remaining highlights on all widgets
        clear_all_highlights()

def clear_all_highlights():
    """Clear highlight styling from all drag-drop widgets"""
    try:
        # Find all DragDropWidget instances and clear their highlights
        for i in range(entries_layout.count()):
            item = entries_layout.itemAt(i)
            if item and item.widget():
                widget = item.widget()
                if isinstance(widget, DragDropWidget) and hasattr(widget, 'highlight_drop_zone'):
                    widget.highlight_drop_zone(False)
    except:
        # Ignore errors if layout is being modified
        pass

class DragDropWidget(QtWidgets.QWidget):
    """Custom widget that handles drag and drop"""
    def __init__(self, index, parent=None):
        super().__init__(parent)
        self.drag_index = index
        self.setAcceptDrops(True)
        self.original_style = ""
        self.is_highlighted = False
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            # Highlight this widget as a drop target
            self.highlight_drop_zone(True)
            event.acceptProposedAction()
    
    def dragLeaveEvent(self, event):
        # Remove highlight when drag leaves this widget
        self.highlight_drop_zone(False)
        super().dragLeaveEvent(event)
    
    def dropEvent(self, event):
        global dragged_index
        # Remove highlight after drop
        self.highlight_drop_zone(False)
        
        if event.mimeData().hasText():
            source_index = int(event.mimeData().text())
            target_index = self.drag_index
            
            if source_index != target_index:
                move_entry_to(source_index, target_index)
            
            event.acceptProposedAction()
    
    def highlight_drop_zone(self, highlight):
        """Add or remove visual highlighting for drop zones"""
        global dark_mode
        if highlight and not self.is_highlighted:
            # Store original style and apply highlight
            self.original_style = self.styleSheet()
            highlight_color = "#4a90e2" if dark_mode else "#87ceeb"
            border_color = "#5aa3f0" if dark_mode else "#4682b4"
            self.setStyleSheet(f"""
                QWidget {{
                    background-color: {highlight_color};
                    border: 2px solid {border_color};
                    border-radius: 8px;
                }}
            """)
            self.is_highlighted = True
        elif not highlight and self.is_highlighted:
            # Restore original style
            self.setStyleSheet(self.original_style)
            self.is_highlighted = False

def move_entry(index, direction):
    """
    Verschiebt den Eintrag (Titel und Text) nach oben oder unten.
    Die Hotkey-Reihenfolge bleibt unverändert.
    """
    titles = data["profiles"][active_profile]["titles"]
    texts = data["profiles"][active_profile]["texts"]
    if direction == "up" and index > 0:
        titles[index], titles[index-1] = titles[index-1], titles[index]
        texts[index], texts[index-1] = texts[index-1], texts[index]
    elif direction == "down" and index < len(titles) - 1:
        titles[index], titles[index+1] = titles[index+1], titles[index]
        texts[index], texts[index+1] = texts[index+1], texts[index]
    update_ui() 

def add_new_entry():
    global data
    if "titles" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["titles"] = []
    if "texts" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["texts"] = []
    if "hotkeys" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["hotkeys"] = []
    erlaubte_zeichen = "1234567890befhmpqvxz§'^"
    belegte_hotkeys = set(data["profiles"][active_profile]["hotkeys"])
    neuer_hotkey = ""
    for zeichen in erlaubte_zeichen:
        test_hotkey = f"ctrl+shift+{zeichen}"
        if test_hotkey not in belegte_hotkeys:
            neuer_hotkey = test_hotkey
            break
    if not neuer_hotkey:
        neuer_hotkey = "ctrl+shift+"  
    data["profiles"][active_profile]["titles"].append("Neuer Eintrag")
    data["profiles"][active_profile]["texts"].append("Neuer Text")
    data["profiles"][active_profile]["hotkeys"].append(neuer_hotkey)
    update_ui()

def delete_entry(index):
    global data
    if "titles" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["titles"] = []
    if "texts" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["texts"] = []
    if "hotkeys" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["hotkeys"] = []
    if index < 0 or index >= len(data["profiles"][active_profile]["titles"]):
        show_critical_message("Fehler", "Ungültiger Eintrag zum Löschen ausgewählt!")
        return
    del data["profiles"][active_profile]["titles"][index]
    del data["profiles"][active_profile]["texts"][index]
    del data["profiles"][active_profile]["hotkeys"][index]
    update_ui() 
    register_hotkeys() 

def move_entry_to(old_index, new_index):
    titles = data["profiles"][active_profile]["titles"]
    texts = data["profiles"][active_profile]["texts"]
    title = titles.pop(old_index)
    text = texts.pop(old_index)
    titles.insert(new_index, title)
    texts.insert(new_index, text)
    update_ui()

def on_drag_start(event, index):
    global dragged_index
    dragged_index = index
    event.widget.configure(bg="gray")

def on_drag_motion(event):
    global current_target
    target_widget = event.widget.winfo_containing(event.x_root, event.y_root)
    while target_widget is not None and not hasattr(target_widget, "drag_index"):
        target_widget = target_widget.master
    if current_target and current_target != target_widget:
        current_target.configure(bg=current_target.orig_bg)
    if target_widget:
        target_widget.configure(bg="gray")
        current_target = target_widget

def on_drag_release(event):
    global dragged_index, current_target
    if current_target:
        current_target.configure(bg=current_target.orig_bg)
        current_target = None
    if dragged_index is None:
        return
    target_widget = event.widget.winfo_containing(event.x_root, event.y_root)
    while target_widget is not None and not hasattr(target_widget, "drag_index"):
        target_widget = target_widget.master
    if target_widget is not None:
        target_index = target_widget.drag_index
        if target_index is not None and target_index != dragged_index:
            swap_entries(dragged_index, target_index)
    dragged_index = None
    update_ui()

def swap_entries(i, j):
    titles = data["profiles"][active_profile]["titles"]
    texts = data["profiles"][active_profile]["texts"]
    hotkeys = data["profiles"][active_profile]["hotkeys"]
    titles[i], titles[j] = titles[j], titles[i]
    texts[i], texts[j] = texts[j], texts[i]

#endregion

#region toggle edit mode

def toggle_edit_mode():
    global edit_mode
    if edit_mode:
        if has_field_changes():
            resp = show_question_message(
                "Ungespeicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Willst du sie speichern?"
            )
            if resp == QtWidgets.QMessageBox.Yes:
                save_data()
            else:
                reset_unsaved_changes()
        edit_mode = False
        update_ui()
        return
    is_sde_only = len(data["profiles"]) == 1 and "SDE" in data["profiles"]
    if active_profile == "SDE" and not is_sde_only:
        show_information_message("Nicht editierbar", "Das SDE-Profil kann nicht bearbeitet werden.")
        return

    edit_mode = True
    update_ui()
#endregion

#region save_data

def save_data(stay_in_edit_mode=False):
    global data, tray, active_profile, profile_entries
    try:
        # Only rename profiles if profile_entries is filled (edit mode)
        if edit_mode and profile_entries:
            updated_profiles = {}
            new_active_profile = active_profile
            for old_name, entry in profile_entries.items():
                new_name = entry.text().strip()
                if old_name == "SDE" or new_name == "SDE":
                    continue
                if new_name and new_name != old_name:
                    if new_name in data["profiles"]:
                        show_critical_message("Fehler", f"Profilname '{new_name}' existiert bereits!")
                        return
                    updated_profiles[new_name] = data["profiles"].pop(old_name)
                    if active_profile == old_name:
                        new_active_profile = new_name
                else:
                    updated_profiles[old_name] = data["profiles"][old_name]
            data["profiles"] = updated_profiles
            data["profiles"]["SDE"] = load_sde_profile()
            available_profiles = {**updated_profiles, "SDE": load_sde_profile()}
            if new_active_profile in available_profiles:
                active_profile = new_active_profile
            else:
                active_profile = list(available_profiles.keys())[0]
            if len(updated_profiles) > 11:
                show_critical_message("Limit erreicht", "Maximal 10 Profile erlaubt!")
                return
            data["active_profile"] = active_profile
        # Always update current profile's entries
        if active_profile != "SDE":
            data["profiles"][active_profile]["titles"] = [entry.text() for entry in title_entries]
            data["profiles"][active_profile]["texts"] = [entry.toHtml() if hasattr(entry, 'toHtml') else entry.text() for entry in text_entries]
            data["profiles"][active_profile]["hotkeys"] = [entry.text() for entry in hotkey_entries]
        profiles_to_save = {k: v for k, v in data["profiles"].items() if k != "SDE"}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"profiles": profiles_to_save, "active_profile": data["active_profile"]}, f, indent=4)
        fehlerhafte_hotkeys = register_hotkeys()
        reset_unsaved_changes()
        update_ui()
        if not fehlerhafte_hotkeys and not stay_in_edit_mode:
            toggle_edit_mode()
        refresh_tray()
    except Exception as e:
        show_critical_message("Fehler", f"Speichern fehlgeschlagen: {e}")
    reset_unsaved_changes()

def mark_unsaved_changes(event=None):
    global unsaved_changes
    unsaved_changes = True

def reset_unsaved_changes():
    global unsaved_changes
    unsaved_changes = False

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

data = load_data()
active_profile = data.get("active_profile", list(data["profiles"].keys())[0])

#endregion

#region Hauptfenster

app = QtWidgets.QApplication(sys.argv)
if sys.platform.startswith("win"):
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")

# Erstelle Hauptfenster
win = QtWidgets.QMainWindow()
win.setWindowTitle("QuickPaste")
win.setMinimumSize(399, 100)
win.setWindowIcon(QtGui.QIcon(os.path.join(os.path.dirname(__file__), "assets", "H.ico")))
QtCore.QTimer.singleShot(500, lambda: win.setWindowIcon(QtGui.QIcon(os.path.join(os.path.dirname(__file__), "assets", "H.ico"))))

# Add status bar for user feedback
win.statusBar().showMessage("Bereit")

# Close‑Event für Save und Tray
def close_event_handler(event):
    """Handle window close event - minimize to tray instead of closing"""
    if not win.isVisible():
        # Window is already hidden, don't process again
        event.ignore()
        return
    save_window_position()
    minimize_to_tray()
    event.ignore()  # Prevent the window from actually closing

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
entries_layout.setSpacing(6)  # Add space between entries
entries_layout.setContentsMargins(8, 8, 8, 8)  # Add margins around the container
scroll_area.setWidget(container)


#endregion

#region UI

def update_ui():
    global toolbar, entries_layout, active_profile, edit_mode, dark_mode, data
    global title_entries, text_entries, hotkey_entries, profile_entries
    # Clear entry lists
    title_entries = []
    text_entries = []
    hotkey_entries = []
    profile_entries = {}
    # Farben definieren
    bg    = "#2e2e2e" if dark_mode else "#eeeeee"
    fg    = "white"   if dark_mode else "black"
    ebg   = "#3c3c3c" if dark_mode else "white"
    bbg   = "#444"    if dark_mode else "#cccccc"

    win.setStyleSheet(f"background:{bg};")
    toolbar.setStyleSheet(f"background:{bg}; border: none;")
    container.setStyleSheet(f"background:{bg};")
    
    # Style status bar to match theme
    win.statusBar().setStyleSheet(f"""
        QStatusBar {{
            background: {bg};
            color: {fg};
            border-top: 1px solid #666;
        }}
    """)

    toolbar.clear()
    profs = [p for p in data["profiles"] if p!="SDE"]
    if not edit_mode and "SDE" in data["profiles"]:
        profs.append("SDE")


    for prof in profs:
        frame = QtWidgets.QWidget()
        frame.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        layout = QtWidgets.QHBoxLayout(frame)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        # In edit mode, show profile name input and buttons for non-SDE profiles
        # But also show clickable buttons for all profiles to allow switching
        if edit_mode and prof != "SDE":
            # Profile name input field
            entry = QtWidgets.QLineEdit(prof)
            entry.setFixedWidth(80)  # Slightly smaller to make room for switch button
            entry.setStyleSheet(
                f"background:{ebg}; color:{fg}; border-radius:5px; padding:4px;"
            )
            profile_entries[prof] = entry
            layout.addWidget(entry)

            # Switch to this profile button (new)
            switch_btn = QtWidgets.QPushButton("🖊️" if prof == active_profile else "🖊️")
            switch_btn.setFixedWidth(28)
            switch_btn.setStyleSheet(
                f"""
                background:{'#4a90e2' if prof == active_profile else bbg}; 
                color:{fg}; 
                border-radius:12px;
                font-size: 14px;
                """
            )
            switch_btn.setToolTip(f"Zu Profil '{prof}' wechseln")
            switch_btn.clicked.connect(lambda _, p=prof: switch_profile(p))
            layout.addWidget(switch_btn)

            # Delete button
            delete_btn = QtWidgets.QPushButton("❌")
            delete_btn.setFixedWidth(28)
            delete_btn.setStyleSheet(
                f"background:{bbg}; color:{fg}; border-radius:12px;"
            )
            delete_btn.clicked.connect(lambda _, p=prof: delete_profile(p))
            layout.addWidget(delete_btn)
        else:
            # Normal profile button (or SDE in edit mode)
            btn = QtWidgets.QPushButton(prof)
            btn.setStyleSheet(
                f"""
                QPushButton {{
                    background:{bbg}; color:{fg};
                    font-weight:{'bold' if prof==active_profile else 'normal'};
                    border-radius: 5px;
                    padding: 6px 16px;
                    margin-right: 4px;
                }}
                QPushButton:hover {{
                    background:#666;
                }}
                """
            )
            btn.clicked.connect(lambda _,p=prof: switch_profile(p))
            layout.addWidget(btn)

        toolbar.addWidget(frame)



    # Spacer to push action buttons to the right
    spacer = QtWidgets.QWidget()
    spacer.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
    toolbar.addWidget(spacer)

    # Add "➕ Profil" button just left of dark mode button in edit mode
    if edit_mode:
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

    # Action buttons (right side, except Help)
    for text, func, tooltip in [
        ("🌙" if not dark_mode else "🌞", toggle_dark_mode, "Dunkelmodus umschalten"),
        ("🔧", toggle_edit_mode, "Bearbeitungsmodus umschalten")
    ]:
        b = QtWidgets.QPushButton(text)
        b.setToolTip(tooltip)
        b.setStyleSheet(
            f"""
            QPushButton {{
                background:{bbg}; color:{fg};
                border-radius: 15px;
                min-width: 30px; min-height: 30px;
                font-size: 18px;
                margin-left: 6px;
                border: none;
            }}
            QPushButton:hover {{
                background:#888;
            }}
            """
        )
        b.clicked.connect(func)
        toolbar.addWidget(b)

    # Help button (always far right)
    help_btn = QtWidgets.QPushButton("❓")
    help_btn.setToolTip("Hilfe anzeigen")
    help_btn.setStyleSheet(
        f"""
        QPushButton {{
            background:{bbg}; color:{fg};
            border-radius: 15px;
            min-width: 30px; min-height: 30px;
            font-size: 18px;
            margin-left: 6px;
            border: none;
        }}
        QPushButton:hover {{
            background:#888;
        }}
        """
    )
    help_btn.clicked.connect(show_help_dialog)
    toolbar.addWidget(help_btn)


    # Einträge im ScrollArea aktualisieren
    # Alt löschen
    while entries_layout.count():
        w = entries_layout.takeAt(0).widget()
        if w: w.deleteLater()

    prof_data = data["profiles"][active_profile]
    titles, texts, hks = prof_data["titles"], prof_data["texts"], prof_data["hotkeys"]
    max_t = 120  # fixed width for title
    max_h = 120  # fixed width for hotkey

    for i, title in enumerate(titles):
        # Use custom drag-drop widget in edit mode
        if edit_mode:
            row = DragDropWidget(i)
        else:
            row = QtWidgets.QWidget()
        
        hl  = QtWidgets.QHBoxLayout(row)
        hl.setContentsMargins(8, 4, 8, 4)  # Slightly more padding
        hl.setSpacing(12)  # More space between elements

        # Add drag & drop functionality in edit mode
        if edit_mode:
            # Add drag handle
            drag_handle = QtWidgets.QLabel("☰")
            drag_handle.setFixedSize(20, 28)  # Set both width and height for more compact appearance
            drag_handle.setStyleSheet(f"""
                color: {fg}; 
                background: {bbg}; 
                padding: 2px 4px;
                border: 1px solid {'#555' if dark_mode else '#ccc'};
                border-radius: 4px;
                font-size: 14px;
                text-align: center;
            """)
            drag_handle.setAlignment(QtCore.Qt.AlignCenter)
            drag_handle.setToolTip("Ziehen zum Verschieben")
            
            # Store index for drag operations
            row.drag_index = i
            drag_handle.drag_index = i
            
            # Enable drag and drop
            row.setAcceptDrops(True)
            drag_handle.mousePressEvent = lambda event, idx=i: start_drag(event, idx, drag_handle)
            
            hl.addWidget(drag_handle)

        # Title
        if edit_mode:
            et = QtWidgets.QLineEdit(title)
            et.setFixedWidth(max_t)
            et.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if dark_mode else '#ccc'}; border-radius: 6px; padding: 8px;")
            et.editingFinished.connect(lambda idx=i, w=et: prof_data['titles'].__setitem__(idx, w.text()))
            hl.addWidget(et)
            title_entries.append(et)
        else:
            lt = QtWidgets.QLabel(title)
            lt.setFixedWidth(max_t)
            lt.setFixedHeight(40)  # Match button height
            lt.setStyleSheet(f"""
                color: {fg}; 
                background: {ebg}; 
                font-weight: bold; 
                padding: 10px 12px;
                border: 1px solid {'#555' if dark_mode else '#ccc'};
                border-radius: 6px;
                font-size: 13px;
            """)
            lt.setAlignment(QtCore.Qt.AlignVCenter)  # Vertical center alignment
            hl.addWidget(lt)

        # Text
        if edit_mode:
            ex = QtWidgets.QTextEdit(texts[i])
            ex.setMaximumHeight(80)  # Limit height for better layout
            ex.setMinimumHeight(60)  # Minimum height for usability
            ex.setStyleSheet(f"background:{ebg}; color:{fg};")
            ex.setAcceptRichText(True)  # Enable rich text support
            ex.setHtml(texts[i])  # Set initial content as HTML
            # Use a lambda that captures the current index correctly
            def make_text_handler(idx):
                return lambda: prof_data['texts'].__setitem__(idx, ex.toHtml())
            ex.textChanged.connect(make_text_handler(i))
            hl.addWidget(ex, 1)
            text_entries.append(ex)
        else:
            # Create a clickable button that shows only the first line of text
            text_btn = QtWidgets.QPushButton()
            text_btn.setStyleSheet(f"""
                QPushButton {{
                    background: {ebg}; 
                    color: {fg}; 
                    text-align: left; 
                    padding: 10px 12px;
                    border: 1px solid {'#555' if dark_mode else '#ccc'};
                    border-radius: 6px;
                    font-size: 13px;
                }}
                QPushButton:hover {{
                    background: {'#4a4a4a' if dark_mode else '#f0f0f0'};
                    border: 1px solid {'#666' if dark_mode else '#999'};
                }}
                QPushButton:pressed {{
                    background: {'#555' if dark_mode else '#e0e0e0'};
                }}
            """)
            text_btn.setFixedHeight(40)  # Fixed height for consistent appearance
            text_btn.setToolTip(f"Klicken zum Kopieren • Hotkey: {hks[i]}")
            
            # Extract and display only the first line of text
            display_text = QtGui.QTextDocument()
            display_text.setHtml(texts[i])
            plain_text = display_text.toPlainText()
            
            # Get only the first line
            first_line = plain_text.split('\n')[0].strip()
            
            # Truncate if too long for display
            if len(first_line) > 50:
                display_text = first_line[:47] + "..."
            else:
                display_text = first_line if first_line else "(Leer)"
            
            text_btn.setText(display_text)
            
            # Connect click to copy function
            def make_copy_handler(idx):
                return lambda: copy_text_to_clipboard(idx)
            text_btn.clicked.connect(make_copy_handler(i))
            
            hl.addWidget(text_btn, 1)

        # Hotkey
        if edit_mode:
            eh = QtWidgets.QLineEdit(hks[i])
            eh.setFixedWidth(max_h)
            eh.setStyleSheet(f"background:{ebg}; color:{fg}; border: 1px solid {'#555' if dark_mode else '#ccc'}; border-radius: 6px; padding: 8px;")
            
            # Add hotkey validation when editing finishes
            def validate_and_set_hotkey(idx, widget):
                hotkey = widget.text().strip().lower()
                erlaubte_zeichen = "1234567890befhmpqvxz§'^"
                
                if hotkey:  # Only validate if not empty
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
                        # Reset to original value
                        widget.setText(hks[idx])
                        return
                    
                    # Check for duplicates in current profile
                    current_hotkeys = [entry.text().strip().lower() for j, entry in enumerate(hotkey_entries) if j != idx]
                    if hotkey in current_hotkeys:
                        show_critical_message(
                            "Fehler", 
                            f"Hotkey \"{widget.text()}\" wird bereits in diesem Profil verwendet!"
                        )
                        # Reset to original value
                        widget.setText(hks[idx])
                        return
                
                # If validation passed, update the data
                prof_data['hotkeys'][idx] = widget.text()
            
            eh.editingFinished.connect(lambda idx=i, w=eh: validate_and_set_hotkey(idx, w))
            hl.addWidget(eh)
            hotkey_entries.append(eh)
            
            # Add delete button after hotkey
            delete_btn = QtWidgets.QPushButton("❌")
            delete_btn.setFixedSize(20, 20)
            delete_btn.setStyleSheet(f"""
                QPushButton {{
                    background: #d32f2f; 
                    color: white; 
                    border: 1px solid #b71c1c;
                    border-radius: 3px;
                    font-size: 10px;
                }}
                QPushButton:hover {{ background: #f44336; }}
            """)
            delete_btn.clicked.connect(lambda _, idx=i: delete_entry(idx))
            delete_btn.setToolTip("Eintrag löschen")
            hl.addWidget(delete_btn)
        else:
            lh = QtWidgets.QLabel(hks[i])
            lh.setFixedWidth(max_h)
            lh.setFixedHeight(40)  # Match button height
            lh.setStyleSheet(f"""
                color: {fg}; 
                background: {ebg}; 
                padding: 10px 12px;
                border: 1px solid {'#555' if dark_mode else '#ccc'};
                border-radius: 6px;
                font-size: 13px;
                font-family: 'Consolas', 'Monaco', monospace;
            """)
            lh.setAlignment(QtCore.Qt.AlignCenter)  # Center alignment for hotkeys
            hl.addWidget(lh)

        entries_layout.addWidget(row)

    # — Save/Add am Ende im Edit‑Mode —
    if edit_mode:
        bw = QtWidgets.QWidget()
        bl = QtWidgets.QHBoxLayout(bw)
        bl.setContentsMargins(0,0,0,0)
        bs = QtWidgets.QPushButton("💾 Speichern")
        bs.setStyleSheet("background:green;color:white;")
        bs.clicked.connect(save_data)
        bl.addWidget(bs)
        ba = QtWidgets.QPushButton("➕ Neuen Eintrag")
        ba.setStyleSheet(f"background:{bbg}; color:{fg};")
        ba.clicked.connect(add_new_entry)
        bl.addWidget(ba)
        entries_layout.addWidget(bw)

    win.show()

#endregion

#region help 

def show_help_dialog():
    help_text = (
        "QuickPaste Hilfe\n\n"
        "Wichtig: In Outlook kann es vorkommen, dass die Tastenkombination (z. B. Ctrl + Shift + 1) nicht sofort reagiert.\n"
        "Stellen Sie sicher, dass Sie die Zahl direkt nach 'Ctrl + Shift' drücken und versuchen Sie es ein zweites Mal.\n\n"
        "• Hotkeys aktiv: Aktiviert/Deaktiviert die Tastenkombinationen.\n"
        "   Wenn deaktiviert, greifen Windows Standardfunktionen.\n\n"
        "• ☾ / 🔆 Dunkelmodus: Wechselt zwischen hell/dunkel.\n\n"
        "• 🔧 Bearbeiten: Titel, Texte und Hotkeys anpassen.\n\n"
        "• ➕ Profil: Neues Textprofil erstellen.\n"
        "• �/📁 Profil wechseln: Im Bearbeitungsmodus zwischen Profilen wechseln.\n"
        "• ❌ Löschen: Profil entfernen (ausser SDE).\n\n"
        "• ↕ Verschieben: Einträge per Drag & Drop umsortieren.\n"
        "• ❌ Eintrag löschen: Klick auf ❌ neben dem Eintrag.\n"
        "• ➕ Eintrag: Neuer Eintrag hinzufügen.\n"
        "• 💾 Speichern: Änderungen sichern.\n\n"
        "Bei Fragen oder Problemen: nico.wagner@bit.admin.ch"
    )
    show_information_message("QuickPaste Hilfe", help_text)

#endregion

#region darkmode

def apply_dark_mode_to_messagebox(msg):
    """Apply dark mode styling to a QMessageBox if dark mode is enabled."""
    if dark_mode:
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
        
        # Additional fallback: manually set button text color
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
    """Show a critical message with proper dark mode styling."""
    if parent is None:
        parent = win
    msg = QtWidgets.QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setIcon(QtWidgets.QMessageBox.Critical)
    apply_dark_mode_to_messagebox(msg)
    return msg.exec_()

def show_question_message(title, text, buttons=QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No, parent=None):
    """Show a question message with proper dark mode styling."""
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
    """Show an information message with proper dark mode styling."""
    if parent is None:
        parent = win
    msg = QtWidgets.QMessageBox(parent)
    msg.setWindowTitle(title)
    msg.setText(text)
    msg.setIcon(QtWidgets.QMessageBox.Information)
    apply_dark_mode_to_messagebox(msg)
    msg.exec_()

def toggle_dark_mode():
    global dark_mode
    dark_mode = not dark_mode
    update_ui()
    save_window_position()  

#endregion



# ——————————————————————————————
# Programmstart
# ——————————————————————————————
load_window_position()
update_ui()
create_tray_icon()
register_hotkeys()
win.show()
sys.exit(app.exec_())