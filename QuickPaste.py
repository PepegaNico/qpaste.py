import tkinter as tk
from tkinter import messagebox
import pyperclip
import keyboard
import json
import time
from pystray import Menu, Icon, MenuItem
from PIL import Image
import threading
import sys
import os
import logging
import win32clipboard
import win32con


APPDATA_PATH = os.path.join(os.environ["APPDATA"], "QuickPaste")
os.makedirs(APPDATA_PATH, exist_ok=True)
CONFIG_FILE = os.path.join(APPDATA_PATH, "config.json")
WINDOW_CONFIG = os.path.join(APPDATA_PATH, "window_config.json")
LOG_FILE = os.path.join(APPDATA_PATH, "qp.log")
SDE_FILE = os.path.join(APPDATA_PATH, "sde.json")
BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
ICON_PATH = os.path.join(BASE_DIR, "assets", "H.ico")
dark_mode = False
registered_hotkey_refs = []
unsaved_changes = False
dragged_index = None
current_target = None
logging.basicConfig(filename=LOG_FILE, filemode="a", level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8")
hotkeys_enabled = True


#region Hyperlink

def make_rtf_hyperlink(url: str, display: str) -> str:
    esc = display.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")
    return (
        r"{\field{\*\fldinst{HYPERLINK \"" + url + r"\"}}"
        r"{\fldrslt{" + esc + r"}}}"
    )

def set_clipboard_rtf(rtf_body: str, plain: str):
    # Umhülle den Body
    full = r"{\rtf1\ansi " + rtf_body + "}"
    try:
        data_bytes = full.encode("cp1252", errors="replace")
    except:
        data_bytes = full.encode("utf-8", errors="ignore")
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, plain)
        cf = win32clipboard.RegisterClipboardFormat("Rich Text Format")
        win32clipboard.SetClipboardData(cf, data_bytes)
    finally:
        win32clipboard.CloseClipboard()

def html_to_rtf_fragment(html: str) -> str:
    """
    Übersetzt alle <a href="URL">Text</a> in RTF-Felder, und entfernt
    Backslashes/Klammern korrekt. Gibt nur den RTF-Body zurück, ohne
    den {\rtf1…}-Wrapper.
    """
    # 1) Escape reiner Text (Backslashes und geschweifte Klammern)
    esc = html.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")

    # 2) Ersetze nacheinander alle Links
    def repl_link(m):
        url, txt = m.group(1), m.group(2)
        # baue RTF-Hyperlink-Feld
        disp = txt.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")
        return (
            r"{\field{\*\fldinst{HYPERLINK \"" + url + r"\"}}"
            r"{\fldrslt{" + disp + r"}}}"
        )
    esc = re.sub(r'<a\s+href="([^"]+)">(.*?)</a>', repl_link, esc, flags=re.DOTALL)

    return esc

import tkinter as tk
from tkinter import simpledialog, messagebox
import re
def open_format_dialog(idx):
    # Popup öffnen
    dlg = tk.Toplevel(root)
    dlg.iconbitmap(ICON_PATH) 
    dlg.transient(root)
    dlg.grab_set()
    dlg.focus_set()
    dlg.title("Hyperlink einfügen")

    # Haupt‑Save-Button deaktivieren
    save_button.config(state="disabled")



    # 2) Einfache Text‑Box (zeigt das rohe HTML, damit Link erhalten bleibt)
    txt = tk.Text(dlg, wrap="word", width=60, height=10)
    raw = data["profiles"][active_profile]["texts"][idx]
    txt.insert("1.0", raw)   # rohes HTML einfügen, damit <a href=...>…</a> stehen bleibt
 
    
    txt.pack(padx=10, pady=10, fill="both", expand=True)
    btnf = tk.Frame(dlg)
    btnf.pack(fill="x", pady=(0,10))

    def add_link():
        try:
            start = txt.index("sel.first")
            end   = txt.index("sel.last")
            sel   = txt.get(start, end)

        except tk.TclError:
            return messagebox.showwarning("Keine Auswahl", "Bitte erst Text markieren!")
        url = simpledialog.askstring("Hyperlink einfügen", "Gib die URL ein:", initialvalue="https://")

        if not url:
            return
        txt.delete(start, end)
        txt.insert(start, f'<a href="{url}">{sel}</a>')
    link_btn = tk.Button(btnf, text="🔗 Link einfügen", command=add_link)
    link_btn.pack(side="left", padx=5)

    def on_ok():
        new_html = txt.get("1.0", "end-1c")
        entry = text_entries[idx]
        entry._html = new_html
        data["profiles"][active_profile]["texts"][idx] = new_html
        entry.delete(0, tk.END)
        entry.insert(0, re.sub(r"<a[^>]*>(.*?)</a>", r"\1", new_html))
        save_button.config(state="normal")
        dlg.destroy()
        mark_unsaved_changes()
    ok_btn = tk.Button(btnf, text="OK", command=on_ok)
    ok_btn.pack(side="right", padx=5)
    
    def on_cancel():
        save_button.config(state="normal")
        dlg.destroy()

    dlg.protocol("WM_DELETE_WINDOW", on_cancel)

#endregion

#region data

def load_sde_profile():
    try:
        with open(SDE_FILE, "r", encoding="utf-8") as f:
            sde = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        logging.warning("⚠ Konnte sde.json nicht laden. Standardwerte für SDE werden gesetzt.")
        sde = {}
    if not sde.get("titles") and not sde.get("texts") and not sde.get("hotkeys"):
        sde = {
            "titles": ["Standard Titel 1", "Standard Titel 2", "Standard Titel 3"],
            "texts": ["Standard Text 1", "Standard Text 2", "Standard Text 3"],
            "hotkeys": ["ctrl+shift+1", "ctrl+shift+2", "ctrl+shift+3"]
        }
    return sde

def load_data():
    try:
        with open(CONFIG_FILE, "r") as file:
            loaded_data = json.load(file)
            if "profiles" not in loaded_data or not isinstance(loaded_data["profiles"], dict):
                loaded_data["profiles"] = {}
            for profile in loaded_data["profiles"]:
                if "titles" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["titles"] = []
                if "texts" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["texts"] = []
                if "hotkeys" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["hotkeys"] = []
            sde_profile = load_sde_profile()
            loaded_data["profiles"]["SDE"] = sde_profile
            if "active_profile" not in loaded_data or loaded_data["active_profile"] not in loaded_data["profiles"]:
                if loaded_data["profiles"]:
                    first_profile = list(loaded_data["profiles"].keys())[0]
                    loaded_data["active_profile"] = first_profile
                else:
                    logging.warning("⚠ Keine Profile gefunden. Erstelle Standardprofil.")
                    loaded_data["profiles"] = {"Profil 1": {"titles": [], "texts": [], "hotkeys": []}}
                    loaded_data["active_profile"] = "Profil 1"
            return loaded_data
    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "profiles": {
                "Profil 1": {
                    "titles": [f"Titel {i}" for i in range(1, 6)],
                    "texts": [f"Text {i}" for i in range(1, 6)],
                    "hotkeys": [f"ctrl+shift+{i}" for i in range(1, 6)]
                },
                "Profil 2": {
                    "titles": [f"Titel {i}" for i in range(1, 6)],
                    "texts": [f"Profil 2 Text {i}" for i in range(1, 6)],
                    "hotkeys": [f"ctrl+shift+{i}" for i in range(1, 6)]
                },
                "SDE": load_sde_profile()
            },
            "active_profile": "Profil 1"
        }

#endregion

#region windows position 

def save_window_position():
    if root:
        window_geometry = root.geometry()
        config_data = {"geometry": window_geometry, "dark_mode": dark_mode}
        try:
            with open(WINDOW_CONFIG + "_tmp", "w", encoding="utf-8") as temp_file:
                json.dump(config_data, temp_file)
            os.replace(WINDOW_CONFIG + "_tmp", WINDOW_CONFIG)
        except Exception as e:
            logging.exception(f"⚠ Fehler beim Speichern der Fensterposition: {e}")

def load_window_position():
    global dark_mode
    try:
        with open(WINDOW_CONFIG, "r") as file:
            content = file.read().strip()
            if not content:
                raise ValueError("⚠ window_config.json ist leer.")
            saved_config = json.loads(content)
            if "dark_mode" in saved_config:
                dark_mode = saved_config["dark_mode"]
            return saved_config.get("geometry")
    except (FileNotFoundError, json.JSONDecodeError, ValueError) as e:
        logging.exception(f"⚠ Fensterposition konnte nicht geladen werden: {e}")
        return None

#endregion

#region profiles

def save_profile_names():
    global data
    if not edit_mode:
        return 
    new_profiles = {}
    for old_name, entry in profile_entries.items():
        new_name = entry.get().strip()
        if new_name and new_name not in new_profiles:
            new_profiles[new_name] = data["profiles"].pop(old_name)  
        else:
            messagebox.showerror("Fehler", f"Profilname \"{new_name}\" ist ungültig oder bereits vergeben!")
            return
    data["profiles"] = new_profiles  
    save_data() 
    update_ui() 

def update_profile_buttons():
    for profile, button in profile_buttons.items():
        if profile == active_profile:
            button.config(bg="lightblue", font=("Arial", 10, "bold"))  
        else:
            button.config(bg="SystemButtonFace", font=("Arial", 10, "normal"))  

def has_field_changes():
    current_profile = data["profiles"].get(active_profile, {})
    edited_titles = [entry.get() for entry in title_entries if entry.winfo_exists()]
    edited_texts = [entry.get() for entry in text_entries if entry.winfo_exists()]
    edited_hotkeys = [entry.get() for entry in hotkey_entries if entry.winfo_exists()]
    return (
        edited_titles != current_profile.get("titles", []) or
        edited_texts != current_profile.get("texts", []) or
        edited_hotkeys != current_profile.get("hotkeys", [])
    )

def switch_profile(profile_name):
    global active_profile, data, edit_mode, tray_icon
    if edit_mode and has_field_changes():
        response = messagebox.askyesnocancel(
            "Nicht gespeicherte Änderungen",
            "Es gibt ungespeicherte Änderungen! Möchtest du sie speichern?",
            icon="warning"
        )
        if response is None:
            return  
        elif response:  
            save_data(stay_in_edit_mode=True)
            reset_unsaved_changes()
        else: 
            reset_unsaved_changes()
    if profile_name not in data["profiles"]:
        messagebox.showerror("Fehler", f"Das Profil '{profile_name}' existiert nicht!")
        return
    active_profile = profile_name
    data["active_profile"] = profile_name
    with open(CONFIG_FILE, "w") as file:
        json.dump(data, file, indent=4)
    reset_unsaved_changes()
    update_profile_buttons()  
    update_ui()  
    refresh_tray()

def add_new_profile():
    global data
    if len(data["profiles"]) >= 11:
        messagebox.showerror("Limit erreicht", "Maximal 10 Profile erlaubt!")
        return
    new_name_base = "Profil"
    counter = 1
    while f"{new_name_base} {counter}" in data["profiles"]:
        counter += 1
    new_profile_name = f"{new_name_base} {counter}"
    titles = [f"Titel {i+1}" for i in range(3)]
    texts = [f"Text {i+1}" for i in range(3)]
    hotkeys = [f"ctrl+shift+{i+1}" for i in range(3)]
    data["profiles"][new_profile_name] = {"titles": titles, "texts": texts, "hotkeys": hotkeys}
    data["active_profile"] = new_profile_name
    update_ui()

def delete_profile(profile_name):
    global data, active_profile
    if profile_name == "SDE":
        messagebox.showerror("Fehler", "Das SDE Profil kann nicht gelöscht werden.")
        return
    if len(data["profiles"]) <= 1:
        messagebox.showerror("Fehler", "Mindestens ein Profil muss vorhanden sein!")
        return
    if messagebox.askyesno("Profil löschen", f"Soll das Profil '{profile_name}' wirklich gelöscht werden?"):
        del data["profiles"][profile_name]
        if active_profile == profile_name:
            active_profile = list(data["profiles"].keys())[0]
            data["active_profile"] = active_profile
        update_ui()
        save_data(stay_in_edit_mode=True)

def make_profile_switcher(profile_name):
    def handler(icon, item):
        switch_profile(profile_name)
    return handler

#endregion

#region Hauptfenster ( main(): ) 

root = tk.Tk()
if sys.platform.startswith("win"): 
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
root.minsize(399, 100)
root.iconbitmap(ICON_PATH) 
root.after(500, lambda: root.iconbitmap(ICON_PATH))
root.title("QuickPaste")
root.protocol("WM_DELETE_WINDOW", lambda: [save_window_position(), minimize_to_tray()])
geometry = load_window_position()
if geometry:
    try:
        root.geometry(geometry)
        logging.info(f"Fensterposition geladen: {geometry}")
    except Exception as e:
        logging.warning(f"⚠ Konnte Fensterposition nicht setzen: {e}")

#endregion

#region insert_text / Register Hotkey

def insert_text(index):
    raw = data["profiles"][active_profile]["texts"][index]
    # 1) kompletten RTF-Body erzeugen (alle Links)
    body = html_to_rtf_fragment(raw)
    full_rtf = r"{\rtf1\ansi " + body + "}"
    # 2) Plain-Text fürs reine Einfügen
    plain = re.sub(r"<[^>]+>", "", raw)
    # 3) Clipboard füllen
    set_clipboard_rtf(full_rtf, plain)
    # 4) Einfügen
    time.sleep(0.1)
    keyboard.send("ctrl+v")

def register_hotkeys():
    if active_profile in data["profiles"]:
        global registered_hotkey_refs
        for ref in registered_hotkey_refs:
            try:
                keyboard.remove_hotkey(ref)
            except Exception:
                pass
        registered_hotkey_refs.clear()
    erlaubte_zeichen = "1234567890befhmpqvxz§'^"
    belegte_hotkeys = set()
    fehlerhafte_hotkeys = False
    data["profiles"].setdefault(active_profile, {}).setdefault("hotkeys", [])
    def hotkey_handler(index):
        gedrückte_tasten = keyboard._pressed_events.keys()
        pfeiltasten = {72, 80, 75, 77}  
        if any(key in gedrückte_tasten for key in pfeiltasten):
            logging.warning("⚠ Windows-Funktion erkannt, kein Text eingefügt.")
            return  
        insert_text(index)
        keyboard.release("shift")
        keyboard.release("ctrl")   
    for i, hotkey in enumerate(data["profiles"][active_profile]["hotkeys"]):
        if hotkey.strip():  
            keys = hotkey.lower().split("+")
            if (
                len(keys) != 3
                or not keys[2] 
                or "ctrl" not in keys
                or "shift" not in keys
                or keys[2] not in erlaubte_zeichen
            ):
                messagebox.showerror("Fehler", f"Ungültiger Hotkey \"{hotkey}\" für Eintrag {i+1}.\nErlaubte Zeichen: {erlaubte_zeichen}")
                fehlerhafte_hotkeys = True
                continue
            if hotkey in belegte_hotkeys:
                messagebox.showerror("Fehler", f"Hotkey \"{hotkey}\" wird bereits verwendet!")
                fehlerhafte_hotkeys = True  
                continue  
            if i >= len(data["profiles"][active_profile]["texts"]):
                logging.warning(f"⚠ Hotkey '{hotkey}' zeigt auf Eintrag {i+1}, aber dieser existiert nicht im Profil '{active_profile}'")
                continue
            def wrapped_handler(index=i, keys=keys):
                pressed = keyboard._pressed_events
                if all(keyboard.is_pressed(k) for k in keys):
                    insert_text(index)
            ref = keyboard.add_hotkey(hotkey, wrapped_handler, suppress=True)
            registered_hotkey_refs.append(ref)
    return fehlerhafte_hotkeys

#endregion

#region tray 

tray_icon = None 

def create_tray_icon():
    global tray_icon
    if tray_icon is not None:
        try:
            tray_icon.stop()
        except:
            pass
        tray_icon = None
    image = Image.open(ICON_PATH)
    profile_items = []
    for profile in data["profiles"].keys():
        label = f"✓ {profile}" if profile == active_profile else f"  {profile}"
        profile_items.append(MenuItem(label, make_profile_switcher(profile)))
    tray_icon = Icon("QuickPaste", image, menu=Menu(
        *profile_items,
        Menu.SEPARATOR,
        MenuItem("Hotkeys aktiv", toggle_hotkeys_from_tray, checked=lambda item: hotkeys_enabled),
        Menu.SEPARATOR,
        MenuItem("↑ Öffnen", show_window),
        MenuItem("✖ Beenden", quit_application)
    ))
    tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
    tray_thread.start()

def refresh_tray():
    global tray_icon
    if tray_icon is not None:
        try:
            tray_icon.stop()
        except:
            pass
        tray_icon = None
    create_tray_icon()

def minimize_to_tray():
    root.withdraw()   

def show_window(icon, item):
    root.deiconify()
    refresh_tray()

def quit_application(icon, item):
    try:
        window_geometry = root.geometry()
        config_data = {"geometry": window_geometry, "dark_mode": dark_mode}
        with open(WINDOW_CONFIG + "_tmp", "w", encoding="utf-8") as temp_file:
            json.dump(config_data, temp_file)
        os.replace(WINDOW_CONFIG + "_tmp", WINDOW_CONFIG)
    except Exception as e:
        logging.exception(f"⚠ Fehler beim Speichern der window_config.json: {e}")
    root.quit()
    root.destroy()
    sys.exit(0)

def toggle_hotkeys_from_tray(icon, item):
    global hotkeys_enabled
    hotkeys_enabled = not hotkeys_enabled
    hotkey_checkbox_var.set(hotkeys_enabled) 
    register_hotkeys()                        
    refresh_tray()

#endregion

#region add/del/move/drag Entry

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
        messagebox.showerror("Fehler", "Ungültiger Eintrag zum Löschen ausgewählt!")
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

edit_mode = False
text_entries = []
title_entries = []
hotkey_entries = []

#region toggle edit mode

def toggle_edit_mode():
    global edit_mode
    if edit_mode:
        if has_field_changes():
            if messagebox.askyesno(
                "Ungespeicherte Änderungen",
                "Du hast ungespeicherte Änderungen. Willst du sie speichern?"
            ):
                save_data()
            else:
                reset_unsaved_changes()
        edit_mode = False
        update_ui()
        return
    is_sde_only = len(data["profiles"]) == 1 and "SDE" in data["profiles"]
    if active_profile == "SDE" and not is_sde_only:
        messagebox.showinfo("Nicht editierbar", "Das SDE-Profil kann nicht bearbeitet werden.")
        return

    edit_mode = True
    update_ui()
#endregion

#region save_data

def save_data(stay_in_edit_mode=False):
    global data, tray_icon, active_profile, profile_entries
    try:
        updated_profiles = {}
        new_active_profile = active_profile 
        for old_name, entry in profile_entries.items():
            new_name = entry.get().strip()
            if old_name == "SDE" or new_name == "SDE":
                continue 
            if new_name and new_name != old_name:
                if new_name in data["profiles"]:
                    messagebox.showerror("Fehler", f"Profilname '{new_name}' existiert bereits!")
                    return
                updated_profiles[new_name] = data["profiles"].pop(old_name) 
                if active_profile == old_name:
                    new_active_profile = new_name
            else:
                updated_profiles[old_name] = data["profiles"][old_name] 
        data["profiles"] = updated_profiles
        refresh_tray()
        data["profiles"]["SDE"] = load_sde_profile()
        available_profiles = {**updated_profiles, "SDE": load_sde_profile()}
        if new_active_profile in available_profiles:
            active_profile = new_active_profile
        else:
            active_profile = list(available_profiles.keys())[0]
        if len(updated_profiles) > 11:
            messagebox.showerror("Limit erreicht", "Maximal 10 Profile erlaubt!")
            return
        data["active_profile"] = active_profile 
        if active_profile != "SDE":
            data["profiles"][active_profile]["titles"] = [entry.get() for entry in title_entries if entry.winfo_exists()]
           
            new_texts = []
            for e in text_entries:
                html = getattr(e, "_html", None)
                if html:
                    new_texts.append(html)
                else:
                    new_texts.append(e.get())
            data["profiles"][active_profile]["texts"] = new_texts

            
            
            data["profiles"][active_profile]["hotkeys"] = [entry.get() for entry in hotkey_entries if entry.winfo_exists()]
        profiles_to_save = {k: v for k, v in data["profiles"].items() if k != "SDE"}
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"profiles": profiles_to_save, "active_profile": data["active_profile"]}, f, indent=4)
        fehlerhafte_hotkeys = register_hotkeys()
        reset_unsaved_changes() 
        update_ui()
        if not fehlerhafte_hotkeys and not stay_in_edit_mode:
            toggle_edit_mode()
    except Exception as e:
        messagebox.showerror("Fehler", f"Speichern fehlgeschlagen: {e}")
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
    result = messagebox.askyesno("Ungespeicherte Änderungen", "Du hast ungespeicherte Änderungen.\nWillst du sie speichern?")
    if result:
        save_data(stay_in_edit_mode=True)
    else:
        reset_unsaved_changes()
    root.after(100, action_if_yes)

data = load_data()
active_profile = data.get("active_profile", list(data["profiles"].keys())[0])

#endregion

#region UI

def update_ui():
    global profile_buttons, profile_entries, title_labels, scrollable_frame, canvas, edit_mode
    global active_profile
    global hotkey_checkbox_var
    is_read_only = active_profile == "SDE"
    is_sde_only = len([p for p in data["profiles"].keys() if p != "SDE"]) == 0 and "SDE" in data["profiles"]
    bg_color = "#2e2e2e" if dark_mode else "SystemButtonFace"
    fg_color = "white" if dark_mode else "black"   
    entry_bg = "#3c3c3c" if dark_mode else "white"
    entry_fg = "white" if dark_mode else "black"
    button_bg = "#444" if dark_mode else "SystemButtonFace"
    root.configure(bg=bg_color)
    profile_buttons = {}
    profile_entries = {}
    title_labels = []
    for widget in root.winfo_children():
        if widget.winfo_name() not in ["frame_container", "top_frame"]:
            widget.destroy()
    top_frame = tk.Frame(root, name="top_frame", bg=bg_color, highlightthickness=0, bd=0)
    top_frame.pack(fill="x", padx=0, pady=0)
    if edit_mode:
        profiles = [p for p in data["profiles"].keys() if p != "SDE"]
    else:
        profiles = [p for p in data["profiles"].keys() if p != "SDE"]
        if "SDE" in data["profiles"]:
            profiles.append("SDE")
    for profile in profiles:
        frame = tk.Frame(top_frame, bg=bg_color)
        frame.pack(side="left", padx=3)
        if edit_mode:
            entry = tk.Entry(frame, width=12, bg=entry_bg, fg=entry_fg)
            entry.insert(0, profile)
            if profile == active_profile:
                entry.config(highlightthickness=3, highlightbackground="black")
            entry.pack(side="left")
            profile_entries[profile] = entry
            edit_btn = tk.Button(frame, text="🖊", command=lambda p=profile: switch_profile(p), width=2, bg=button_bg, fg=fg_color, activebackground=button_bg, activeforeground=fg_color)
            edit_btn.pack(side="left", padx=2)
            delete_btn = tk.Button(frame, text="❌", command=lambda p=profile: delete_profile(p), bg=button_bg, fg=fg_color, activebackground=button_bg, activeforeground=fg_color)
            delete_btn.pack(side="left", padx=2)
        else:
            btn = tk.Button(frame, text=profile, command=lambda p=profile: switch_profile(p), bg=button_bg, fg=fg_color, activebackground=button_bg, activeforeground=fg_color, font=("Arial", 10, "bold") if profile == active_profile else ("Arial", 10, "normal"))
            btn.pack(side="left", padx=3)
            profile_buttons[profile] = btn
    help_button = tk.Button(top_frame, text="❓", command=show_help_dialog, bg=button_bg, fg=fg_color)
    help_button.pack(side="right", padx=5)          
    settings_icon = tk.Button(top_frame, text="🔧", command=toggle_edit_mode, width=3, bg=button_bg, fg=fg_color)
    settings_icon.pack(side="right", padx=5, pady=5)
    dark_mode_button = tk.Button(top_frame, text="☾" if not dark_mode else "🔆", command=toggle_dark_mode, bg=button_bg, fg=fg_color)
    dark_mode_button.pack(side="right", padx=5)
    hotkey_checkbox_var = tk.BooleanVar(value=hotkeys_enabled)
    hotkey_checkbox = tk.Checkbutton( top_frame, text="Hotkeys aktiv", variable=hotkey_checkbox_var, command=lambda: toggle_hotkeys_from_ui(hotkey_checkbox_var.get()), bg=bg_color, fg=fg_color, selectcolor=bg_color)
    hotkey_checkbox.pack(side="right", padx=5, pady=5)
    if edit_mode:
        add_profile_btn = tk.Button(top_frame, text="➕ Profil", command=lambda: confirm_and_then(add_new_profile), bg=button_bg, fg=fg_color)
        add_profile_btn.pack(side="right", padx=5)
    frame_container = tk.Frame(root, name="frame_container", bg=bg_color)
    frame_container.pack(fill="both", expand=True)
    canvas = tk.Canvas(frame_container, bg=bg_color, highlightthickness=0)
    scrollbar = tk.Scrollbar(frame_container, orient="vertical", command=canvas.yview, bg=bg_color)
    scrollable_frame = tk.Frame(canvas, bg=bg_color)
    def update_scroll_region(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
    scrollable_frame.bind("<Configure>", update_scroll_region)
    window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=frame_container.winfo_width())
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    def resize_canvas(event):
        canvas.itemconfig(window, width=event.width)
    canvas.bind("<Configure>", resize_canvas)
    def _on_mouse_wheel(event):
        if root.winfo_height() < scrollable_frame.winfo_height():
            canvas.yview_scroll(-1 * (event.delta // 120), "units")
    canvas.bind("<Enter>", lambda e: root.bind_all("<MouseWheel>", _on_mouse_wheel))
    canvas.bind("<Leave>", lambda e: root.unbind_all("<MouseWheel>"))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)
    global title_entries, text_entries, hotkey_entries
    title_entries, text_entries, hotkey_entries = [], [], []
    titles = data["profiles"][active_profile]["titles"]
    max_title_length = max((len(t) for t in titles), default=10)
    texts = data["profiles"][active_profile]["texts"]
    min_text_lenght = min((len(t) for t in texts), default=10)
    hotkeys = data["profiles"][active_profile]["hotkeys"]
    max_hotkey_lenght = max((len(t) for t in hotkeys), default=10)
    for i, title in enumerate(data["profiles"][active_profile]["titles"]):
        frame = tk.Frame(scrollable_frame, bg=bg_color)
        frame.orig_bg = bg_color 
        frame.pack(fill="x", expand=True, padx=5, pady=2)
        frame.drag_index = i
        if is_read_only:
            tk.Label(frame, text=title, width=max_title_length, bg=bg_color, fg=fg_color, anchor="w").pack(side="left", padx=5)
            tk.Label(frame, text=data["profiles"][active_profile]["texts"][i], bg=bg_color, fg=fg_color, anchor="center", justify="center").pack(side="left", padx=5, expand=True, fill="x")
            tk.Label(frame, text=data["profiles"][active_profile]["hotkeys"][i], width=12, bg=bg_color, fg=fg_color, anchor="e").pack(side="right", padx=5)
            continue
        if edit_mode:
            drag_handle = tk.Label(frame, text="↕", bg=bg_color, fg=fg_color, cursor="fleur")
            drag_handle.pack(side="left", padx=2)
            drag_handle.bind("<ButtonPress-1>", lambda e, idx=i: on_drag_start(e, idx))
            drag_handle.bind("<B1-Motion>", on_drag_motion)
            drag_handle.bind("<ButtonRelease-1>", on_drag_release)
        if edit_mode:
            title_entry = tk.Entry(frame, width=max_title_length, bg=bg_color, fg=fg_color)
            title_entry.insert(0, title)
            title_entry.pack(side="left", padx=5)
            title_entry.bind("<KeyRelease>", mark_unsaved_changes)
            title_entries.append(title_entry)
        else:
            tk.Label(frame, text=title, width=max_title_length, bg=bg_color, fg=fg_color, anchor="w").pack(side="left", padx=5)
        if edit_mode:
            text_entry = tk.Entry(frame, bg=bg_color, fg=fg_color)
            text_entry.insert(0, data["profiles"][active_profile]["texts"][i])
            text_entry.pack(side="left", fill="x", expand=True, padx=5)
            text_entry.bind("<KeyRelease>", mark_unsaved_changes)
            text_entries.append(text_entry)
        else:
            raw = data["profiles"][active_profile]["texts"][i]
            # Alle <…>‑Tags wegfiltern für die Anzeige
            plain = re.sub(r"<[^>]+>", "", raw)
            lines = [line.strip() for line in plain.split("\n") if line.strip()]
            display_text = "\n".join(lines[:2])
            text_button = tk.Button(frame, text=display_text, command=lambda i=i: insert_text(i), wraplength=frame.winfo_width() - 150, justify="left", anchor="w", padx=5, pady=5, width=min_text_lenght, height=2,bg=bg_color, fg=fg_color)
            text_button.pack(side="left", fill="x", expand=True, padx=5, pady=2)
        if edit_mode:
            format_btn = tk.Button(frame, text="✎ Editor", command=lambda i=i: open_format_dialog(i), bg=button_bg, fg=fg_color)
            format_btn.pack(side="left", padx=2)
        if edit_mode:
            hotkey_entry = tk.Entry(frame, width=12, bg=bg_color, fg=fg_color)
            hotkey_entry.insert(0, data["profiles"][active_profile]["hotkeys"][i])
            hotkey_entry.pack(side="left", padx=5)
            hotkey_entry.bind("<KeyRelease>", mark_unsaved_changes)
            hotkey_entries.append(hotkey_entry)
        else:
            hotkey_label = tk.Label(frame, text=data["profiles"][active_profile]["hotkeys"][i], width=max_hotkey_lenght, anchor="e", bg=bg_color, fg=fg_color)
            hotkey_label.pack(side="right", padx=5)
        if edit_mode:
            delete_button = tk.Button(frame, text="❌", width=2, height=0, bg=button_bg, fg=fg_color, font=("Arial", 12), command=lambda i=i: confirm_and_then(lambda: delete_entry(i)))
            delete_button.pack(side="left", padx=5)
    if edit_mode:
        buttons_frame = tk.Frame(root, bg=bg_color)
        buttons_frame.pack(fill="x", pady=5)
        global save_button
        save_button = tk.Button(buttons_frame, text="💾 Speichern", command=lambda: save_data(stay_in_edit_mode=False), fg="white", bg="green", height=2)
        save_button.pack(side="left", expand=True, fill="x", padx=5)
        add_button = tk.Button(buttons_frame, text="➕ Neuen Eintrag", command=lambda: confirm_and_then(add_new_entry), bg=button_bg, fg=fg_color, activebackground=button_bg, activeforeground=fg_color, height=2)
        add_button.pack(side="right", expand=True, fill="x", padx=5)
    root.update_idletasks()

#endregion

def toggle_hotkeys_from_ui(value):
    global hotkeys_enabled
    hotkeys_enabled = value
    if hotkey_checkbox_var.get() != value:
        hotkey_checkbox_var.set(value)

    # NEU: Bei Änderung Hotkeys neu registrieren bzw. entfernen
    register_hotkeys()
    refresh_tray()

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
        "• 🖊 Bearbeiten: Profilnamen ändern.\n"
        "• ❌ Löschen: Profil entfernen (ausser SDE).\n\n"
        "• ↕ Verschieben: Einträge per Drag & Drop umsortieren.\n"
        "• ❌ Eintrag löschen: Klick auf ❌ neben dem Eintrag.\n"
        "• ➕ Eintrag: Neuer Eintrag hinzufügen.\n"
        "• 💾 Speichern: Änderungen sichern.\n\n"
        "Bei Fragen oder Problemen: nico.wagner@bit.admin.ch"
    )
    tk.messagebox.showinfo("QuickPaste Hilfe", help_text)

#endregion

#region darkmode

def toggle_dark_mode():
    global dark_mode
    dark_mode = not dark_mode
    update_ui()
    save_window_position()  

#endregion

saved_geometry = load_window_position()
if saved_geometry:
    root.geometry(saved_geometry) 
else:
    root.geometry("300x650") 
update_ui()
register_hotkeys()  
create_tray_icon()
logging.info("QuickPaste gestartet")   
root.mainloop()