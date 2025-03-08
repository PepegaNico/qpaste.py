import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import keyboard
import json
import os
import time
import pystray
from pystray import Icon, Menu, MenuItem
from PIL import Image
import threading
import sys

CONFIG_FILE = "config.json"
WINDOW_CONFIG = "window_config.json"


def debug_tasten(event):
    print(f"🔍 Taste gedrückt: {event.name} (Scan-Code: {event.scan_code})")

keyboard.hook(debug_tasten)

#region data

default_data = {
    "titles": ["Titel"] * 10, 
    "texts": [f"Text {i+1}" for i in range(10)],
    "hotkeys": [
        "ctrl+shift+1", "ctrl+shift+2", "ctrl+shift+3", "ctrl+shift+4", "ctrl+shift+5",
        "ctrl+shift+6", "ctrl+shift+7", "ctrl+shift+8", "ctrl+shift+9", "ctrl+shift+0"
    ]
}


def load_data():
    try:
        with open(CONFIG_FILE, "r") as file:
            loaded_data = json.load(file)

            # **Sicherstellen, dass "profiles" existiert**
            if "profiles" not in loaded_data or not isinstance(loaded_data["profiles"], dict):
                loaded_data["profiles"] = {}

            # **Falls keine Profile existieren, Standardprofil erstellen**
            if not loaded_data["profiles"]:
                loaded_data["profiles"] = {
                    "Standard": {"titles": [], "texts": [], "hotkeys": []}
                }

            # **Alle Profile prüfen und fehlende Felder ergänzen**
            for profile in loaded_data["profiles"]:
                if "titles" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["titles"] = []
                if "texts" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["texts"] = []
                if "hotkeys" not in loaded_data["profiles"][profile]:
                    loaded_data["profiles"][profile]["hotkeys"] = []

            # **Sicherstellen, dass `active_profile` existiert**
            if "active_profile" not in loaded_data or loaded_data["active_profile"] not in loaded_data["profiles"]:
                if loaded_data["profiles"]:  # Falls Profile existieren
                    first_profile = list(loaded_data["profiles"].keys())[0]
                    print(f"⚠ `active_profile` existiert nicht mehr. Setze `{first_profile}` als neues aktives Profil.")
                    loaded_data["active_profile"] = first_profile
                else:  # Falls keine Profile mehr existieren
                    print("⚠ Keine Profile gefunden. Erstelle Standardprofil.")
                    loaded_data["profiles"] = {"Standard": {"titles": [], "texts": [], "hotkeys": []}}
                    loaded_data["active_profile"] = "Standard"

            return loaded_data


    except (FileNotFoundError, json.JSONDecodeError):
        return {
            "profiles": {
                "Standard": {"titles": [], "texts": [], "hotkeys": []}
            },
            "active_profile": list(loaded_data["profiles"].keys())[0] if loaded_data["profiles"] else "Standard"

        }


data = load_data()
active_profile = data.get("active_profile", list(data["profiles"].keys())[0])  # Erstes verfügbares Profil nehmen

#endregion

#region windows position 

def save_window_position():
    """Speichert die aktuelle Fensterposition & Größe in einer JSON-Datei."""
    if root:
        window_geometry = root.geometry()
        try:
            with open(WINDOW_CONFIG, "w") as file:
                json.dump({"geometry": window_geometry}, file)
        except Exception as e:
            print(f"⚠ Fehler beim Speichern der Fensterposition: {e}")

def load_window_position():
    """Lädt gespeicherte Fensterposition & gibt sie zurück."""
    try:
        with open(WINDOW_CONFIG, "r") as file:
            saved_config = json.load(file)
            if "geometry" in saved_config:
                return saved_config["geometry"]
    except (FileNotFoundError, json.JSONDecodeError) as e:
        print(f"⚠ Fensterposition konnte nicht geladen werden: {e}")
        return None  # Falls nichts gefunden wird, zurückgeben

#endregion

#region profiles

def save_profile_names():
    global data

    if not edit_mode:
        return  # Falls Bearbeitungsmodus nicht aktiv ist, nichts tun

    new_profiles = {}

    for old_name, entry in profile_entries.items():
        new_name = entry.get().strip()

        if new_name and new_name not in new_profiles:
            new_profiles[new_name] = data["profiles"].pop(old_name)  # **Alte Daten übernehmen**
        else:
            messagebox.showerror("Fehler", f"Profilname \"{new_name}\" ist ungültig oder bereits vergeben!")
            return

    data["profiles"] = new_profiles  # **Profilnamen aktualisieren**
    save_data()  # **Neue Namen speichern**
    update_ui()  # **UI aktualisieren**

def update_profile_buttons():
    """Hervorhebung des aktiven Profils durch andere Farbe."""
    for profile, button in profile_buttons.items():
        if profile == active_profile:
            button.config(bg="lightblue", font=("Arial", 10, "bold"))  # **Aktives Profil optisch hervorheben**
        else:
            button.config(bg="SystemButtonFace", font=("Arial", 10, "normal"))  # **Normale Darstellung**

def switch_profile(profile_name):
    global active_profile, data

    if profile_name not in data["profiles"]:
        print(f"⚠ Fehler: Profil {profile_name} existiert nicht!")
        return

    active_profile = profile_name
    update_profile_buttons()  # **Buttons visuell aktualisieren**
    update_ui()  # **Neues Profil laden**

#endregion

#region Hauptfenster ( main(): ) 

root = tk.Tk()
if sys.platform.startswith("win"): 
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")

root.iconbitmap("H.ico")  # Setze das Icon direkt nach dem Fenster
root.after(500, lambda: root.iconbitmap("H.ico"))
root.title("QuickPaste")

# Speichern der Position beim Schließen
root.protocol("WM_DELETE_WINDOW", lambda: [save_window_position(), minimize_to_tray()])


# Lade gespeicherte Fensterposition
saved_geometry = load_window_position()
if saved_geometry:
    root.geometry(saved_geometry)  # Falls gespeichert, setze diese Position
else:
    root.geometry("300x650")  # Falls keine gespeicherte Position existiert, Standardwert

# Lade gespeicherte Fensterposition beim Start
try:
    with open(WINDOW_CONFIG, "r") as file:
        saved_config = json.load(file)
        if "geometry" in saved_config:
            root.geometry(saved_config["geometry"])  
except FileNotFoundError:
    root.geometry("300x650")  # Standardgröße, falls keine gespeicherte Position vorhanden ist

#endregion

#region insert_text


def insert_text(index):
    global data  

    # **Falls die Pfeiltasten allein oder mit `ctrl+shift` gedrückt sind, breche ab**
    if keyboard.is_pressed("up") or keyboard.is_pressed("down") or keyboard.is_pressed("left") or keyboard.is_pressed("right"):
        if keyboard.is_pressed("ctrl") and keyboard.is_pressed("shift"):
            print(f"🔄 Windows-Funktion aktiv: `ctrl+shift+Pfeiltaste`, kein Text eingefügt.")
            return  

    print(f"🔍 insert_text() wurde aufgerufen mit Index: {index}")

    if active_profile not in data["profiles"]:
        print(f"🚨 Fehler: Aktives Profil '{active_profile}' nicht gefunden!")
        return  

    if "texts" not in data["profiles"][active_profile]:
        print(f"🚨 Fehler: 'texts' nicht in Profil '{active_profile}' vorhanden!")
        return  

    if index < 0 or index >= len(data["profiles"][active_profile]["texts"]):
        print(f"🚨 Fehler: Ungültiger Index {index}")
        return  

    text = data["profiles"][active_profile]["texts"][index]
    print(f"✅ Einfügen von: '{text}' mit Hotkey {index}")  

    pyperclip.copy(text)
    keyboard.send("ctrl+v")



#endregion

#region hotkeys

#OKKKKKK

def safe_insert_text(idx):
    # Verhindert Einfügen bei normalen Windows-Shortcuts
    if keyboard.is_pressed("ctrl") and keyboard.is_pressed("shift"):
        event_keys = keyboard._pressed_events.keys()
        arrow_keys = {72, 80, 75, 77}  # Scan-Codes für Pfeiltasten
        if any(scan_code in event_keys for scan_code in arrow_keys):
            print(f"🔄 Windows-Funktion aktiv: `ctrl+shift+Pfeiltaste`, kein Text eingefügt.")
            return  
    
    insert_text(idx)


def register_hotkeys():
    global data
    erlaubte_zeichen = "0123456789befhmpqvxzäöü§'^,.-$"
    belegte_hotkeys = set()
    fehlerhafte_hotkeys = False

    # **Bestehende Hotkeys bereinigen**
    for hotkey in list(keyboard._hotkeys.copy()):
        try:
            keyboard.remove_hotkey(hotkey)
        except KeyError:
            pass  

    # **Prüfen, ob Hotkeys vorhanden sind**
    if "hotkeys" not in data["profiles"].get(active_profile, {}):
        data["profiles"].setdefault(active_profile, {})["hotkeys"] = []

    erlaubte_kombinationen = {"ctrl", "shift"}  # Erlaubte Modifier

    for i, hotkey in enumerate(data["profiles"][active_profile]["hotkeys"]):
        if not hotkey.strip():
            continue  # **Unvollständige Hotkeys überspringen**

        keys = hotkey.lower().split("+")
        
        # **Sicherstellen, dass `keys` nicht leer ist, um IndexError zu vermeiden**
        if not keys or len(keys) < 2 or keys[-1] not in erlaubte_zeichen:
            messagebox.showerror("Fehler", f"Ungültiger Hotkey \"{hotkey}\" für Eintrag {i+1}. Erlaubt: 0-9, A-Z, §, ^, #")
            fehlerhafte_hotkeys = True  
            continue  

        # **Falls Hotkey nur `ctrl+shift` oder ungültige Kombination enthält, ignorieren**
        if set(keys) == {"ctrl", "shift"}:
            print(f"⚠ Hotkey '{hotkey}' ignoriert (nur Modifier-Kombination).")
            continue  

        # **Prüfen, ob es sich um eine gültige Hotkey-Kombination handelt**
        if not any(k in erlaubte_kombinationen for k in keys[:-1]):
            print(f"⚠ Hotkey '{hotkey}' ignoriert (ungültige Kombination).")
            continue

        # **Prüfen, ob der Hotkey bereits vergeben ist**
        if hotkey in belegte_hotkeys:
            messagebox.showerror("Fehler", f"Hotkey \"{hotkey}\" wird bereits verwendet!")
            fehlerhafte_hotkeys = True  # **Merke, dass es einen Fehler gab**
            continue  
        
        # **Hotkey registrieren**
        belegte_hotkeys.add(hotkey)
        keyboard.add_hotkey(hotkey, lambda i=i: safe_insert_text(i), suppress=True, trigger_on_release=True)

    return fehlerhafte_hotkeys  # **Gibt `True` zurück, wenn Fehler existieren**













#endregion

#region tray 


tray_icon = None 

def create_tray_icon():
    global tray_icon  
    if tray_icon is None:  
        image = Image.open("H.ico")
        tray_icon = Icon("QuickPaste", image, menu=Menu(
            MenuItem("Öffnen", show_window),
            MenuItem("Beenden", quit_application)
        ))
        tray_thread = threading.Thread(target=tray_icon.run, daemon=True)
        tray_thread.start()

def minimize_to_tray():
    root.withdraw()   

def show_window(icon, item):
    global tray_icon
    root.deiconify()  
    tray_icon = None  

def quit_application(icon, item):
    global tray_icon
    
    window_geometry = root.geometry() 
    with open("window_config.json", "w") as file:
        json.dump({"geometry": window_geometry}, file)

    if tray_icon is not None:
        try:
            tray_icon.stop()
        except Exception as e:
            print(f"Fehler beim Stoppen des Tray-Icons: {e}")  

    root.quit()
    root.destroy()
    sys.exit(0)

#endregion

#region add/del Entry


def add_new_entry():
    global data

    # Sicherstellen, dass das aktive Profil existiert
    if active_profile not in data["profiles"]:
        data["profiles"][active_profile] = {"titles": [], "texts": [], "hotkeys": []}

    # Sicherstellen, dass alle Listen existieren
    if "titles" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["titles"] = []
    if "texts" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["texts"] = []
    if "hotkeys" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["hotkeys"] = []

    # Neuen Eintrag hinzufügen
    data["profiles"][active_profile]["titles"].append("Neuer Eintrag")
    data["profiles"][active_profile]["texts"].append("Neuer Text")
    data["profiles"][active_profile]["hotkeys"].append("ctrl+shift+")  # **Jetzt Standardwert setzen!**

    update_ui()  # UI aktualisieren
    register_hotkeys()  # Hotkeys neu registrieren
   
def delete_entry(index):
    global data

    # Sicherstellen, dass das aktive Profil existiert
    if active_profile not in data["profiles"]:
        messagebox.showerror("Fehler", "Kein aktives Profil gefunden!")
        return

    # Sicherstellen, dass die Listen existieren
    if "titles" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["titles"] = []
    if "texts" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["texts"] = []
    if "hotkeys" not in data["profiles"][active_profile]:
        data["profiles"][active_profile]["hotkeys"] = []

    # Sicherstellen, dass der Index gültig ist
    if index < 0 or index >= len(data["profiles"][active_profile]["titles"]):
        messagebox.showerror("Fehler", "Ungültiger Eintrag zum Löschen ausgewählt!")
        return

    # Eintrag entfernen
    del data["profiles"][active_profile]["titles"][index]
    del data["profiles"][active_profile]["texts"][index]
    del data["profiles"][active_profile]["hotkeys"][index]

    update_ui()  # UI aktualisieren
    register_hotkeys()  # Hotkeys neu registrieren


#endregion

edit_mode = False
text_entries = []
title_entries = []
hotkey_entries = []

#region toggle edit mode

def toggle_edit_mode():
    global edit_mode
    edit_mode = not edit_mode  # Modus wechseln

    if not edit_mode:  # Falls der Bearbeitungsmodus verlassen wird
        save_profile_names()  # **Profilnamen speichern**

    update_ui()  # UI neu laden

#endregion

#region save_data

def save_data():
    global data, active_profile, profile_entries

    try:
        # **Profilnamen aktualisieren**
        updated_profiles = {}
        new_active_profile = active_profile  # Standardmäßig beibehalten

        for old_name, entry in profile_entries.items():
            new_name = entry.get().strip()

            # Falls der Name geändert wurde
            if new_name and new_name != old_name:
                if new_name in data["profiles"]:
                    messagebox.showerror("Fehler", f"Profilname '{new_name}' existiert bereits!")
                    return

                updated_profiles[new_name] = data["profiles"].pop(old_name)  # Umbenennen des Profils
                
                # Falls das aktive Profil umbenannt wurde, aktualisieren
                if active_profile == old_name:
                    new_active_profile = new_name
            else:
                updated_profiles[old_name] = data["profiles"][old_name]  # Unveränderte Profile übernehmen

        data["profiles"] = updated_profiles
        active_profile = new_active_profile  # Neues aktives Profil setzen

        # **Stelle sicher, dass `active_profile` existiert**
        if new_active_profile in updated_profiles:
            active_profile = new_active_profile  # Falls umbenannt, auf neuen Namen setzen
        else:
            active_profile = list(updated_profiles.keys())[0]  # Fallback: Erstes Profil wählen

        data["active_profile"] = active_profile  # Speichern des neuen aktiven Profils

        # **Textbausteine & Hotkeys speichern**
        data["profiles"][active_profile]["titles"] = [entry.get() for entry in title_entries if entry.winfo_exists()]
        data["profiles"][active_profile]["texts"] = [entry.get() for entry in text_entries if entry.winfo_exists()]
        data["profiles"][active_profile]["hotkeys"] = [entry.get() for entry in hotkey_entries if entry.winfo_exists()]

        # **Daten in die Datei speichern**
        with open(CONFIG_FILE, "w") as file:
            json.dump(data, file, indent=4)

        # **Hotkeys neu registrieren**
        fehlerhafte_hotkeys = register_hotkeys()

        # **Nur in den normalen Modus wechseln, wenn es keine Hotkey-Fehler gibt**
        if not fehlerhafte_hotkeys:
            toggle_edit_mode()

    except Exception as e:
        messagebox.showerror("Fehler", f"Speichern fehlgeschlagen: {e}")

#endregion

#region UI

def update_ui():
    global content_frame, profile_buttons, profile_entries, title_labels, scrollable_frame, canvas
    global active_profile
    profile_buttons = {}
    profile_entries = {}
    title_labels = []

    # **Falls `active_profile` nicht existiert, ein gültiges setzen**
    if active_profile not in data["profiles"]:
        active_profile = list(data["profiles"].keys())[0]  # Erstes verfügbares Profil wählen




    # **Nur relevante UI-Elemente entfernen, um Performance zu verbessern**
    for widget in root.winfo_children():
        if widget.winfo_name() not in ["frame_container", "top_frame"]:
            widget.destroy()

    # **Oberer Frame für Profile + Einstellungen**
    top_frame = tk.Frame(root, name="top_frame")
    top_frame.pack(fill="x", padx=5, pady=5)

    profiles = list(data["profiles"].keys())
    for profile in profiles:
        frame = tk.Frame(top_frame)
        frame.pack(side="left", padx=3)
        
        if edit_mode:
            entry = tk.Entry(frame, width=12, bg="white")
            entry.insert(0, profile)
            entry.pack(side="left")
            profile_entries[profile] = entry
            edit_btn = tk.Button(frame, text="🖊", command=lambda p=profile: switch_profile(p), width=2, bg="lightblue" if profile == active_profile else "SystemButtonFace")
            edit_btn.pack(side="left", padx=2)
        else:
            if active_profile not in data["profiles"]:
                active_profile = list(data["profiles"].keys())[0]  # Erstes verfügbares Profil setzen

            btn = tk.Button(
                frame,
                text=profile,
                command=lambda p=profile: switch_profile(p),
                bg="lightblue" if profile == active_profile else "SystemButtonFace",
                font=("Arial", 10, "bold") if profile == active_profile else ("Arial", 10, "normal")
            )

            btn.pack(side="left", padx=3)
            profile_buttons[profile] = btn

    settings_icon = ttk.Button(top_frame, text="⚙️", command=toggle_edit_mode, width=3)
    settings_icon.pack(side="right", padx=5, pady=5)

    frame_container = tk.Frame(root, name="frame_container")
    frame_container.pack(fill="both", expand=True)

    canvas = tk.Canvas(frame_container)
    scrollbar = tk.Scrollbar(frame_container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
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

    if active_profile not in data["profiles"]:
        print(f"⚠ Fehler: Profil '{active_profile}' existiert nicht! Setze Standard-Profil.")
        active_profile = list(data["profiles"].keys())[0]  # Erstes existierendes Profil wählen

    max_title_length = max(len(t) for t in data["profiles"][active_profile]["titles"])


    for i, title in enumerate(data["profiles"][active_profile]["titles"]):
        frame = tk.Frame(scrollable_frame)
        frame.pack(fill="x", expand=True, padx=5, pady=2)

        if edit_mode:
            title_entry = tk.Entry(frame, width=max_title_length)
            title_entry.insert(0, title)
            title_entry.pack(side="left", padx=5)
            title_entries.append(title_entry)
        else:
            tk.Label(frame, text=title, width=max_title_length, anchor="w").pack(side="left", padx=5)

        if edit_mode:
            text_entry = tk.Entry(frame)
            text_entry.insert(0, data["profiles"][active_profile]["texts"][i])
            text_entry.pack(side="left", fill="x", expand=True, padx=5)
            text_entries.append(text_entry)
        else:
            text_button = tk.Button(frame, text=data["profiles"][active_profile]["texts"][i], command=lambda i=i: insert_text(i), width=10, height=2)
            text_button.pack(side="left", fill="both", expand=True, padx=5, pady=2)

        if edit_mode:
            hotkey_entry = tk.Entry(frame, width=12)
            hotkey_entry.insert(0, data["profiles"][active_profile]["hotkeys"][i])
            hotkey_entry.pack(side="left", fill="x", expand=True, padx=5)
            hotkey_entries.append(hotkey_entry)
        else:
            tk.Label(frame, text=data["profiles"][active_profile]["hotkeys"][i], width=9, anchor="w", fg="gray").pack(side="left", padx=5)

        if edit_mode:
            delete_button = tk.Button(frame, text="❌", width=2, height=0, font=("Arial", 12), command=lambda i=i: delete_entry(i))
            delete_button.pack(side="left", padx=5)

    if edit_mode:
        buttons_frame = tk.Frame(root)
        buttons_frame.pack(fill="x", pady=5)

        save_button = tk.Button(buttons_frame, text="💾 Speichern", command=save_data, fg="white", bg="green", height=2)
        save_button.pack(side="left", expand=True, fill="x", padx=5)
        add_button = tk.Button(buttons_frame, text="➕ Neuen Eintrag", command=add_new_entry, height=2)
        add_button.pack(side="right", expand=True, fill="x", padx=5)
    
    root.update_idletasks()

    print(f"✅ {len(title_entries)} Titel-Elemente hinzugefügt")
    print(f"✅ {len(text_entries)} Text-Elemente hinzugefügt")
    print(f"✅ {len(hotkey_entries)} Hotkey-Elemente hinzugefügt")

#endregion


data = load_data()
active_profile = data.get("active_profile", "Standard")
update_ui()
register_hotkeys()  
create_tray_icon()   
root.mainloop()

#region


#endregion
