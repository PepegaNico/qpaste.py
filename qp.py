import tkinter as tk
from tkinter import ttk
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

default_data = {
    "titles": ["Titel"] * 10,  # Start mit 10 Einträgen
    "texts": [f"Text {i+1}" for i in range(10)],
    "hotkeys": [
        "ctrl+shift+1", "ctrl+shift+2", "ctrl+shift+3", "ctrl+shift+4", "ctrl+shift+5",
        "ctrl+shift+6", "ctrl+shift+7", "ctrl+shift+8", "ctrl+shift+9", "ctrl+shift+0"
    ]
}

def load_data():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as file:
                return json.load(file)
        except json.JSONDecodeError:
            return default_data

data = load_data()

def load_window_geometry():
    if os.path.exists(WINDOW_CONFIG):
        try:
            with open(WINDOW_CONFIG, "r") as file:
                config = json.load(file)
                return config.get("geometry")
        except json.JSONDecodeError:
            pass
    return None



def insert_text(index):
    text = data["texts"][index]
    pyperclip.copy(text)
    time.sleep(0.1)
    keyboard.send("ctrl+v")

#region Hotkeys
def register_hotkeys():
    # Alle existierenden Hotkeys sicher entfernen
    for hotkey in list(keyboard._hotkeys.copy()):
        try:
            keyboard.remove_hotkey(hotkey)
        except KeyError:
            pass  # Ignoriere Fehler, falls der Hotkey nicht existiert

    # Neue Hotkeys registrieren
    for i, hotkey in enumerate(data["hotkeys"]):
        if hotkey.strip():  # Ignoriere leere Strings
            keyboard.add_hotkey(hotkey, insert_text, args=[i], suppress=True)

#endregion

tray_icon = None 

#region Tray
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
    
    # Fenstergröße und Position speichern
    window_geometry = root.geometry()  # Holt z. B. "800x600+100+50"
    with open("window_config.json", "w") as file:
        json.dump({"geometry": window_geometry}, file)

    # Tray-Icon sicher beenden
    if tray_icon is not None:
        try:
            tray_icon.stop()
        except Exception as e:
            print(f"Fehler beim Beenden des Tray-Icons: {e}")

    root.quit()
    root.destroy()
    sys.exit(0)

#endregion

#region add/del entry

    #add entry
def add_new_entry():
    """Fügt einen neuen leeren Eintrag im Bearbeitungsmodus hinzu, ohne den Modus zu verlassen."""
    data["titles"].append("Neuer Eintrag")
    data["texts"].append("Neuer Text")
    data["hotkeys"].append("ctrl+shift+")  # Neuer Eintrag mit leerem Hotkey

    update_ui()  # UI sofort aktualisieren, aber nicht speichern oder Hotkeys registrieren


    #del entry
def delete_entry(index):
    """Löscht einen Eintrag im Bearbeitungsmodus, ohne die UI zu verlassen."""
    if len(data["titles"]) > 1:  # Mindestens 1 Eintrag muss immer bleiben
        del data["titles"][index]
        del data["texts"][index]
        del data["hotkeys"][index]

        update_ui()  # UI aktualisieren, aber nicht speichern oder Hotkeys registrieren
#endregion

#region UI
root = tk.Tk()
saved_geometry = load_window_geometry()
root.geometry(saved_geometry if saved_geometry else "300x650")  # Falls keine Datei existiert, Standardgröße nutzen

if sys.platform.startswith("win"):  # Nur für Windows
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("QuickPaste")
root.iconbitmap("H.ico")
root.title("QuickPaste")
root.protocol("WM_DELETE_WINDOW", lambda: minimize_to_tray())

#Taskbar
def force_taskbar_icon():
    """Erzwingt das Laden des Taskleisten-Icons nach einer Verzögerung."""
    root.after(500, lambda: root.iconbitmap("H.ico"))  # Verzögertes Laden nach 500ms

edit_mode = False
text_entries = []
title_entries = []
hotkey_entries = []

def toggle_edit_mode():
    global edit_mode
    edit_mode = not edit_mode
    update_ui()

def save_data():
    global data

    # Stelle sicher, dass die Widgets noch existieren, bevor wir sie abrufen
    try:
        data["titles"] = [entry.get() for entry in title_entries if entry.winfo_exists()]
        data["texts"] = [entry.get() for entry in text_entries if entry.winfo_exists()]
        data["hotkeys"] = [entry.get() for entry in hotkey_entries if entry.winfo_exists()]
    except Exception as e:
        print(f"Fehler beim Speichern der Daten: {e}")
        return

    # Speichern in die JSON-Datei
    with open(CONFIG_FILE, "w") as file:
        json.dump(data, file, indent=4)

    # Nach dem Speichern den Bearbeitungsmodus verlassen
    toggle_edit_mode()
    register_hotkeys()

def update_ui():
    # Alle Widgets löschen
    for widget in root.winfo_children():
        widget.destroy()

    # **Fixierter Oberer Frame für den Einstellungsbutton**
    top_frame = tk.Frame(root)
    top_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)

    # **Einstellungsbutton: Wechselt zwischen Bearbeitungs- und Normalmodus** 
    settings_icon = ttk.Button(top_frame, text="⚙️" if not edit_mode else " ↩️",
                               command=toggle_edit_mode, width=3)
    settings_icon.pack(side="right", padx=5, pady=5)

    # **Haupt-Frame für Inhalte**
    # **Haupt-Frame mit Scrollbar**
    # **Scrollbarer Bereich für die Einträge**
    # **Scrollbarer Bereich für die Einträge**
    # **Scrollbarer Bereich für die Einträge**
    frame_container = tk.Frame(root)  # Container hält Canvas & Scrollbar
    frame_container.grid(row=1, column=0, sticky="nsew")

    canvas = tk.Canvas(frame_container)
    scrollbar = tk.Scrollbar(frame_container, orient="vertical", command=canvas.yview)

    scrollable_frame = tk.Frame(canvas)  # Das eigentliche UI-Frame

    # Scrollregion immer aktualisieren, wenn sich das UI verändert
    def update_scroll_region(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))

    scrollable_frame.bind("<Configure>", update_scroll_region)

    # Fenster in Canvas setzen
    window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=frame_container.winfo_width())

    canvas.configure(yscrollcommand=scrollbar.set)

    # Layout setzen
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Fenstergröße anpassen, damit das Layout korrekt dargestellt wird
    def resize_canvas(event):
        canvas.itemconfig(window, width=event.width)

    canvas.bind("<Configure>", resize_canvas)

    # Mausrad-Scrolling aktivieren (nur wenn sich die Maus über dem Canvas befindet)
    def _on_mouse_wheel(event):
        if root.winfo_height() < scrollable_frame.winfo_height():
            canvas.yview_scroll(-1 * (event.delta // 120), "units")

    canvas.bind("<Enter>", lambda e: root.bind_all("<MouseWheel>", _on_mouse_wheel))
    canvas.bind("<Leave>", lambda e: root.unbind_all("<MouseWheel>"))



    # Konfiguration für dynamische Größenanpassung
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)


    # **Liste mit Einträgen**
    for i, title in enumerate(data["titles"]):
        frame = tk.Frame(scrollable_frame)

        frame.pack(fill="x", expand=True, padx=5, pady=2, anchor="w")


        # **Titel-Feld**
        if edit_mode:
            title_entry = tk.Entry(frame, width=6,)
            title_entry.insert(0, title)
            title_entry.pack(side="left", fill="x", expand=True, padx=5)
            title_entries.append(title_entry)
        else:
            tk.Label(frame, text=title, width=6, anchor="w").pack(side="left", padx=5)

        # **Textfeld oder Button**
        if edit_mode:
            text_entry = tk.Entry(frame, width=10,)
            text_entry.insert(0, data["texts"][i])
            text_entry.pack(side="left", fill="x", expand=True, padx=5)
            text_entries.append(text_entry)
        else:
            text_button = tk.Button(frame, text=data["texts"][i], command=lambda i=i: insert_text(i), width=10,
                                    height=2)
            text_button.pack(side="left", fill="x", expand=True, padx=5)

        # **Hotkey-Feld**
        if edit_mode:
            hotkey_entry = tk.Entry(frame, width=9,)
            hotkey_entry.insert(0, data["hotkeys"][i])
            hotkey_entry.pack(side="left", fill="x", expand=True, padx=5)
            hotkey_entries.append(hotkey_entry)
        else:
            tk.Label(frame, text=data["hotkeys"][i], width=9, anchor="w", fg="gray").pack(side="left", padx=5)

        # **Löschen-Button (Nur im Bearbeitungsmodus)**
        if edit_mode:
            delete_button = tk.Button(frame, text="❌", width=2, height=0, font=("Arial", 12),
                                      command=lambda i=i: delete_entry(i))
            delete_button.pack(side="left", padx=5)

    # **Speichern & Neuer Eintrag-Button (Nur im Bearbeitungsmodus)**
    if edit_mode:
        buttons_frame = tk.Frame(root)
        buttons_frame.grid(row=2, column=0, sticky="ew", pady=5)

        save_button = tk.Button(buttons_frame, text="💾 Speichern", command=save_data, fg="white", bg="green", height=2)
        save_button.pack(side="left", expand=True, fill="x", padx=5)

        add_button = tk.Button(buttons_frame, text="➕ Neuen Eintrag", command=add_new_entry, height=2)
        add_button.pack(side="right", expand=True, fill="x", padx=5)

    # **Fenster aktualisieren**
    root.update_idletasks()










#endregion 

update_ui()
register_hotkeys()  
create_tray_icon()  
root.mainloop()
