import tkinter as tk
from tkinter import ttk
import pyperclip
import keyboard
import json
import os
import time

CONFIG_FILE = "config.json"

default_data = { 
    "hotkeys": [
        "ctrl+shift+1", "ctrl+shift+2", "ctrl+shift+3", "ctrl+shift+4", "ctrl+shift+5",
        "ctrl+shift+6", "ctrl+shift+7", "ctrl+shift+8", "ctrl+shift+9", "ctrl+shift+0",
        "ctrl+shift+q", "ctrl+shift+w", "ctrl+shift+e", "ctrl+shift+r", "ctrl+shift+t",
        "ctrl+shift+y", "ctrl+shift+u", "ctrl+shift+i", "ctrl+shift+o", "ctrl+shift+p"
    ]
}

def load_data():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as file:
                data = json.load(file)
                return data
        except json.JSONDecodeError:
            return default_data

data = load_data()

root = tk.Tk()
root.iconbitmap("H.ico")
root.title("QuickPaste")

edit_mode = False
text_entries = []
title_entries = []
hotkey_entries = []

def insert_text(index):
    text = data["texts"][index]
    pyperclip.copy(text)
    time.sleep(0.1)
    keyboard.send("ctrl+v")

def toggle_edit_mode():
    global edit_mode
    edit_mode = not edit_mode
    update_ui()

def save_data():
    global data

    data["titles"] = [entry.get() for entry in title_entries]
    data["texts"] = [entry.get() for entry in text_entries]
    data["hotkeys"] = [entry.get() for entry in hotkey_entries]

    with open(CONFIG_FILE, "w") as file:
        json.dump(data, file, indent=4)

    register_hotkeys()
    toggle_edit_mode()

def register_hotkeys():
    global data

    current_hotkeys = list(data["hotkeys"])

    for hotkey in current_hotkeys:
        try:
            keyboard.remove_hotkey(hotkey)
        except KeyError:
            pass 

    for i in range(20):
        keyboard.add_hotkey(data["hotkeys"][i], lambda i=i: insert_text(i))

def update_ui():
    for widget in root.winfo_children():
        widget.destroy()

    settings_icon = ttk.Button(root, text="⚙️", command=toggle_edit_mode, width=3)
    settings_icon.grid(row=0, column=1, sticky="e", padx=5, pady=5)

    global title_entries, text_entries, hotkey_entries
    title_entries = []
    text_entries = []
    hotkey_entries = []

    for i in range(20):
        frame = tk.Frame(root)
        frame.grid(row=i+1, column=0, padx=5, pady=2, sticky="w")

        if edit_mode:
            title_entry = tk.Entry(frame, width=15)
            title_entry.insert(0, data["titles"][i])
            title_entry.pack(side="left", padx=5)
            title_entries.append(title_entry)

            text_entry = tk.Entry(frame, width=30)
            text_entry.insert(0, data["texts"][i])
            text_entry.pack(side="left", padx=5)
            text_entries.append(text_entry)

            hotkey_entry = tk.Entry(frame, width=15)
            hotkey_entry.insert(0, data["hotkeys"][i])
            hotkey_entry.pack(side="left", padx=5)
            hotkey_entries.append(hotkey_entry)
        else:

            title_label = tk.Label(frame, text=f"{data['titles'][i]}", width=15, anchor="w")
            title_label.pack(side="left", padx=5)

            btn = tk.Button(frame, text=data["texts"][i], command=lambda i=i: insert_text(i), width=30, height=1)
            btn.pack(side="left", padx=5)

            hotkey_label = tk.Label(frame, text=f"{data['hotkeys'][i]}", width=15, anchor="w", fg="gray")
            hotkey_label.pack(side="left")

    if edit_mode:
        save_button = tk.Button(root, text="Speichern", command=save_data, fg="white", bg="green")
        save_button.grid(row=21, column=0, pady=10)

register_hotkeys()
update_ui()
root.mainloop()