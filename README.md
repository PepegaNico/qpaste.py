# QuickPaste 

QuickPaste ist eine Windows-Anwendung, die es ermöglicht, vordefinierte Texte per Tastenkombination (Hotkeys) oder über eine grafische Benutzeroberfläche (GUI) einzufügen. Die Anwendung ist ideal für den schnellen Zugriff auf Standardtexte.

## Features

- **Hotkey-Unterstützung**: Füge Texte mit Strg+Shift+Zahl ein.
- **Eigene Texte & Hotkeys verwalten**: Texte und Hotkeys können in der Anwendung geändert und gespeichert werden.

## Installation

1. Stelle sicher, dass Python 3.10 oder höher installiert ist.
2. Installiere die benötigten Abhängigkeiten mit folgendem Befehl:
   ```sh
   pip install pyperclip keyboard
   ```

## Nutzung

1. Starte das Skript mit:
   ```sh
   python qp.py
   ```
2. Die Anwendung öffnet sich und ist einsatzbereit.
3. Texte können über die GUI oder definierte Hotkeys eingefügt werden.

## Hotkeys (Standardmäßig gesetzt)

```
Strg + Shift + 1  → Text 1 einfügen
Strg + Shift + 2  → Text 2 einfügen
...
Strg + Shift + Q  → Text 11 einfügen
Strg + Shift + W  → Text 12 einfügen
```

Die Hotkeys können in der Anwendung angepasst werden.

## Fehlerbehebung

Falls die Hotkeys nicht funktionieren:

1. Prüfe, ob die `keyboard`-Bibliothek korrekt installiert ist:
   ```sh
   python -c "import keyboard; print(keyboard.__version__)"
   ```
2. Falls nötig, reinstalliere `keyboard`:
   ```sh
   pip install --force-reinstall keyboard
   ```



