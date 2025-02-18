## **QuickPaste – Fast Clipboard Tool for Windows** 

## Description

QuickPaste is a Windows application that allows you to insert predefined text using keyboard shortcuts (hotkeys) or a graphical user interface (GUI).  
The application is ideal for quickly accessing frequently used text snippets.  

## Features  

- **Hotkey Support**: Insert text using `Ctrl + Shift + Number`.  
- **Manage Custom Text & Hotkeys**: Modify and save text snippets and hotkeys within the application.  
- **Add New Entries**: Easily add new text snippets via the GUI.  
- **Delete Entries**: Remove unused text entries with a single click.  

## Installation  

1. Ensure that **Python 3.10 or later** is installed.  
2. Install the required dependencies with the following command:  
   ```sh
   pip install pyperclip keyboard
   ```

## Usage

1. Start the script with:
   ```sh
   python qp.py
   ```
2. The application will launch and be ready to use.
3. Text can be inserted via the GUI or predefined hotkeys.

## Default Hotkeys

```
Strg + Shift + 1  → Insert Text 1
Strg + Shift + 2  → Insert Text 2
```

Hotkeys can be customized within the application.

## Troubleshooting

### Hotkeys not working?  
- Ensure **no other application is using the same shortcuts**
- Check if the `keyboard` library is installed correctly:
   ```sh
   python -c "import keyboard; print(keyboard.__version__)"
   ```

- If necessary, reinstall `keyboard`:
   ```sh
   pip install --force-reinstall keyboard
   ```



If you encounter bugs or have suggestions, please create an issue on GitHub:  
[GitHub Issues](https://github.com/PepegaNico/qpaste/issues)  




