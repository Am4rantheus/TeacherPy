import os
import sys
import shutil
import filecmp
import json
from colorama import init, Fore, Style

# Füge den scripts-Ordner zum Python-Pfad hinzu
SCRIPT_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR = os.path.dirname(SCRIPT_BASE_DIR)
sys.path.append(SCRIPT_BASE_DIR)

# Importiere die Pfade aus paths.py (falls benötigt)
# from paths import SOME_PATH

# Initialize colorama
init(autoreset=True)

def load_config(file_name):
    config_path = os.path.join(PARENT_DIR, file_name)

    try:
        with open(config_path, 'r') as config_file:
            return json.load(config_file)
    except FileNotFoundError:
        print(f"Fehler: Die Konfigurationsdatei wurde nicht gefunden: {config_path}")
        print(f"Bitte stellen Sie sicher, dass die '{file_name}'-Datei im übergeordneten Verzeichnis des Skripts liegt.")
        exit(1)
    except json.JSONDecodeError:
        print(f"Fehler: Die Datei '{config_path}' enthält kein gültiges JSON-Format.")
        exit(1)

config = load_config('project_config.json')
layout = load_config('layout.json')

BASE_DIR = config['base_folder_path']
USB_PATH = config['usb_path']

def get_color(key):
    return getattr(Fore, layout['colors'].get(key, 'WHITE'))

def colored_filename(filename):
    if filename.startswith("SVP_"):
        return f"{get_color('SVP_') if not filename.endswith('.pdf') else get_color('SVP_PDF')}{filename}{Style.RESET_ALL}"
    elif filename.startswith("PPP_") or filename.endswith((".pptx", ".odp")):
        return f"{get_color('PRESENTATION')}{filename}{Style.RESET_ALL}"
    elif filename.endswith((".docx", ".odt")):
        return f"{get_color('DOCX')}{filename}{Style.RESET_ALL}"
    elif filename.endswith(".pdf"):
        if "_Erwartungsbild" in filename:
            return f"{get_color('ERWARTUNGSBILD')}{filename}{Style.RESET_ALL}"
        else:
            return f"{get_color('PDF')}{filename}{Style.RESET_ALL}"
    elif filename.endswith((".mp4", ".avi", ".mov", ".wmv")):
        return f"{get_color('VIDEO')}{filename}{Style.RESET_ALL}"
    else:
        for prefix in ['AB_', 'LB_']:
            if filename.startswith(prefix):
                return f"{get_color(prefix)}{filename}{Style.RESET_ALL}"
    return filename

def colored_foldername(foldername):
    return f"{get_color('FOLDER')}{foldername}{Style.RESET_ALL}"

def get_subfolder(base_path):
    subfolders = [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]
    print("Verfügbare Unterordner:")
    for i, folder in enumerate(subfolders, 1):
        print(f"{i}. {colored_foldername(folder)}")
    while True:
        try:
            choice = int(input("Wählen Sie einen Unterordner (Nummer): ")) - 1
            if 0 <= choice < len(subfolders):
                return os.path.join(base_path, subfolders[choice])
            else:
                print("Ungültige Auswahl. Bitte versuchen Sie es erneut.")
        except ValueError:
            print("Bitte geben Sie eine Zahl ein.")

def list_files(folder):
    print(f"\nDateien in {colored_foldername(os.path.basename(folder))}:")
    for root, _, files in os.walk(folder):
        for file in files:
            print(colored_filename(file))

def compare_and_sync(usb_folder, base_folder, option):
    for root, _, files in os.walk(usb_folder):
        for file in files:
            usb_file = os.path.join(root, file)
            base_file = os.path.join(base_folder, os.path.relpath(usb_file, usb_folder))
            
            if option == 1:  # Veränderte Dateien überschreiben
                if os.path.exists(base_file):
                    if not filecmp.cmp(usb_file, base_file, shallow=False):
                        shutil.copy2(usb_file, base_file)
                        print(f"Aktualisiert: {colored_filename(file)}")
                else:
                    shutil.copy2(usb_file, base_file)
                    print(f"Neu hinzugefügt: {colored_filename(file)}")
            elif option == 2:  # Nur neue Dateien hinzufügen
                if not os.path.exists(base_file):
                    shutil.copy2(usb_file, base_file)
                    print(f"Neu hinzugefügt: {colored_filename(file)}")

def archive_folder(source_folder, archive_folder):
    shutil.copytree(source_folder, archive_folder, dirs_exist_ok=True)
    print(f"Ordner archiviert: {colored_foldername(os.path.basename(archive_folder))}")

def main():
    print(f"Wählen Sie den zu archivierenden Ordner auf dem USB-Stick ({colored_foldername(os.path.basename(USB_PATH))}):")
    usb_folder = get_subfolder(USB_PATH)
    base_folder = os.path.join(BASE_DIR, os.path.basename(usb_folder))

    list_files(usb_folder)

    print("\nWählen Sie eine Option:")
    print("1. Veränderte Dateien auf dem USB-Stick in Basisordner archivieren (überschreiben)")
    print("2. Nur neue Dateien auf dem USB-Stick in Basisordner archivieren")
    print("3. Gesamten Ordner archivieren")
    
    while True:
        choice = input("Ihre Wahl (1-3): ").strip()
        if choice in ['1', '2', '3']:
            break
        print("Ungültige Eingabe. Bitte wählen Sie 1, 2 oder 3.")

    if choice in ['1', '2']:
        compare_and_sync(usb_folder, base_folder, int(choice))
    elif choice == '3':
        archive_folder(usb_folder, base_folder)

    keep_usb = input("Möchten Sie die Dateien auf dem USB-Stick behalten? (j/n): ").lower()
    if keep_usb != 'j':
        shutil.rmtree(usb_folder)
        print(f"Ordner auf dem USB-Stick gelöscht: {colored_foldername(os.path.basename(usb_folder))}")

if __name__ == "__main__":
    main()