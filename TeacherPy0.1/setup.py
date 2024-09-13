import os
import sys
import subprocess
import json
import time
import tkinter as tk
from tkinter import messagebox
import colorama
from colorama import Fore, Style

# Initialisiere colorama für farbige Ausgabe
colorama.init(autoreset=True)

# Ermittle den absoluten Pfad zum aktuellen Skript
current_script_path = os.path.abspath(__file__)

# Ermittle das Verzeichnis, in dem sich das aktuelle Skript befindet
current_dir = os.path.dirname(current_script_path)

# Definiere den Pfad für die Konfigurationsdatei
PROJECT_CONFIG_PATH = os.path.join(current_dir, 'project_config.json')

# Definiere die Platzhalter
PLACEHOLDERS = ["<PFAD_ZUM_BASISORDNER>", "<PFAD_ZUM_NOTIZORDNER>", "<PFAD_ZUM_USB_LAUFWERK>", "Pfad\\zum\\Ordner"]

def read_requirements():
    requirements_path = os.path.join(current_dir, 'requirements.txt')
    with open(requirements_path, 'r') as f:
        return [line.strip() for line in f if line.strip() and not line.startswith('#')]

def install(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
    except subprocess.CalledProcessError:
        print(f"{Fore.RED}Fehler beim Installieren von {package}.")

def ensure_basic_packages():
    basic_packages = ['setuptools', 'wheel', 'pip']
    for package in basic_packages:
        install(package)

def install_missing_libraries():
    required = read_requirements()
    for package in required:
        try:
            install(package)
            print(f"{Fore.GREEN}{package} wurde erfolgreich installiert.")
        except Exception as e:
            print(f"{Fore.RED}Fehler beim Installieren von {package}: {str(e)}")
    print(f"{Fore.GREEN}Installation abgeschlossen.")

def create_default_config():
    return {
        "base_folder_path": "Pfad\\zum\\Ordner",
        "notes_folder_path": "Pfad\\zum\\Ordner",
        "usb_path": "Pfad\\zum\\Ordner",
        "create_note": True,
        "svp_options": {
            "time_slots": {
                "1": ["8:15", "9:00"],
                "2": ["9:10", "10:40"],
                "3": ["11:00", "12:30"],
                "4": ["13:30", "15:00"]
            },
            "prompts": {
                "Klasse/Kurs": True,
                "Raumnummer": True,
                "Datum": True,
                "Lernbereich": True,
                "Stundenthema": True,
                "Lehrperson": True
            },
            "default_teacher": "<STANDARD_LEHRERNAME>",
            "include_abbreviations": True
        }
    }

def is_valid_directory(path):
    return os.path.isabs(path) and os.path.exists(os.path.dirname(path))

def update_config():
    if not os.path.exists(PROJECT_CONFIG_PATH):
        config = create_default_config()
    else:
        with open(PROJECT_CONFIG_PATH, 'r') as f:
            config = json.load(f)

    def update_path(key, prompt):
        current_value = config.get(key, "")
        if current_value in PLACEHOLDERS or current_value == "":
            print(f"{Fore.YELLOW}Es wurde noch kein {prompt} festgelegt.")
            while True:
                new_value = input(f"{Fore.CYAN}Bitte geben Sie den {prompt} ein: {Style.RESET_ALL}")
                if is_valid_directory(new_value):
                    return new_value
                else:
                    print(f"{Fore.RED}Ungültiger Pfad. Bitte geben Sie einen gültigen absoluten Pfad ein.")
        else:
            print(f"{Fore.GREEN}Aktueller {prompt}: {current_value}")
            change = input(f"{Fore.CYAN}Möchten Sie den {prompt} ändern? (j/n): {Style.RESET_ALL}").lower()
            if change == 'j':
                while True:
                    new_value = input(f"{Fore.CYAN}Neuer {prompt}: {Style.RESET_ALL}")
                    if is_valid_directory(new_value):
                        return new_value
                    else:
                        print(f"{Fore.RED}Ungültiger Pfad. Bitte geben Sie einen gültigen absoluten Pfad ein.")
            else:
                return current_value

    config["base_folder_path"] = update_path("base_folder_path", "Pfad zum Basisordner")
    config["usb_path"] = update_path("usb_path", "Pfad zum USB-Laufwerk")

    use_notes = input(f"{Fore.CYAN}Möchten Sie einen Notizordner verwenden? (j/n): {Style.RESET_ALL}").lower() in ['j', 'ja', 'y', 'yes']
    config["create_note"] = use_notes
    if use_notes:
        config["notes_folder_path"] = update_path("notes_folder_path", "Pfad zum Notizordner")
    else:
        config["notes_folder_path"] = ""

    config["svp_options"]["default_teacher"] = input(f"{Fore.CYAN}Standard-Lehrername: {Style.RESET_ALL}") or config["svp_options"].get("default_teacher", "")

    include_abbreviations = input(f"{Fore.CYAN}Möchten Sie Abkürzungen einschließen? (j/n): {Style.RESET_ALL}").lower() in ['j', 'ja', 'y', 'yes']
    config["svp_options"]["include_abbreviations"] = include_abbreviations

    change_time_slots = input(f"{Fore.CYAN}Möchten Sie die Zeitslots ändern? (j/n): {Style.RESET_ALL}").lower() in ['j', 'ja', 'y', 'yes']
    if change_time_slots:
        for slot in config["svp_options"]["time_slots"]:
            print(f"{Fore.GREEN}Aktueller Zeitslot {slot}: {config['svp_options']['time_slots'][slot][0]} - {config['svp_options']['time_slots'][slot][1]}")
            start_time = input(f"{Fore.CYAN}Neue Startzeit für Slot {slot} (leer lassen für unverändert): {Style.RESET_ALL}")
            end_time = input(f"{Fore.CYAN}Neue Endzeit für Slot {slot} (leer lassen für unverändert): {Style.RESET_ALL}")
            if start_time:
                config["svp_options"]["time_slots"][slot][0] = start_time
            if end_time:
                config["svp_options"]["time_slots"][slot][1] = end_time

    with open(PROJECT_CONFIG_PATH, 'w') as f:
        json.dump(config, f, indent=4)

    print(f"{Fore.GREEN}Konfiguration wurde erfolgreich aktualisiert.")

def open_and_grant_permissions_to_templates():
    template_dir = os.path.join(current_dir, 'Vorlage')
    template_file = 'SVP_Vorlage.dotx'
    template_path = os.path.join(template_dir, template_file)

    if not os.path.exists(template_path):
        print(f"{Fore.RED}Die Datei {template_file} wurde nicht im Verzeichnis {template_dir} gefunden.")
        return

    print(f"{Fore.YELLOW}WICHTIGER HINWEIS:")
    print(f"{Fore.CYAN}Bitte öffnen Sie manuell die Datei {template_file} im folgenden Ordner:")
    print(f"{Fore.GREEN}{template_dir}")
    print(f"{Fore.CYAN}Erteilen Sie die Berechtigungen zum Bearbeiten und speichern Sie die Datei.")
    
    while True:
        user_input = input(f"{Fore.YELLOW}Haben Sie die Berechtigungen erteilt und die Datei gespeichert? (ja/nein): {Style.RESET_ALL}").lower()
        if user_input in ['ja', 'j', 'yes', 'y']:
            print(f"{Fore.GREEN}Vielen Dank! Das Skript wird fortgesetzt.")
            break
        elif user_input in ['nein', 'n', 'no']:
            print(f"{Fore.RED}Bitte erteilen Sie die Berechtigungen, bevor Sie fortfahren.")
        else:
            print(f"{Fore.YELLOW}Ungültige Eingabe. Bitte antworten Sie mit 'ja' oder 'nein'.")

    print(f"{Fore.GREEN}Vorlagendatei wurde verarbeitet.")

def create_shortcut():
    try:
        import winshell
        from win32com.client import Dispatch
    except ImportError:
        print(f"{Fore.RED}Konnte winshell oder win32com nicht importieren. Verknüpfung wird nicht erstellt.")
        return

    main_path = os.path.join(current_dir, 'main.py')
    icon_path = os.path.join(current_dir, 'icon.ico')
    shortcut_path = os.path.join(current_dir, 'TeacherPy.lnk')

    if not os.path.exists(main_path):
        print(f"{Fore.RED}main.py nicht gefunden.")
        return
    if not os.path.exists(icon_path):
        print(f"{Fore.RED}icon.ico nicht gefunden.")
        return

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = sys.executable
    shortcut.Arguments = f'"{main_path}"'
    shortcut.WorkingDirectory = current_dir
    shortcut.IconLocation = icon_path
    shortcut.save()

    print(f"{Fore.GREEN}Verknüpfung erstellt: {shortcut_path}")

def main():
    print(f"{Fore.CYAN}Überprüfe und installiere grundlegende Pakete...")
    ensure_basic_packages()

    print(f"{Fore.CYAN}Installiere erforderliche Bibliotheken...")
    install_missing_libraries()

    print(f"{Fore.CYAN}Öffne Vorlagendateien und erteile Berechtigungen...")
    open_and_grant_permissions_to_templates()

    print(f"{Fore.CYAN}Aktualisiere die Konfiguration...")
    update_config()

    print(f"{Fore.CYAN}Erstelle die Verknüpfung...")
    create_shortcut()

    print(f"{Fore.GREEN}Setup abgeschlossen.")

if __name__ == "__main__":
    main()
