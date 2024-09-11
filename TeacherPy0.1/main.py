import os
import sys
import json
import subprocess
from colorama import init, Fore

init(autoreset=True)  # Initialize colorama

# Füge den scripts-Ordner zum Python-Pfad hinzu
SCRIPT_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(SCRIPT_BASE_DIR, "scripts")
sys.path.append(SCRIPTS_DIR)

# Importiere die Pfade aus paths.py
from paths import NEU_SCRIPT_PATH, FINAL_SCRIPT_PATH, ARCHIVE_SCRIPT_PATH

def load_config(file_name):
    json_path = os.path.join(SCRIPT_BASE_DIR, file_name)
    
    if not os.path.exists(json_path):
        print(f"Fehler: Die Konfigurationsdatei '{json_path}' wurde nicht gefunden.")
        exit(1)

    try:
        with open(json_path, 'r') as config_file:
            return json.load(config_file)
    except json.JSONDecodeError:
        print(f"Fehler: Die Datei '{json_path}' enthält kein gültiges JSON-Format.")
        exit(1)

def colored_print(text, color):
    print(color + text)

def main():
    config = load_config('project_config.json')
    layout = load_config('layout.json')

    while True:
        print("\nWählen Sie eine Option:")
        colored_print("1. Eine neue Stunde erstellen", Fore.GREEN)
        colored_print("2. Eine fertiggestellte Stunde finalisieren", Fore.BLUE)
        colored_print("3. Eine abgeschlossene Stunde archivieren", Fore.YELLOW)
        colored_print("4. Beenden", Fore.RED)
        
        choice = input("Ihre Wahl (1-4): ").strip()
        
        if choice == '1':
            subprocess.run([sys.executable, NEU_SCRIPT_PATH])
        elif choice == '2':
            subprocess.run([sys.executable, FINAL_SCRIPT_PATH])
        elif choice == '3':
            subprocess.run([sys.executable, ARCHIVE_SCRIPT_PATH])
        elif choice == '4':
            print("Programm wird beendet.")
            break
        else:
            print("Ungültige Eingabe. Bitte wählen Sie 1, 2, 3 oder 4.")

if __name__ == "__main__":
    main()