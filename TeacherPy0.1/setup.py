import os
import sys
import subprocess
import json

# Ermittle den absoluten Pfad zum aktuellen Skript
current_script_path = os.path.abspath(__file__)

# Ermittle das Verzeichnis, in dem sich das aktuelle Skript befindet
current_dir = os.path.dirname(current_script_path)

# Definiere den Pfad für die Konfigurationsdatei
PROJECT_CONFIG_PATH = os.path.join(current_dir, 'project_config.json')

# Definiere die Platzhalter
PLACEHOLDERS = ["<PFAD_ZUM_BASISORDNER>", "<PFAD_ZUM_NOTIZORDNER>", "<PFAD_ZUM_USB_LAUFWERK>", "Pfad\\zum\\Ordner"]

def install(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
    except subprocess.CalledProcessError:
        print(f"Fehler beim Installieren von {package}.")

def ensure_basic_packages():
    basic_packages = ['setuptools', 'wheel', 'pip']
    for package in basic_packages:
        install(package)

def get_required_libraries():
    return [
        'winshell', 
        'pywin32', 
        'python-docx',
        'colorama',
        'PyPDF2',
    ]

def install_missing_libraries(required):
    for package in required:
        try:
            if package == 'win32com':
                install('pywin32')
            elif package == 'python-docx':
                install('python-docx')
            elif package == 'docx':
                install('python-docx')
            else:
                install(package)
            print(f"{package} wurde erfolgreich installiert.")
        except Exception as e:
            print(f"Fehler beim Installieren von {package}: {str(e)}")
    print("Installation abgeschlossen.")

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
            print(f"Es wurde noch kein {prompt} festgelegt.")
            while True:
                new_value = input(f"Bitte geben Sie den {prompt} ein: ")
                if is_valid_directory(new_value):
                    return new_value
                else:
                    print(f"Ungültiger Pfad. Bitte geben Sie einen gültigen absoluten Pfad ein.")
        else:
            print(f"Aktueller {prompt}: {current_value}")
            change = input(f"Möchten Sie den {prompt} ändern? (j/n): ").lower()
            if change == 'j':
                while True:
                    new_value = input(f"Neuer {prompt}: ")
                    if is_valid_directory(new_value):
                        return new_value
                    else:
                        print(f"Ungültiger Pfad. Bitte geben Sie einen gültigen absoluten Pfad ein.")
            else:
                return current_value

    config["base_folder_path"] = update_path("base_folder_path", "Pfad zum Basisordner")
    config["usb_path"] = update_path("usb_path", "Pfad zum USB-Laufwerk")

    use_notes = input("Möchten Sie einen Notizordner verwenden? (j/n): ").lower() in ['j', 'ja', 'y', 'yes']
    config["create_note"] = use_notes
    if use_notes:
        config["notes_folder_path"] = update_path("notes_folder_path", "Pfad zum Notizordner")
    else:
        config["notes_folder_path"] = ""

    config["svp_options"]["default_teacher"] = input("Standard-Lehrername: ") or config["svp_options"].get("default_teacher", "")

    include_abbreviations = input("Möchten Sie Abkürzungen einschließen? (j/n): ").lower() in ['j', 'ja', 'y', 'yes']
    config["svp_options"]["include_abbreviations"] = include_abbreviations

    change_time_slots = input("Möchten Sie die Zeitslots ändern? (j/n): ").lower() in ['j', 'ja', 'y', 'yes']
    if change_time_slots:
        for slot in config["svp_options"]["time_slots"]:
            print(f"Aktueller Zeitslot {slot}: {config['svp_options']['time_slots'][slot][0]} - {config['svp_options']['time_slots'][slot][1]}")
            start_time = input(f"Neue Startzeit für Slot {slot} (leer lassen für unverändert): ")
            end_time = input(f"Neue Endzeit für Slot {slot} (leer lassen für unverändert): ")
            if start_time:
                config["svp_options"]["time_slots"][slot][0] = start_time
            if end_time:
                config["svp_options"]["time_slots"][slot][1] = end_time

    with open(PROJECT_CONFIG_PATH, 'w') as f:
        json.dump(config, f, indent=4)

    print("Konfiguration wurde erfolgreich aktualisiert.")

def create_shortcut():
    try:
        import winshell
        from win32com.client import Dispatch
    except ImportError:
        print("Konnte winshell oder win32com nicht importieren. Verknüpfung wird nicht erstellt.")
        return

    main_path = os.path.join(current_dir, 'main.py')
    icon_path = os.path.join(current_dir, 'icon.ico')
    shortcut_path = os.path.join(current_dir, 'TeacherPy.lnk')

    if not os.path.exists(main_path):
        print("main.py nicht gefunden.")
        return
    if not os.path.exists(icon_path):
        print("icon.ico nicht gefunden.")
        return

    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.Targetpath = sys.executable
    shortcut.Arguments = f'"{main_path}"'
    shortcut.WorkingDirectory = current_dir
    shortcut.IconLocation = icon_path
    shortcut.save()

    print(f"Verknüpfung erstellt: {shortcut_path}")

def main():
    print("Überprüfe und installiere grundlegende Pakete...")
    ensure_basic_packages()

    print("Installiere erforderliche Bibliotheken...")
    required = get_required_libraries()
    install_missing_libraries(required)

    print("Aktualisiere die Konfiguration...")
    update_config()

    print("Erstelle die Verknüpfung...")
    create_shortcut()

    print("Setup abgeschlossen.")

if __name__ == "__main__":
    main()