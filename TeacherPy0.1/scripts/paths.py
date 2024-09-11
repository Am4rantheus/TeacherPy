import os

# Finde das Basisverzeichnis basierend auf dem aktuellen Skriptpfad
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
TEACHERPY_DIR = os.path.dirname(CURRENT_DIR)  # Ein Verzeichnis höher

# Basis-Verzeichnis für die Skripte
SCRIPT_BASE_DIR = os.path.join(TEACHERPY_DIR, "scripts")

# Pfade zu den einzelnen Skripten
SVP_SCRIPT_PATH = os.path.join(SCRIPT_BASE_DIR, "svp.py")
FINAL_SCRIPT_PATH = os.path.join(SCRIPT_BASE_DIR, "Final.py")
NEU_SCRIPT_PATH = os.path.join(SCRIPT_BASE_DIR, "neu.py")
ARCHIVE_SCRIPT_PATH = os.path.join(SCRIPT_BASE_DIR, "archive.py")

# Pfade zu den Konfigurationsdateien
PROJECT_CONFIG_PATH = os.path.join(TEACHERPY_DIR, "project_config.json")
LAYOUT_CONFIG_PATH = os.path.join(TEACHERPY_DIR, "layout.json")

# Pfade zu den Vorlagen
TEMPLATE_PATH = os.path.join(TEACHERPY_DIR, "Vorlage", "SVP_Vorlage.dotx")
TEMPLATE_NOTE_PATH = os.path.join(TEACHERPY_DIR, "Vorlage", "template.note")