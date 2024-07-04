import subprocess
import sys

def install_package(package_name):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

# Liste des packages nécessaires
required_packages = [
    'pandas',
    'matplotlib',
    'openpyxl'
]

# Installation des packages
for package in required_packages:
    try:
        __import__(package)
        print(f"{package} est déjà installé.")
    except ImportError:
        print(f"{package} n'est pas installé. Installation...")
        install_package(package)
        print(f"{package} a été installé.")

# Vérification de tkinter
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    print("tkinter est déjà installé.")
except ImportError:
    print("tkinter n'est pas installé. Veuillez installer tkinter manuellement.")
    print("Sur Windows: tkinter est inclus avec Python. Vous n'avez rien à faire.")
    print("Sur Linux: Vous pouvez l'installer via votre gestionnaire de paquets.")
    print("Exemple: sudo apt-get install python3-tk")

print("Toutes les dépendances nécessaires ont été installées.")
