import os
import sys
import subprocess
import tkinter as tk
from tkinter import messagebox


# Fonction pour s'assurer que pandas est installé
def installer_pandas():
    try:
        import pandas
    except ImportError:
        # Si pandas n'est pas installé, installez-le avec pip
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas"])
# Appelez la fonction pour s'assurer que pandas est installé
installer_pandas()

# Fonction pour exécuter un script choisi
def executer_script(script_name):
    try:
        # Obtenir le répertoire actuel
        repertoire_actuel = os.path.dirname(os.path.abspath(__file__))
        
        # Créer le chemin complet vers le script
        chemin_script = os.path.join(repertoire_actuel, script_name)
        
        subprocess.run(["python", chemin_script], check=True)
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erreur", f"Le script a rencontré une erreur : {e}")

# Fonction pour choisir un script à exécuter
def choisir_script():
    # Dictionnaire des scripts disponibles avec des descriptions lisibles
    scripts = {
        "tekinter_NonRepSonnerie.py": "Non Rép. par Durée de la Sonnerie",
        "tekinter_rep_nonrep.py": "Répondus vs Non Répondus",
        "tekinter_stats_appel.py": "Temps Moyens des appels",
        "tekinter_appelParSTATS.py": "Nombre appels par tag"
    }

    # Créez une nouvelle fenêtre pour la sélection de script
    selection_window = tk.Toplevel(root)
    selection_window.title("Sélectionner un script")
    selection_window.geometry("300x200")

    # Créez un message de sélection
    label = tk.Label(selection_window, text="Choisissez un script à exécuter :")
    label.pack(pady=10)

    # Ajoutez des boutons pour chaque script avec des descriptions lisibles
    for script, description in scripts.items():
        btn_script = tk.Button(selection_window, text=description, command=lambda s=script: executer_script(s))
        btn_script.pack(pady=5)


# Création de la fenêtre principale
root = tk.Tk()
root.title("Exécution de scripts Tkinter")
root.geometry("350x150")

# Bouton pour ouvrir la sélection de script
btn_choisir_script = tk.Button(root, text="Choisir un script", command=choisir_script)
btn_choisir_script.pack(pady=20)

# Lancer l'application
root.mainloop()
