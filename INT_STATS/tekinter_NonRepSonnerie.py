import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Fonction pour compter les appels non répondus par durée de sonnerie
def compter_appels_non_repondus_par_duree(chemin_fichier_excel):
    try:
        # Charger le fichier Excel
        df = pd.read_excel(chemin_fichier_excel)

        # Convertir la colonne "Durée de la sonner. (s.)" en numérique
        df['Durée de la sonner. (s.)'] = pd.to_numeric(df['Durée de la sonner. (s.)'], errors='coerce')

        # Filtrer les lignes où la valeur de la colonne "Durée de la sonner. (s.)" est supérieure à 10 et le statut est "Non rép."
        df_non_repondus = df[(df['Durée de la sonner. (s.)'] > 10) & (df['Statut'] == 'Non rép.')]

        # Compter le nombre d'appels non répondus pour chaque valeur de Durée de la sonnerie supérieure à 10
        stats_non_repondus = df_non_repondus.groupby('Durée de la sonner. (s.)').size()
        
        return stats_non_repondus
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

# Fonction pour choisir un fichier et afficher les résultats dans un tableau
def choisir_fichier():
    chemin_fichier_excel = filedialog.askopenfilename(
        title="Choisir le fichier Excel",
        filetypes=(("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*"))
    )

    if not chemin_fichier_excel:
        return

    stats_non_repondus_par_duree = compter_appels_non_repondus_par_duree(chemin_fichier_excel)

    # Supprimer les entrées précédentes du tableau
    for row in treeview.get_children():
        treeview.delete(row)

    # Ajouter de nouvelles entrées au tableau avec la durée de la sonnerie et le nombre d'appels non répondus
    for duree, count in stats_non_repondus_par_duree.items():
        treeview.insert("", "end", values=(duree, count))

# Création de la fenêtre principale
root = tk.Tk()
root.title("Appels non répondus par durée de la sonnerie")

# Définir la taille de la fenêtre
root.geometry("600x400")  # Largeur x Hauteur

# Bouton pour choisir le fichier Excel
btn_choisir_fichier = tk.Button(root, text="Choisir un fichier Excel", command=choisir_fichier)
btn_choisir_fichier.pack(pady=10)

# Tableau pour afficher les résultats
treeview = ttk.Treeview(root, columns=("Duree", "Appels Non Rép."), show="headings")
treeview.heading("Duree", text="Durée de la sonnerie (s.)")
treeview.heading("Appels Non Rép.", text="Nombre d'appels non répondus")
treeview.pack(fill="both", expand=True)

# Lancer l'application
root.mainloop()
