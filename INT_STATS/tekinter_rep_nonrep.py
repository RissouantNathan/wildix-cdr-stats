import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Fonction pour compter les appels par statut et par tag
def compter_appels_par_statut(chemin_fichier_excel):
    try:
        # Charger le fichier Excel
        df = pd.read_excel(chemin_fichier_excel)
        
        # Convertir la colonne "Durée de la sonner. (s.)" en numérique
        df['Durée de la sonner. (s.)'] = pd.to_numeric(df['Durée de la sonner. (s.)'], errors='coerce')

        # Filtrer les lignes où la valeur de la colonne "Durée de la sonner. (s.)" est supérieure à 10
        df = df[df['Durée de la sonner. (s.)'] > 10]

        # Filtrer les lignes où la valeur de la colonne "Tag" est égale à chaque tag
        tags = ['adv', 'compta', 'depl', 'deb adv', 'deb support', 'support', 'svi']
        comptes = {}
        for tag in tags:
            df_tag = df[df['Tag'] == tag]
            comptes[tag] = {
                'Rép.': df_tag[df_tag['Statut'] == 'Rép.'].shape[0],
                'Non rép.': df_tag[df_tag['Statut'] == 'Non rép.'].shape[0]
            }

        return comptes
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

    comptes = compter_appels_par_statut(chemin_fichier_excel)

    # Supprimer les entrées précédentes du tableau
    for row in treeview.get_children():
        treeview.delete(row)

    # Ajouter de nouvelles entrées au tableau avec les tags
    for tag, compte in comptes.items():
        treeview.insert("", "end", values=(tag, compte['Rép.'], compte['Non rép.']))

# Création de la fenêtre principale
root = tk.Tk()
root.title("Comptage des appels par statut")

# Définir la taille de la fenêtre
root.geometry("600x400")  # Largeur x Hauteur

# Bouton pour choisir le fichier Excel
btn_choisir_fichier = tk.Button(root, text="Choisir un fichier Excel", command=choisir_fichier)
btn_choisir_fichier.pack(pady=10)

# Tableau pour afficher les résultats
treeview = ttk.Treeview(root, columns=("Tag", "Rep", "Non Rep"), show="headings")
treeview.heading("Tag", text="Tag")
treeview.heading("Rep", text="Appels Rép.")
treeview.heading("Non Rep", text="Appels Non Rép.")
treeview.pack(fill="both", expand=True)

# Lancer l'application
root.mainloop()
