import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Fonction pour calculer la moyenne des temps de conversation pour chaque tag
def calculer_moyenne_temps_conv(chemin_fichier_excel):
    try:
        # Charger le fichier Excel
        df = pd.read_excel(chemin_fichier_excel)
        
        # Filtrer les lignes où la valeur de la colonne "Tag" est égale à chaque tag
        tags = ['adv', 'compta', 'depl', 'deb adv', 'deb support', 'support', 'svi', 'svi deb support']
        moyennes = {}
        for tag in tags:
            df_tag = df[df['Tag'] == tag]
            moyenne_seconde = df_tag['Temps de conv. (s.)'].mean()
            # Convertir en minutes et secondes
            minutes = int(moyenne_seconde // 60)
            secondes = int(moyenne_seconde % 60)
            moyennes[tag] = (minutes, secondes)
        
        return moyennes
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

# Fonction pour choisir un fichier et afficher les moyennes dans un tableau
def choisir_fichier():
    chemin_fichier_excel = filedialog.askopenfilename(
        title="Choisir le fichier Excel",
        filetypes=(("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*"))
    )

    if not chemin_fichier_excel:
        return

    moyennes = calculer_moyenne_temps_conv(chemin_fichier_excel)

    # Supprimer les entrées précédentes du tableau
    for row in treeview.get_children():
        treeview.delete(row)

    # Ajouter de nouvelles entrées au tableau
    for tag, moyenne in moyennes.items():
        treeview.insert("", "end", values=(tag, moyenne[0], moyenne[1]))

# Création de la fenêtre principale
root = tk.Tk()
root.title("Moyenne des temps de conversation par tag")

# Définir la taille de la fenêtre
root.geometry("600x400")  # Largeur x Hauteur

# Bouton pour choisir le fichier Excel
btn_choisir_fichier = tk.Button(root, text="Choisir un fichier Excel", command=choisir_fichier)
btn_choisir_fichier.pack(pady=10)

# Tableau pour afficher les résultats
treeview = ttk.Treeview(root, columns=("Tag", "Minutes", "Secondes"), show="headings")
treeview.heading("Tag", text="Tag")
treeview.heading("Minutes", text="Minutes")
treeview.heading("Secondes", text="Secondes")
treeview.pack(fill="both", expand=True)

# Lancer l'application
root.mainloop()
