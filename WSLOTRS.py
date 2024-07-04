import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
# un script pour les gouverner tous !

def analyser_appels_non_repondus(df):
    # Convertir les colonnes pertinentes en numérique
    df['Durée de la sonner. (s.)'] = pd.to_numeric(df['Durée de la sonner. (s.)'], errors='coerce')
    df['Temps de conv. (s.)'] = pd.to_numeric(df['Temps de conv. (s.)'], errors='coerce')

    # Analyse des appels non répondus par durée de sonnerie
    df_non_repondus = df[(df['Durée de la sonner. (s.)'] > 10) & (df['Statut'] == 'Non rép.')]
    stats_non_repondus = df_non_repondus.groupby('Durée de la sonner. (s.)').size()

    # Analyse des temps de conversation par tag
    tags = ['adv', 'compta', 'depl', 'deb adv', 'deb support', 'support', 'svi', 'svi deb support']
    moyennes_conv = {}
    moyennes_sonnerie = {}
    for tag in tags:
        df_tag = df[df['Tag'] == tag]
        moyenne_seconde_conv = df_tag['Temps de conv. (s.)'].mean()
        moyenne_seconde_sonnerie = df_tag['Durée de la sonner. (s.)'].mean()
        minutes_conv = int(moyenne_seconde_conv // 60)
        secondes_conv = int(moyenne_seconde_conv % 60)
        moyennes_conv[tag] = (minutes_conv, secondes_conv)
        moyennes_sonnerie[tag] = moyenne_seconde_sonnerie
    
    return stats_non_repondus, moyennes_conv, moyennes_sonnerie

def afficher_graphique_moyennes_temps_conv(moyennes):
    # Créer un graphique en barres pour les moyennes des temps de conversation
    tags = list(moyennes.keys())
    moyennes_minutes_secondes = [f'{m[0]} minutes {m[1]} secondes' for m in moyennes.values()]
    moyennes_minutes = [m[0] + m[1] / 60 for m in moyennes.values()]

    plt.figure(figsize=(12, 8))
    bars = plt.bar(tags, moyennes_minutes, color='skyblue')

    plt.xlabel('Tags')
    plt.ylabel('Durée Moyenne des Conversations (minutes)')
    plt.title('Durée Moyenne des Conversations par Tag')
    plt.xticks(rotation=45)
    
    # Ajouter les valeurs au-dessus des barres
    for bar, moyenne in zip(bars, moyennes_minutes_secondes):
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval, f'{moyenne}', ha='center', va='bottom')

    plt.tight_layout()
    plt.show()

def tracer_graphique(stats_non_repondus):
    # Tracer le graphique en barres pour les appels non répondus
    plt.figure(figsize=(10, 6))
    stats_non_repondus.plot(kind='bar', color='skyblue')
    plt.xlabel('Durée de la sonnerie (s.)')
    plt.ylabel('Nombre d\'appels non répondus')
    plt.title('Nombre d\'appels non répondus par durée de sonnerie (> 10s)')
    plt.xticks(rotation=45)
    plt.grid(axis='y')
    plt.tight_layout()
    plt.show()

def afficher_graphique_moyennes_sonnerie(moyennes_sonnerie):
    # Créer un graphique en barres pour les moyennes des durées de sonnerie par tag
    tags = list(moyennes_sonnerie.keys())
    moyennes_sonnerie_vals = [m for m in moyennes_sonnerie.values()]

    plt.figure(figsize=(12, 8))
    bars = plt.bar(tags, moyennes_sonnerie_vals, color='lightcoral')

    plt.xlabel('Tags')
    plt.ylabel('Durée Moyenne de la Sonnerie (secondes)')
    plt.title('Durée Moyenne de la Sonnerie par Tag')
    plt.xticks(rotation=45)
    
    # Ajouter les valeurs au-dessus des barres
    for bar, moyenne in zip(bars, moyennes_sonnerie_vals):
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width() / 2, yval, f'{moyenne:.2f} s', ha='center', va='bottom')

    plt.tight_layout()
    plt.show()

# Fonction pour compter les appels par TAG
def compter_appels_par_tag(df):
    try:
        # Créer un dictionnaire pour stocker le nombre total d'appels par TAG
        comptes = {}

        # Itérer sur chaque TAG unique pour compter les appels
        for tag in df['Tag'].unique():
            comptes[tag] = df[df['Tag'] == tag].shape[0]

        return comptes
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {e}")

# Fonction pour choisir un fichier et afficher les résultats dans un tableau
def choisir_fichier():
    # Choisir le fichier Excel
    chemin_fichier_excel = filedialog.askopenfilename(
        title="Choisir le fichier Excel",
        filetypes=[("Fichiers Excel", "*.xlsx"), ("Tous les fichiers", "*.*")]
    )

    if not chemin_fichier_excel:
        return

    try:
        # Charger le fichier Excel
        df = pd.read_excel(chemin_fichier_excel)

        # Compter les appels par TAG
        comptes = compter_appels_par_tag(df)

        # Supprimer les entrées précédentes du tableau
        for row in treeview.get_children():
            treeview.delete(row)

        # Ajouter les nouvelles entrées au tableau
        for tag, compte in comptes.items():
            treeview.insert("", "end", values=(tag, compte))

        # Analyser et afficher les graphiques
        stats_non_repondus, moyennes_conv, moyennes_sonnerie = analyser_appels_non_repondus(df)
        tracer_graphique(stats_non_repondus)
        afficher_graphique_moyennes_temps_conv(moyennes_conv)
        afficher_graphique_moyennes_sonnerie(moyennes_sonnerie)
    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite lors de l'analyse des données : {e}")

# Création de la fenêtre principale Tkinter
root = tk.Tk()
root.title("Analyse des appels")

# Définir la taille de la fenêtre
root.geometry("600x400")

# Bouton pour choisir le fichier Excel
btn_choisir_fichier = tk.Button(root, text="Choisir un fichier Excel", command=choisir_fichier)
btn_choisir_fichier.pack(pady=10)

# Tableau pour afficher les résultats
treeview = ttk.Treeview(root, columns=("Tag", "Nombre d'Appels"), show="headings")
treeview.heading("Tag", text="Tag")
treeview.heading("Nombre d'Appels", text="Nombre d'Appels")
treeview.pack(fill="both", expand=True)

# Lancer l'application Tkinter
root.mainloop()
