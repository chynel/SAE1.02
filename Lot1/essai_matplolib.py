# date : 7 janvier 2025
# auteurs : Ferrand SOKI, Salma ALAHYAN, Yasmina TARAZA
# version 6

import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt

# Définir le dossier contenant les fichiers texte
directory = 'data'  # chemin relatif vers le répertoire de données
output_xlsx = 'Infos.xlsx'  # Nom du fichier de sortie

# Créer un nouveau classeur Excel
workbook = Workbook()
sheet = workbook.active
sheet.title = "Etudiants"

# Préparer les en-têtes pour la feuille Excel
entete = ['NUMERO_ETUDIANT', 'NOM', 'PRENOM', 'EMAIL']
sheet.append(entete)

# Variables pour collecter les erreurs
erreurs = {}

# Parcourir les fichiers dans le dossier
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    if os.path.isfile(f):  # Vérifier si c'est bien un fichier
        with open(f, 'r', encoding='utf-8', errors='ignore') as fichier:
            lignes = fichier.readlines()

            # Variables pour stocker les informations
            numero_etudiant = nom = prenom = email = None

            # Parcourir les lignes pour trouver les champs nécessaires
            for ligne in lignes:
                if ligne.startswith("NUMERO_ETUDIANT"):
                    numero_etudiant = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("NOM"):
                    nom = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("PRENOM"):
                    prenom = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("EMAIL"):
                    email = ligne.split(maxsplit=1)[1].strip()
                elif '-' in ligne:  # Ligne contenant une erreur
                    # Extraire le type d'erreur
                    parts = ligne.split(' - ')
                    if len(parts) >= 3:
                        erreur_type = parts[2].strip()
                        erreurs[erreur_type] = erreurs.get(erreur_type, 0) + 1

            # Si toutes les informations sont trouvées, les écrire dans la feuille Excel
            if numero_etudiant and nom and prenom and email:
                sheet.append([numero_etudiant, nom, prenom, email])

# Générer un histogramme des erreurs
plt.figure(figsize=(10, 6))
plt.bar(erreurs.keys(), erreurs.values(), color='skyblue')
plt.title("Distribution des erreurs", fontsize=16)
plt.xlabel("Types d'erreurs", fontsize=12)
plt.ylabel("Nombre de fois", fontsize=12)
plt.xticks(rotation=45, ha='right', fontsize=10)
plt.tight_layout()

# Sauvegarder le graphique comme image
graph_image = 'erreurs_histogramme.png'
plt.savefig(graph_image)
plt.close()

# Ajouter le graphique à une nouvelle feuille dans le classeur Excel
sheet_visualisation = workbook.create_sheet(title="Visualisation")
img = Image(graph_image)
sheet_visualisation.add_image(img, "A1")

# Sauvegarder le fichier Excel
workbook.save(output_xlsx)

print(f"Extraction terminée. Les données et la visualisation ont été écrites dans {output_xlsx}.")
