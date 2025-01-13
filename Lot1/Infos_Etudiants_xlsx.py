# date : 7 janvier 2025
# auteurs : Ferrand SOKI, Salma ALAHYAN, Yasmina TARAZA
# version 6

# Ce programme liste les fichiers contenus dans le dossier data 
# et affiche le contenu des fichiers

import os
# Le module os permet d'interagir avec le système d'exploitation, 
# il permet ainsi de gérer l’arborescence des fichiers, 
# de fournir des informations sur le système d'exploitation processus, 
# variables systèmes, ainsi que de nombreuses fonctionnalités du systèmes.

from openpyxl import Workbook  # Permet de manipuler les fichiers Excel
from collections import Counter  # Fournit un moyen de compter les occurrences d'éléments dans un iterable




# Fonction pour extraire la partie après un mot-clé donné
#Inspirer de l'autre projet en programmation
def extraire_valeur_apres_mot_cle(ligne, mot_cle):
    # Trouvercle mot-clé se termine ici par exemple NOM
    debut = ligne.find(mot_cle) + len(mot_cle)
    # Ignorer les espaces ou tabulations après le mot-clé
    while debut < len(ligne) and ligne[debut] in " \t":
        debut += 1
    # Extraire la partie restante
    valeur = ""
    while debut < len(ligne) and ligne[debut] not in "\n":
        valeur += ligne[debut]
        debut += 1
    return valeur


# Définition du dossier contenant les fichiers texte
directory = 'data'  # Chemin relatif vers le répertoire de données
output_xlsx = 'Infos_Etudiants1.xlsx'  # Nom du fichier de sortie

# Créeation d'un nouveau classeur Excel
workbook = Workbook()

# Ajout de la feuille des étudiants
feuille_etudiants = workbook.active
feuille_etudiants.title = "Etudiants"

# Préparation des en-têtes pour la feuille Excel des étudiants
entete = ['NOM', 'PRENOM', 'EMAIL'] # Définition des colonnes de la feuille Excel
# Écriture des en-têtes dans la première ligne
feuille_etudiants.append(entete)

# Initialisation d'un dictionnaire pour les erreurs
erreurs = Counter()

# Parcourt des fichiers dans le dossier data
for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    if os.path.isfile(f):  # Vérification de la veracité du fichier
        with open(f, 'r', encoding='utf-8', errors='ignore') as fichier:
            lignes = fichier.readlines()

            # Variables pour stocker les informations
            nom = prenom = email = None

            # Parcourt des lignes pour extraire les informations
            for ligne in lignes:
                # Extraction des informations de l'étudiant
                if ligne.startswith("NOM"):
                    nom = extraire_valeur_apres_mot_cle(ligne, "NOM")
                elif ligne.startswith("PRENOM"):
                    prenom = extraire_valeur_apres_mot_cle(ligne, "PRENOM")
                elif ligne.startswith("EMAIL"):
                    email = extraire_valeur_apres_mot_cle(ligne, "EMAIL")

                # Extraction des erreurs
                if "SyntaxError" in ligne or "TypeError" in ligne or "NameError" in ligne or "IndentationError" in ligne:
                    erreurs[ligne] += 1

            # Si toutes les informations sont trouvées, les écrire dans la feuille Excel des étudiants
            if nom and prenom and email:
                feuille_etudiants.append([nom, prenom, email])


# Ajoute d'une nouvelle feuille pour les erreurs
sheet_erreurs = workbook.create_sheet(title="Erreurs")

# En-têtes pour la feuille des erreurs
entete_erreurs = ['Erreur', 'Nombre de fois']
sheet_erreurs.append(entete_erreurs)

# Ajout des erreurs dans la feuille des erreurs
for erreur, count in erreurs.items():
    sheet_erreurs.append([erreur, count])


# Sauvegarder le fichier Excel
workbook.save(output_xlsx)

print(f"Extraction terminée. Les données ont été écrites dans {output_xlsx}.")
