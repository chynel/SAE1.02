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

from openpyxl import Workbook
from collections import Counter

# Définition du dossier contenant les fichiers texte
directory = 'data'  # Chemin relatif vers le répertoire de données
output_xlsx = 'Infos_Etudiants.xlsx'  # Nom du fichier de sortie

# Créeation d'un nouveau classeur Excel
workbook = Workbook()

# Ajout de la feuille des étudiants
feuille_etudiants = workbook.active
feuille_etudiants.title = "Etudiants"

# Préparation des en-têtes pour la feuille Excel des étudiants
entete = ['NUMERO_ETUDIANT', 'NOM', 'PRENOM', 'EMAIL']
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
            numero_etudiant = nom = prenom = email = None

            # Parcourt des lignes pour extraire les informations
            for ligne in lignes:
                # Extraction des informations de l'étudiant
                if ligne.startswith("NUMERO_ETUDIANT"):
                    """Le formatage dans le fichier .txt étant comme suit: 
                        NOM              CHARREAU
                        PRENOM           BRAS
                        EMAIL            bras.charreau@orangefr
                        On utilise maxsplit = 1 pour séparer dès le premier espace
                        rencontré.
                        Exemple: ligne = "NOM CHARREAU"
                                 result = ligne.split(maxsplit=1)
                                 print(result)  # ['NOM', 'CHARREAU']

                                 ligne = "NOM              CHARREAU MARIE"
                                 result = ligne.split(maxsplit=1)
                                 print(result)  # ['NOM', 'CHARREAU MARIE']
                    """
                    numero_etudiant = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("NOM"):
                    nom = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("PRENOM"):
                    prenom = ligne.split(maxsplit=1)[1].strip()
                elif ligne.startswith("EMAIL"):
                    email = ligne.split(maxsplit=1)[1].strip()

                # Extraction des erreurs
                if "SyntaxError" in ligne or "TypeError" in ligne or "NameError" in ligne or "IndentationError" in ligne:
                    erreurs[ligne.strip()] += 1

            # Si toutes les informations sont trouvées, les écrire dans la feuille Excel des étudiants
            if numero_etudiant and nom and prenom and email:
                feuille_etudiants.append([numero_etudiant, nom, prenom, email])


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
