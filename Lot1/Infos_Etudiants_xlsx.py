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

# Définir le dossier contenant les fichiers texte
directory = 'data' #chemin relatif vers le repertoire de données
output_xlsx = 'Infos_Etudiants.xlsx' #Nom du fichier de sorti

# Créer un nouveau classeur Excel
workbook = Workbook()
# Sélectionner la feuille active
sheet = workbook.active
# Renommer la feuille
sheet.title = "Etudiants"

# Préparer les en-têtes pour la feuille Excel
entete = ['NUMERO_ETUDIANT', 'NOM', 'PRENOM', 'EMAIL']
# Écrire les en-têtes dans la première ligne
sheet.append(entete)

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

            # Si toutes les informations sont trouvées, les écrire dans la feuille Excel
            if numero_etudiant and nom and prenom and email:
                sheet.append([numero_etudiant, nom, prenom, email])

# Sauvegarder le fichier Excel
workbook.save(output_xlsx)

print(f"Extraction terminée. Les données ont été écrites dans {output_xlsx}.")
