# date : 7 janvier 2025
# auteurs : Ferrand SOKI, Salma ALAHYAN, Yasmina TARAZA
# version 8

# Ce programme liste les fichiers contenus dans le dossier data 
# et affiche le contenu des fichiers

import os
# Le module os permet d'interagir avec le système d'exploitation, 
# il permet ainsi de gérer l’arborescence des fichiers, 
# de fournir des informations sur le système d'exploitation processus, 
# variables systèmes, ainsi que de nombreuses fonctionnalités du systèmes.

import csv


# Définir le dossier contenant les fichiers texte
directory = 'data' #chemin relatif vers le repertoire de données
output_csv = 'mailingList .csv' #Nom du fichier de sorti

# Préparer les en-têtes du fichier CSV
entete = ['ETUDIANTS EMAIL']

# Ouvrir le fichier CSV en mode écriture
with open(output_csv, mode='w', encoding='utf-8', newline='') as fichier:
    ecrire = csv.writer(fichier)
    # Écrire les en-têtes
    ecrire.writerow(entete)

    # Parcourir les fichiers dans le dossier
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        if os.path.isfile(f):  # Vérifier si c'est bien un fichier
            with open(f, 'r') as fichier:
                lignes = fichier.readlines()
                
                # Variables pour stocker les informations
                numero_etudiant = nom = prenom = email = None

                # Parcourir les lignes pour trouver les champs nécessaires
                for ligne in lignes:
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
                    if ligne.startswith("EMAIL"):
                        email = ligne.split(maxsplit=1)[1].strip()

                # Si toutes les informations sont trouvées, les écrire dans le fichier CSV
                if email:
                    ecrire.writerow([email])
                

print(f"Extraction terminée. Les données ont été écrites dans {output_csv}.")