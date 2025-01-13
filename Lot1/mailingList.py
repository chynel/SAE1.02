# date : 7 janvier 2025
# auteurs : Ferrand SOKI, Salma ALAHYAN, Yasmine TARAZA
# version 8

# Ce programme liste les fichiers contenus dans le dossier data 
# et affiche le contenu des fichiers

import os
# Le module os permet d'interagir avec le système d'exploitation, 
# il permet ainsi de gérer l’arborescence des fichiers, 
# de fournir des informations sur le système d'exploitation processus, 
# variables systèmes, ainsi que de nombreuses fonctionnalités du systèmes.

import csv # Le module csv permet de lire et écrire dans des fichiers CSV (Comma-Separated Values).


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

# Définir le dossier contenant les fichiers texte
directory = 'data' #chemin relatif vers le repertoire de données
output_csv = 'mailingList.csv' #Nom du fichier de sorti

# Préparer les en-têtes du fichier CSV
entete = ['NOM', 'PRENOM', 'MAIL_ETUDIANTS']

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
                nom = prenom = email = None

                # Parcourir les lignes pour trouver les champs nécessaires
                for ligne in lignes:
                    if ligne.startswith("NOM"):
                        nom = extraire_valeur_apres_mot_cle(ligne, "NOM")
                    elif ligne.startswith("PRENOM"):
                        prenom = extraire_valeur_apres_mot_cle(ligne, "PRENOM")
                    elif ligne.startswith("EMAIL"):
                        email = extraire_valeur_apres_mot_cle(ligne, "EMAIL")

                # Si toutes les informations sont trouvées, les écrire dans le fichier CSV
                if nom and prenom and email:
                    ecrire.writerow([nom, prenom, email])
                

print(f"Extraction terminée. Les données ont été écrites dans {output_csv}.")