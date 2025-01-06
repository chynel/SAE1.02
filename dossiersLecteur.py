# date : 6 janvier 2025
# auteur : votre nom
# version 1

# Ce programme liste les fichiers contenus dans le dossier data 
# et affiche le contenu des fichiers

import os
# Le module os permet d'interagir avec le système d'exploitation, 
# il permet ainsi de gérer l’arborescence des fichiers, 
# de fournir des informations sur le système d'exploitation processus, 
# variables systèmes, ainsi que de nombreuses fonctionnalités du systèmes.


#La variable directory est déclarée par le nom du dossier : data
directory = 'data'
for filename in os.listdir(directory):
    # filename est le nom d'un fichier dans le dossier directory
    # f continent le chemin relatif vers un fichier du dossier directory
    f = os.path.join(directory, filename)
    # verification que le fichier adressé par f est bien un fichier
    if os.path.isfile(f):
        print("-------------------- Nom du fichier : ", filename, "--------------------")
        # ouverture de fichier 
        fichier = open(f, "r", encoding = 'utf-8', errors='ignore')
        #lecture de fichier
        lignes = fichier.readlines()
        print(lignes)
        # fermeture de fichier 
        fichier.close()

