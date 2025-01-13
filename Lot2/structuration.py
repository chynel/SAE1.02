from openpyxl import load_workbook, Workbook

# Chemins relatifs pour les fichiers
input_file = "Infos_Etudiants.xlsx"  # Le fichier source doit être dans le même répertoire que le script
output_file = "Erreurs_Separees.xlsx"  # Le fichier de sortie sera généré dans le même répertoire

# Charger le classeur Excel
workbook_source = load_workbook(input_file)
feuille_source = workbook_source.active

# Créer un nouveau classeur Excel pour le résultat
workbook_result = Workbook()
feuille_resultat = workbook_result.active
feuille_resultat.title = "Erreurs"

# En-têtes pour le fichier de sortie
entetes = ['DATE', 'NUMERO', 'TYPE', 'SPECIFICITE', 'REPONSE']
feuille_resultat.append(entetes)

# Fonction pour séparer les données d'une ligne
def extraire_champs(ligne):
    champs = [part.strip() for part in ligne.split(' - ')]  # Divise la ligne à chaque " - " et nettoie les espaces
    return champs

# Compteur pour suivre les lignes traitées et ignorées après constat de la réduction des occurences
lignes_traitees = 0
lignes_ignorees = 0

# Parcourir les lignes du fichier source
for row in feuille_source.iter_rows(min_row=2, values_only=True):  # Ignorer les en-têtes
    for partie in row:
        if partie and isinstance(partie, str) and "-" in partie:
            champs = extraire_champs(partie)
            if len(champs) >= 5:
                feuille_resultat.append(champs[:5])  # Ajouter les 5 champs extraits à la feuille
                lignes_traitees += 1
            else:
                lignes_ignorees += 1  # Lignes ignorées si elles ont moins de 5 champs
                print(f"Ligne ignorée (moins de 5 champs) : {partie}")  # Affiche la ligne ignorée
        else:
            lignes_ignorees += 1  # Lignes ignorées si elles ne contiennent pas de "-"
            print(f"Ligne ignorée (pas de '-') : {partie}")  # Affiche la ligne ignorée
    
# Sauvegarder le fichier Excel résultant après ajout des données
workbook_result.save(output_file)

# Afficher le nombre de lignes traitées et ignorées
print(f"Nombre de lignes traitées : {lignes_traitees}")
print(f"Nombre de lignes ignorées : {lignes_ignorees}")
print(f"Les données ont été extraites et sauvegardées dans le fichier {output_file}.")
