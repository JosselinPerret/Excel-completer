# Excel Completer

Une application simple pour mettre à jour un fichier Excel avec des données de couverture de test extraites d'un rapport texte.

![GIF de présentation](presentation.gif)

## Version en ligne

Vous pouvez accéder à l'application Streamlit en ligne ici : [https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/](https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/)

## Description

Cette application permet de :
- Importer un fichier Excel/CSV contenant des informations de composants
- Importer un fichier texte de rapport de couverture de test
- Mettre à jour le fichier Excel avec les données de couverture extraites du rapport
- Télécharger le fichier Excel mis à jour

## Fonctionnalités

- Interface utilisateur intuitive avec Streamlit
- Extraction automatique des données de couverture à partir du rapport texte
- Mise à jour du fichier Excel avec les pourcentages de couverture
- Aperçu des données mises à jour avant téléchargement
- Support pour les fichiers Excel (.xlsx) et CSV (.csv)

## Comment utiliser l'application

### Utilisation en ligne
1. Accédez à l'application sur [https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/](https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/)
2. Téléchargez votre fichier Excel/CSV
3. Téléchargez votre rapport de couverture (fichier texte)
4. Cliquez sur "Traiter les fichiers"
5. Téléchargez le fichier mis à jour

### Exécution locale

#### Prérequis
- Python 3.6 ou supérieur
- pip (gestionnaire de paquets Python)

#### Installation
1. Clonez ce dépôt :
   ```bash
   git clone https://github.com/JosselinPerret/Excel-completer.git
   cd Excel-completer
   ```

2. Installez les dépendances requises :
   ```bash
   pip install -r requirements.txt
   ```

#### Exécution
Pour lancer l'application Streamlit :
```bash
streamlit run streamlit_app.py
```

Pour lancer la version Tkinter (ancienne version) :
```bash
python coverage_excel.py
```

## Structure du projet

- `streamlit_app.py` : Application principale utilisant Streamlit
- `coverage_excel.py` : Ancienne version utilisant Tkinter
- `requirements.txt` : Fichier listant les dépendances Python
- `Plan_de_Test_par_Composant - CEBB3_SA_SB.xlsx` : Exemple de fichier Excel
- `Plan_de_Test_par_Composant - CEBB3_SA_SB.csv` : Version CSV du fichier Excel
- `ANALYZEREPORT CEBB3_SA_SB.txt` : Exemple de rapport de couverture

## Fonctionnement technique

1. L'application extrait les données de couverture du rapport texte en utilisant des expressions régulières
2. Elle recherche les composants dans le fichier Excel et met à jour la colonne "COVERAGE %"
3. Pour les composants non trouvés dans le rapport, une valeur de "0%" est attribuée
4. Le fichier Excel mis à jour est généré et prêt à être téléchargé

## Format du rapport de couverture

L'application s'attend à trouver des informations de couverture dans le format suivant :
```
Test Summary for U10 (MC14519)
...
Totals: 100.00%
```

Où "U10" est l'identifiant du composant et "100.00%" est le pourcentage de couverture.

## Format du fichier Excel/CSV

Le fichier Excel/CSV doit contenir une colonne nommée "COMP." qui liste les identifiants des composants.
Une colonne "COVERAGE %" sera ajoutée ou mise à jour avec les valeurs extraites du rapport.
