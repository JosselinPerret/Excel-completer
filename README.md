# 📊 Excel Completer

<div align="center">

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Une application élégante pour automatiser la mise à jour de vos fichiers Excel avec des données de couverture de test.**

![GIF de présentation](presentation.gif)

[🌐 Application en ligne](https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/) | [📖 Documentation](#-fonctionnalités) | [🚀 Installation](#-installation-locale)

</div>

## 🎯 Description

Cette application web simple mais puissante vous permet de :

- Importer un fichier Excel/CSV contenant des informations de composants
- Importer un fichier texte de rapport de couverture de test
- Mettre à jour automatiquement le fichier Excel avec les données de couverture
- Télécharger le fichier mis à jour tout en préservant le formatage original

## ✨ Fonctionnalités

| Fonctionnalité | Description |
|----------------|-------------|
| 🌐 **Interface moderne** | Interface utilisateur web intuitive propulsée par Streamlit |
| 🔍 **Extraction intelligente** | Analyse automatique des données de couverture par expressions régulières |
| 🔄 **Mise à jour précise** | Ajout ou mise à jour de la colonne "COVERAGE %" avec formatage approprié |
| 👁️ **Aperçu instantané** | Visualisation des données mises à jour avant téléchargement |
| 📁 **Format flexible** | Support pour les fichiers Excel (.xlsx) et CSV (.csv) |
| 🎨 **Préservation du formatage** | Conservation du style et du formatage des fichiers Excel existants |

## 🚀 Guide d'utilisation

### 🌐 Utilisation en ligne

Accédez directement à l'application déployée pour une utilisation immédiate sans installation :

1. Ouvrez l'application sur [Streamlit Cloud](https://josselinperret-excel-completer-streamlit-app-7ep3fi.streamlit.app/)
2. Importez votre fichier Excel/CSV via le sélecteur de fichier
3. Importez votre rapport de couverture de test (fichier texte)
4. Cliquez sur le bouton "Traiter les fichiers"
5. Prévisualisez et téléchargez le fichier mis à jour

### 💻 Installation locale

#### Prérequis

- Python 3.6 ou supérieur
- pip (gestionnaire de paquets Python)

#### Installation

1. Clonez le dépôt sur votre machine locale :

```bash
git clone https://github.com/JosselinPerret/Excel-completer.git
cd Excel-completer
```

2. Créez un environnement virtuel (recommandé) et installez les dépendances :

```bash
python -m venv .venv
source .venv/bin/activate  # Sur Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

#### Exécution

Lancez l'application avec l'une des commandes suivantes :

```bash
# Pour l'interface Streamlit (recommandé)
streamlit run streamlit_app.py
```

```bash
# Pour l'ancienne interface Tkinter
python coverage_excel.py
```

## 📂 Structure du projet

| Fichier | Description |
|---------|-------------|
| `streamlit_app.py` | Application principale avec interface web Streamlit |
| `coverage_excel.py` | Version alternative avec interface Tkinter |
| `requirements.txt` | Liste des dépendances Python requises |
| `Plan_de_Test_par_Composant - CEBB3_SA_SB.xlsx` | Exemple de fichier Excel d'entrée |
| `Plan_de_Test_par_Composant - CEBB3_SA_SB.csv` | Version CSV de l'exemple |
| `ANALYZEREPORT CEBB3_SA_SB.txt` | Exemple de rapport de couverture |
| `README.md` | Documentation du projet |

## ⚙️ Fonctionnement technique

1. **Extraction des données** : L'application analyse le rapport texte en utilisant des expressions régulières pour identifier les pourcentages de couverture par composant.
   
2. **Correspondance des données** : Elle associe les identifiants des composants du rapport avec ceux du fichier Excel.
   
3. **Mise à jour du fichier** : La colonne "COVERAGE %" est créée ou mise à jour avec les valeurs extraites, préservant le formatage original du fichier.
   
4. **Valeurs par défaut** : Pour les composants non trouvés dans le rapport, une valeur de "0%" est attribuée.
   
5. **Génération du fichier** : Le fichier Excel est mis à jour en mémoire puis proposé au téléchargement.

## 📄 Format des fichiers

### Format du rapport de couverture

L'application s'attend à trouver des informations de couverture dans le format suivant :

```text
Test Summary for U10 (MC14519)
...
Totals: 100.00%
```

Où "U10" est l'identifiant du composant et "100.00%" est le pourcentage de couverture.

### Format du fichier Excel/CSV

Le fichier Excel/CSV doit contenir une colonne nommée "COMP." qui liste les identifiants des composants.
Une colonne "COVERAGE %" sera ajoutée ou mise à jour avec les valeurs extraites du rapport.

## 📊 Exemple de résultat

Après traitement, le fichier Excel contiendra une colonne "COVERAGE %" formatée avec les pourcentages de couverture pour chaque composant identifié dans le rapport.

## 📝 Licence

Ce projet est distribué sous licence MIT. Vous êtes libre de l'utiliser, le modifier et le distribuer selon les termes de cette licence.
