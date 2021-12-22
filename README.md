# ELE216 Scripts de correction

Quelques scripts que j'ai écris pour faciliter certaines tâches redondantes dans mon travail de correction en ELE216 et autres charges de lab que j'ai données à l'ETS. Je les rends disponible ici au cas où ça pourrait être utile à d'autres.

## Outils nécessaires

Node.js 16.13.1
NPM 8.3.0

Ces scripts pourraient fonctionner avec d'autres versions de Node mais je ne peux le garantir.

Une fois que Node et NPM sont installés, exécuter `npm install` dans le dossier des scripts (là où se trouve `package.json`) pour installer les dépendances.

## Récupérer une liste d'étudiants

Certains des scripts suivants nécessitent une liste d'étudiants `csv` générée par Moodle. Pour récupérer cette liste:

1. Visiter le site Moodle du cours;
1. Visiter Notes -> Exporter -> Fichier Texte;
1. Laisser les options par défaut;
1. Cliquer "Exporter".

## Exécuter les scripts

Tous les scripts peuvent être exécutés ainsi: `node <script> <paramètres>`.

## Scripts

### Création des grilles d'évaluation (individuelles)
Ce script reçoit un fichier Excel (`xlsx`) contenant au minimum une _worksheet_ qui est une grille d'évaluation vierge, et une liste d'étudiants au format CSV générée par Moodle (voir point précédent). Ce script va générer, dans le fichier Excel, une copie de la grille d'évaluation vierge au nom de chaque étudiant(e).

### Création des grilles d'évaluation (équipes)
Ce script reçoit un fichier Excel (`xlsx`) contenant au minimum une _worksheet_ qui est une grille d'évaluation vierge. Ce script va générer, dans le fichier Excel, une copie de la grille d'évaluation vierge pour chaque équipe.

|Paramètre|Description|Obligatoire|Par défaut|
|-|-|-|-|
|`-f|--fichier-evaluations <file>`|Fichier Excel dans lequel seront créées les grilles d'évaluation.|Oui||
|`-n|--nombre-equipes <nbr>`|Nombre d'équipes. Des grilles d'évaluation seront créées pour les équipes 1 à n.|Oui||
|`-e|--exclure-equipes <nbrs>`|Numéros d'équipes inutilisés. Écrire comme une string, séparer les numéros par une virgule, ex. "1, 5, 7".|||
|`-t|--template-name <name>`|Nom de la worksheet qui sera utilisée comme modèle pour les grilles d'évaluation.||\_\_TEMPLATE\_\_|
|`-g|--numero-groupe <groupe>`|Numéro du groupe.|||
|`-c|--cellule-groupe <cellule>`|Identifiant de la cellule dans lequel indiquer le numéro du groupe. ex. "A3"|||

### Extraction des grilles d'évaluation (individuelles et équipes)
Ce script parcours un fichier Excel contenant les grilles d'évaluation d'une classe (voir scripts précédents), et pour chaque étudiant/équipe, crée un nouveau fichier Excel contenant uniquement la grille d'évaluation de cet étudiant/cette équipe.

### Extraction des notes (individuelles)
Ce script reçoit un fichier Excel contenant les grilles d'évaluation des étudiants (voir script précédent), et une liste de classe CSV générée par Moodle. Ce script parcours la liste d'étudiants, puis les grilles d'évaluation, et recopie la note de chaque étudiant dans la liste de classe.

Une fois que les notes sont extraites dans la liste d'étudiants, cette liste peut être ré-importée dans Moodle afin d'entrer toutes les notes d'un coup. Pour faire ça:

1. Visiter le site Moodle du cours;
1. Visiter Notes -> Importation -> Fichier CSV;
1. Choisir le fichier `csv`;
1. Laisser les autres options par défaut;
1. Cliquer "Déposer les notes";
1. Suivre les autres étapes pour terminer l'importation.

### Impression des notes (équipes)
Malheureusement, je n'ai pas trouvé de moyen facile d'extraire les notes dans la liste de classe pour les travaux d'équipe. À la place, j'ai fait ce script qui va imprimer les notes de toutes les équipe dans la console. Il faut quand même entrer les notes manuellement dans Moodle, mais c'est plus rapide que de devoir retourner voir la grille Excel de chaque équipe.
