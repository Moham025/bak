# Architecture du Projet CivilEngTools Pro

Ce document décrit l'architecture du projet CivilEngTools Pro, en détaillant la structure des fichiers et la manière dont les composants frontend (Flutter/Dart) et backend (Python) interagissent.

## Arborescence des Fichiers Principaux

```
.
├── backend/                   # Code source du serveur backend Python
│   ├── bon_a_envoye.py        # Logique Python pour l'outil "Bon à Envoyer"
│   ├── covnumletter.py        # Module Python pour la conversion de nombres en lettres
│   ├── __pycache__/           # Cache Python (généré automatiquement)
│   ├── Aramture/              # Potentiellement pour les fichiers liés aux armatures (à confirmer)
│   └── recup bon file envoi.txt # Fichier texte (rôle à déterminer)
├── build/                     # Contient les artefacts de compilation (généré par Flutter)
├── lib/                       # Code source principal de l'application Flutter (Dart)
│   ├── main.dart              # Point d'entrée de l'application Flutter
│   ├── pages/                 # Widgets représentant les différentes pages/écrans
│   │   └── home_page.dart     # Widget pour la page d'accueil
│   └── widgets/               # Widgets réutilisables pour l'interface utilisateur
│       ├── app_footer.dart    # Widget pour le pied de page
│       ├── app_header.dart    # Widget pour l'en-tête
│       ├── excel_upload_modal.dart # Widget pour la modale de téléversement Excel
│       ├── tool_card.dart     # Widget pour afficher une carte d'outil
│       └── tools_grid.dart    # Widget pour afficher la grille des outils
├── test/                      # Fichiers de tests pour Flutter
├── web/                       # Fichiers spécifiques à la plateforme web pour Flutter
│   └── index.html             # Point d'entrée HTML pour l'application web
├── .gitignore                 # Fichiers et dossiers ignorés par Git
├── analysis_options.yaml      # Configuration pour l'analyse statique Dart
├── pubspec.lock               # Versions exactes des dépendances Flutter/Dart
├── pubspec.yaml               # Dépendances et métadonnées du projet Flutter
├── README.md                  # Informations générales sur le projet (à compléter)
└── ... (autres fichiers de configuration et générés par l'IDE)
```

## Fonctionnement Général et Interaction Dart/Python

Le projet est une application web construite avec **Flutter** pour la partie frontend (interface utilisateur) et **Python (Flask)** pour la partie backend (logique métier et calculs).

### 1. Frontend (Flutter/Dart - dans le dossier `lib/`)

*   **Point d'entrée (`lib/main.dart`)**: Initialise l'application Flutter, configure le thème global (couleurs, polices) et lance la `HomePage`.
*   **Structure des Pages (`lib/pages/`)**:
    *   `home_page.dart`: Définit la structure de la page d'accueil, qui est généralement composée d'un en-tête, d'une grille d'outils et d'un pied de page.
*   **Widgets Réutilisables (`lib/widgets/`)**:
    *   `app_header.dart` & `app_footer.dart`: Composants pour l'en-tête et le pied de page.
    *   `tools_grid.dart`: Affiche une liste d'outils disponibles. Chaque outil est défini avec une icône, un titre, une description et une action à exécuter lorsqu'on clique dessus.
    *   `tool_card.dart`: Représente visuellement chaque outil dans la grille.
    *   `excel_upload_modal.dart`: Fournit une interface utilisateur (fenêtre modale) pour permettre aux utilisateurs de sélectionner et de téléverser des fichiers (spécifiquement des fichiers Excel pour l'outil "Bon à Envoyer"). Cette modale est susceptible d'être réutilisée ou adaptée pour d'autres outils nécessitant une interaction avec des fichiers.
*   **Interaction Utilisateur**: L'utilisateur navigue sur l'interface web construite avec Flutter. Lorsqu'il clique sur un outil (par exemple, "Bon à Envoyer" ou le futur "Combine arm. Poutre"), une action est déclenchée.

### 2. Backend (Python/Flask - dans le dossier `backend/`)

*   **Serveur Flask**: Un serveur web léger (probablement défini dans un fichier comme `bon_a_envoye.py` ou un fichier principal du backend) écoute les requêtes HTTP provenant du frontend Flutter.
*   **Points d'API (Endpoints)**:
    *   Le backend expose des URL spécifiques (endpoints) que le frontend peut appeler. Par exemple, pour l'outil "Bon à Envoyer", il y a des endpoints comme `/get-sheet-names` et `/process-excel` dans `bon_a_envoye.py`.
    *   Chaque endpoint est associé à une fonction Python qui exécute une logique métier.
*   **Logique Métier**:
    *   `bon_a_envoye.py`: Contient la logique pour traiter les fichiers Excel téléversés : lecture des feuilles, manipulation des données, application de règles spécifiques (comme la conversion de montants en lettres), et la génération d'un nouveau fichier Excel.
    *   `covnumletter.py`: Un module utilitaire fournissant la fonction `conv_number_letter` pour la conversion de nombres en toutes lettres, utilisé par `bon_a_envoye.py`.
    *   *Futur outil "Combine arm. Poutre"*: Il nécessitera la création de nouvelles fonctions Python et potentiellement de nouveaux endpoints pour gérer la sélection de multiples fichiers (ou la réception de leurs données) et leur combinaison.

### 3. Communication Frontend <-> Backend

*   **Requêtes HTTP**: Lorsque l'utilisateur interagit avec un outil dans l'interface Flutter (par exemple, en téléversant un fichier via `excel_upload_modal.dart`), le code Dart effectue une requête HTTP (généralement `POST` ou `GET`) vers l'URL correspondante du backend Python.
    *   Pour le téléversement de fichiers, Flutter envoie les données du fichier (octets) dans le corps de la requête.
    *   Pour d'autres actions, il peut envoyer des paramètres dans l'URL ou sous forme de JSON.
*   **Traitement Backend**: Le serveur Flask reçoit la requête, identifie l'endpoint appelé, et exécute la fonction Python associée. Cette fonction peut lire les données envoyées (fichiers, paramètres), effectuer des calculs, manipuler des fichiers sur le serveur, etc.
*   **Réponses HTTP**: Une fois le traitement terminé, le backend Python renvoie une réponse HTTP au frontend Flutter.
    *   Cela peut être un fichier (comme le fichier Excel traité pour "Bon à Envoyer").
    *   Cela peut être des données JSON (comme la liste des noms de feuilles d'un fichier Excel).
    *   Ou simplement un statut de succès ou d'erreur.
*   **Mise à jour de l'Interface Utilisateur**: Le code Dart dans Flutter reçoit la réponse du backend. En fonction de cette réponse, il met à jour l'interface utilisateur : affiche un message de succès/erreur, télécharge un fichier, affiche des nouvelles données, etc.

### Flux Typique pour un Outil (Ex: "Bon à Envoyer")

1.  **Flutter**: L'utilisateur clique sur l'outil "Bon à Envoyer" dans `ToolsGrid`.
2.  **Flutter**: L'action `onTapAction` associée (définie dans `tools_grid.dart`) appelle `showExcelUploadModal(context)`.
3.  **Flutter**: Le widget `excel_upload_modal.dart` s'affiche, permettant à l'utilisateur de sélectionner un fichier Excel.
4.  **Flutter**: Après sélection, l'utilisateur clique sur un bouton pour traiter le fichier. Le code Dart de la modale envoie le fichier Excel et les noms des feuilles sélectionnées à l'endpoint `/process-excel` du backend Python via une requête HTTP `POST`.
5.  **Python**: Le serveur Flask (dans `bon_a_envoye.py`) reçoit la requête sur `/process-excel`. La fonction `process_excel_file_route` appelle `traiter_fichier_excel_core`.
6.  **Python**: `traiter_fichier_excel_core` effectue toutes les manipulations nécessaires sur le fichier Excel (conversion de colonnes, nettoyage du total en lettres en utilisant `cl_conv_number_letter` de `covnumletter.py`, suppression de colonnes, etc.).
7.  **Python**: Si le traitement réussit, un nouveau fichier Excel est généré en mémoire. Le backend renvoie ce fichier dans la réponse HTTP.
8.  **Flutter**: Le code Dart reçoit la réponse contenant le fichier Excel. Il déclenche alors le téléchargement du fichier dans le navigateur de l'utilisateur.

Ce modèle d'interaction sera similaire pour le nouvel outil "Combine arm. Poutre", où Flutter collectera l'information nécessaire (probablement une liste de fichiers Excel), l'enverra à un nouvel endpoint Python, et le backend Python effectuera la logique de combinaison et renverra le fichier résultant.

Ce document devrait servir de guide pour comprendre l'organisation actuelle et faciliter l'intégration de nouvelles fonctionnalités. 