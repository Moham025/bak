# 🚀 Backend Flask - Engineering Tools

**Backend Flask seulement - Prêt pour Railway deployment**

## 📋 Description
Backend Python Flask pour les outils d'ingénierie :
- **Bon à Envoyer** : Traitement de fichiers Excel
- **Combine Armatures** : Traitement de fichiers CSV

## 🛠 Technologies
- **Flask** : Serveur web
- **Pandas/Openpyxl** : Traitement de données
- **Gunicorn** : Serveur WSGI pour production

## 🚀 Déploiement Railway
Ce dossier est configuré pour Railway avec :
- `Procfile` : Configuration du serveur
- `requirements.txt` : Dépendances Python
- `.gitignore` : Fichiers exclus

## 📡 API Endpoints
- `POST /get-sheet-names` : Obtenir les feuilles Excel
- `POST /process-excel` : Traiter le fichier Excel  
- `POST /combine-armatures` : Combiner les CSV armatures

## 🔧 Variables d'environnement
- `PORT` : Port du serveur (défini automatiquement par Railway)
- `FLASK_ENV` : Environment (production par défaut)
