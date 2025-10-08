# CRM Local

Ce projet est une application FastAPI simple permettant de gérer des clients locaux et d'importer des fichiers Excel.

## Installation des dépendances

Assurez-vous d'utiliser un environnement virtuel Python, puis installez les dépendances avec :

```bash
pip install -r requirements.txt
```

## Lancer le serveur de développement

Initialisez la base de données et démarrez le serveur avec uvicorn :

```bash
uvicorn app:app --reload
```

L'application sera disponible sur http://127.0.0.1:8000.
