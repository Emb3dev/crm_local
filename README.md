# CRM Local

CRM Local est une application FastAPI pensée pour les équipes de service client qui gèrent des interventions locales. Elle rassemble la
base clients, les prestations sous-traitées, les consommables (filtres & courroies) ainsi qu'un plan de charge interactif afin de
faciliter le pilotage opérationnel au quotidien.

## Fonctionnalités principales

- ✔️ Authentification sécurisée (JWT + mots de passe hachés) et tableau de bord compte utilisateur.
- 👥 Gestion des clients, entreprises et contacts avec import Excel et fiches détaillées.
- 🛠️ Pilotage des prestations : création, édition, suivi du statut et import/export depuis Excel.
- 🧰 Module « Filtres & Courroies » avec suivi des références, quantités et imports dédiés.
- 📅 Plan de charge dynamique exposé via des endpoints API pour l'édition et l'import/export.
- 🛡️ Espace d'administration pour gérer les utilisateurs et le référentiel de prestations.

## Prérequis

- Python 3.11+ (recommandé)
- `pip` et un environnement virtuel (`python -m venv .venv`)
- SQLite (fourni nativement avec Python, aucun serveur externe requis)

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Configuration

Les variables d'environnement suivantes permettent d'adapter l'application (valeurs par défaut entre parenthèses) :

| Variable | Rôle |
| --- | --- |
| `CRM_SECRET_KEY` (`change-me`) | Clé utilisée pour signer les JWT. À remplacer en production. |
| `CRM_TOKEN_EXPIRE_MINUTES` (`480`) | Durée de validité (en minutes) des tokens d'accès. |
| `CRM_ADMIN_USERNAME` / `CRM_ADMIN_PASSWORD` (`admin` / `admin`) | Identifiants du compte super-administrateur créé au démarrage. |
| `CRM_SESSION_COOKIE_NAME` (`session_token`) | Nom du cookie qui stocke le token JWT. |
| `CRM_SESSION_COOKIE_SECURE` (`false`) | Forcer l'attribut `Secure` sur le cookie (utiliser `true` derrière HTTPS). |

> ℹ️ Les paramètres ci-dessus sont définis dans `app.py` et peuvent être fournis via un fichier `.env` ou votre orchestrateur (Docker,
> systemd, etc.).

## Initialisation de la base de données

La base SQLite est stockée dans `crm.db`. Lors du premier lancement (ou après avoir supprimé ce fichier), exécutez :

```bash
python - <<'PY'
from database import init_db
init_db()
PY
```

Cela crée les tables (clients, prestations, filtres, plan de charge, utilisateurs…) et alimente le référentiel des prestations par défaut
(`defaults.py`).

## Lancer le serveur de développement

```bash
uvicorn app:app --reload
```

L'application est alors disponible sur http://127.0.0.1:8000. Les assets statiques sont servis depuis `/static` et les templates sont
dans `/templates`.

## Connexion & comptes

1. Rendez-vous sur `/login`.
2. Identifiez-vous avec le compte administrateur (`CRM_ADMIN_USERNAME` / `CRM_ADMIN_PASSWORD`).
3. Créez ensuite des utilisateurs supplémentaires depuis `/admin/utilisateurs` et attribuez-leur des mots de passe générés ou
   personnalisés.

Chaque utilisateur peut modifier son mot de passe dans `/mon-compte` et la dernière activité est tracée automatiquement.

## Import / Export

Plusieurs modules acceptent des fichiers Excel. Vous pouvez télécharger des modèles à partir de l'interface ou directement via :

- `/clients/import/template`
- `/prestations/import/template`
- `/filtres-courroies/filtres/import/template`
- `/filtres-courroies/courroies/import/template`
- `/prestations/referentiel/export/` (export du référentiel des prestations)

Les fichiers importés sont traités dans `importers.py` et alimentent les modèles SQLModel (`models.py`).

## Structure du projet

```text
crm_local/
├── app.py                # Routes FastAPI, authentification, dépendances et logique applicative
├── models.py             # Modèles SQLModel pour les clients, prestations, filtres, plan de charge, etc.
├── crud.py               # Fonctions de persistance et d'interrogation de la base de données
├── database.py           # Initialisation SQLite + migrations légères
├── importers.py          # Parsing des fichiers Excel (openpyxl)
├── templates/            # Pages Jinja2 (base, login, listes clients, plan de charge…)
├── static/styles.css     # Feuille de style principale
└── requirements.txt      # Dépendances Python
```

## Développement & contributions

- Activez le rechargement automatique avec `uvicorn app:app --reload`.
- Ajoutez de nouvelles dépendances dans `requirements.txt` puis exécutez `pip install -r requirements.txt`.
- Utilisez `black`, `ruff` ou votre outil favori pour garder un code cohérent (non configurés par défaut mais recommandés).
- Les pull requests doivent inclure une description claire des changements et, si possible, des captures d'écran pour les modifications UI.

Bon développement ! 🚀
