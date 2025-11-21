# Revue du code

## Observations clés
- **Cohérence du cookie de session corrigée** : `get_token_from_request` récupère désormais le jeton via le nom de cookie configurable `SESSION_COOKIE_NAME`, ce qui évite les échecs d'authentification quand le nom est personnalisé via les variables d'environnement.
- **Validation des mots de passe administrateurs perfectible** : la création d'utilisateurs admin ne vérifie actuellement pas la longueur minimale (seule la présence du mot de passe est contrôlée). Les changements de mot de passe côté compte imposent `PASSWORD_MIN_LENGTH`, mais la création via `/admin/utilisateurs` peut accepter des mots de passe trop courts.

## Recommandations
- **Centraliser le nom du cookie** : conserver l'utilisation systématique de `SESSION_COOKIE_NAME` pour lire, écrire et supprimer le cookie de session afin d'éviter les divergences de configuration.
- **Aligner les contrôles de robustesse des mots de passe** : ajouter une vérification de la longueur minimale et des règles communes lors de la création d'utilisateurs dans `/admin/utilisateurs` afin d'assurer un niveau de sécurité cohérent pour tous les comptes, en réutilisant idéalement la même logique que la mise à jour de mot de passe du compte.
