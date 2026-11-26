# Consignes pour Claude — DASHBOARD OPUS

À lire avant toute modification de ce dépôt, quel que soit le projet Claude d'où vient la demande.

1. **Toujours partir de la dernière version de la branche `main` de ce dépôt** (`git fetch` puis `git log`). Ne jamais republier un `index.html` construit à partir d'une copie plus ancienne (copie iCloud, fichier d'une conversation précédente, fichier téléchargé plus tôt) : cela efface les évolutions faites entre-temps. *Incident du 26/11/2026 : une publication faite depuis le projet SOPROPECHE, construite sur une version antérieure, a effacé les évolutions publiées une heure plus tôt (phases, ligne cliquable, couleurs) ; il a fallu les réappliquer.*
2. **Juste avant de publier, revérifier que `main` n'a pas bougé** depuis le début du travail. S'il a bougé : fusionner (fusion à 3 voies avec la version de départ comme base), retester, puis publier.
3. Le brief complet (`BRIEF_DASHBOARD_OPUS.md`) est dans le dossier iCloud « CLAUDE/GESTION DE PROJET » de Nicolas et dans le Projet claude.ai « GESTION DE PROJET » : le lire, et compléter son journal des évolutions après chaque publication.
4. Décrire chaque publication dans le message de commit.
