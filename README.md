# PANT Togo P179138 — Tableau de bord V2

Version 2 du classeur de l’équipe intégrée **sans modification des données**. Projet distinct de SIRA Guinée.

## Contenu
- `index.html`, `styles.css`, `app.js`, `data.js` : tableau de bord en français avec drapeau togolais horizontal.
- `PACT_100_JOURS_MASTER.xlsx` : copie **identique** du fichier V2 transmis, renommée pour la synchronisation.
- `sync_data.py` : synchronisation Excel → `data.js`.
- `.github/workflows/update-dashboard.yml` : mise à jour GitHub.
- `RAPPORT_CONTROLE.md` : informations constatées sans correction.

## Publication
Créer un dépôt GitHub indépendant. Importer tous les fichiers en conservant `.github`. Activer GitHub Pages (branche `main`, répertoire `/`). À chaque mise à jour, remplacer `PACT_100_JOURS_MASTER.xlsx` et exécuter le workflow.
