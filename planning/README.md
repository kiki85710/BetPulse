# Planning des épreuves BetPulse

| Fichier | Rôle |
|---|---|
| `epreuves.csv` | Données de référence : une ligne par épreuve, 25 compétitions, saison 2026-2027 et suivantes. |
| `planning_betpulse.xlsx` | Classeur généré à partir du CSV (onglets Lisez-moi, Épreuves, Référentiels), à importer dans Google Sheets. |
| `build_planning.py` | Régénère le classeur depuis `epreuves.csv`. |
| `consolider.py` | Fusionne des CSV de recherche, contrôle les valeurs et attribue les identifiants. |
| `lire_planning.py` | Lit l'onglet Épreuves du Google Sheet et renvoie un DataFrame pandas. |

## Import dans Google Sheets

1. Google Drive > Nouveau > Importer un fichier > `planning_betpulse.xlsx`.
2. Ouvrir le fichier avec Google Sheets, puis Fichier > Enregistrer au format Google Sheets.
3. Partager > Tous les utilisateurs disposant du lien : Lecteur (nécessaire pour le mode CSV de `lire_planning.py`).
4. Copier l'identifiant du classeur (entre `/d/` et `/edit` dans l'URL).

## Lecture en Python

```bash
pip install pandas            # + gspread pour le mode api
export BETPULSE_SHEET_ID=<identifiant>
python planning/lire_planning.py --jours 14
python planning/lire_planning.py --competition "Formule 1" --jours 60
```

## Mise à jour

Le Google Sheet devient la référence une fois importé. `epreuves.csv` et le `.xlsx` restent l'état vérifié au 16/10/2026.
