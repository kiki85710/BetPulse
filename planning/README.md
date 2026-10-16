# Planning des épreuves BetPulse

| Fichier | Rôle |
|---|---|
| `epreuves.csv` | Données de référence : une ligne par épreuve, 25 compétitions, saison 2026-2027 et suivantes. |
| `planning_betpulse.xlsx` | Classeur généré à partir du CSV (onglets Lisez-moi, Épreuves, Référentiels), à importer dans Google Sheets. |
| `build_planning.py` | Régénère le classeur depuis `epreuves.csv`. |
| `consolider.py` | Fusionne des CSV de recherche, contrôle les valeurs et attribue les identifiants. |
| `lire_planning.py` | Lit l'onglet Épreuves du Google Sheet et renvoie un DataFrame pandas. |

## Google Sheet de référence

[Planning des épreuves BetPulse](https://docs.google.com/spreadsheets/d/1qBgAZn_C7Q5KkNmbjBqAiGqa_e1FXKfNeCGssGBhnBY/edit) : c'est la version à jour et modifiable.
`epreuves.csv` et `planning_betpulse.xlsx` sont l'instantané vérifié du 16/10/2026.

Pour le mode CSV de `lire_planning.py` (sans clé API), partager le classeur en « Tous les utilisateurs disposant du lien : Lecteur ».

## Lecture en Python

```bash
pip install pandas            # + gspread pour le mode api
python planning/lire_planning.py --jours 14
python planning/lire_planning.py --competition "Formule 1" --jours 60
```

