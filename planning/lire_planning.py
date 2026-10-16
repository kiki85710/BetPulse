"""Lit le planning BetPulse depuis Google Sheets et renvoie un DataFrame pandas.

Deux modes :
  1. Export CSV (sans clé) : le classeur doit être partagé « Tous les utilisateurs disposant du lien : lecteur ».
  2. API Google Sheets (gspread) : un compte de service a accès au classeur ; GOOGLE_APPLICATION_CREDENTIALS
     pointe vers son fichier JSON.

Variables d'environnement :
  BETPULSE_SHEET_ID  identifiant du classeur (facultatif, SHEET_ID par défaut)

Exemples :
  python planning/lire_planning.py                 # épreuves des 14 prochains jours
  python planning/lire_planning.py --jours 30 --competition "Formule 1"
"""
import argparse
import os
from datetime import date, timedelta
from urllib.parse import quote

import pandas as pd

SHEET_ID = "1qBgAZn_C7Q5KkNmbjBqAiGqa_e1FXKfNeCGssGBhnBY"  # classeur « Planning des épreuves BetPulse »
ONGLET = "Épreuves"
COLONNES_DATE = ["date_debut", "date_fin", "date_verification"]


def lire_csv(sheet_id: str) -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(ONGLET)}"
    return pd.read_csv(url, dtype=str)


def lire_api(sheet_id: str) -> pd.DataFrame:
    import gspread  # pip install gspread

    gc = gspread.service_account(filename=os.environ["GOOGLE_APPLICATION_CREDENTIALS"])
    valeurs = gc.open_by_key(sheet_id).worksheet(ONGLET).get_all_records(value_render_option="FORMATTED_VALUE")
    return pd.DataFrame(valeurs, dtype=str)


def charger(sheet_id: str | None = None, mode: str = "csv") -> pd.DataFrame:
    sheet_id = sheet_id or os.environ.get("BETPULSE_SHEET_ID", SHEET_ID)
    df = lire_api(sheet_id) if mode == "api" else lire_csv(sheet_id)
    df = df.dropna(subset=["id"])
    for col in COLONNES_DATE:
        df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce").dt.date
    df["manche"] = pd.to_numeric(df["manche"], errors="coerce").astype("Int64")
    return df


def a_venir(df: pd.DataFrame, jours: int = 14, competition: str | None = None) -> pd.DataFrame:
    """Épreuves ouvertes aux pronostics qui commencent dans les `jours` prochains jours."""
    debut, fin = date.today(), date.today() + timedelta(days=jours)
    masque = (df["ouvert_pronos"] == "Oui") & (df["date_debut"] >= debut) & (df["date_debut"] <= fin)
    masque &= ~df["statut"].isin(["Annulé", "Reporté"])
    if competition:
        masque &= df["competition"] == competition
    return df[masque].sort_values(["date_debut", "competition"])


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--mode", choices=["csv", "api"], default="csv")
    p.add_argument("--jours", type=int, default=14)
    p.add_argument("--competition")
    a = p.parse_args()
    sel = a_venir(charger(mode=a.mode), a.jours, a.competition)
    print(sel[["id", "competition", "intitule", "date_debut", "date_fin", "heure_debut", "statut"]].to_string(index=False))
