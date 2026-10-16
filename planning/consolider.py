"""Fusionne les CSV de recherche, contrôle les valeurs et attribue les identifiants.

Usage : python planning/consolider.py DOSSIER_CSV
Écrit planning/epreuves.csv (source de vérité versionnée dans le dépôt).
"""
import csv
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

from build_planning import CODE, GENRES, SPORTS, STATUTS

CHAMPS = ["id", "sport", "competition", "saison", "phase", "manche", "intitule", "date_debut", "date_fin",
          "heure_debut", "lieu", "pays", "genre", "statut", "ouvert_pronos", "source_url", "notes"]


def main(dossier):
    lignes, erreurs = [], []
    for f in sorted(Path(dossier).glob("*.csv")):
        with f.open(encoding="utf-8-sig") as fh:
            for n, r in enumerate(csv.DictReader(fh), start=2):
                r = {k.strip(): (v or "").strip() for k, v in r.items() if k}
                ou = f"{f.name}:{n}"
                if r["competition"] not in CODE:
                    erreurs.append(f"{ou} compétition inconnue : {r['competition']}")
                if r["sport"] not in SPORTS:
                    erreurs.append(f"{ou} sport inconnu : {r['sport']}")
                if r["statut"] not in STATUTS:
                    erreurs.append(f"{ou} statut inconnu : {r['statut']}")
                if r["genre"] not in GENRES:
                    erreurs.append(f"{ou} genre inconnu : {r['genre']}")
                for k in ("date_debut", "date_fin"):
                    if r[k]:
                        try:
                            date.fromisoformat(r[k])
                        except ValueError:
                            erreurs.append(f"{ou} {k} invalide : {r[k]}")
                if r["date_debut"] and not r["date_fin"]:
                    r["date_fin"] = r["date_debut"]
                if r["date_debut"] and r["date_fin"] and r["date_fin"] < r["date_debut"]:
                    erreurs.append(f"{ou} date_fin avant date_debut")
                lignes.append(r)
    if erreurs:
        print("\n".join(erreurs))
        sys.exit(1)

    lignes.sort(key=lambda r: (list(CODE).index(r["competition"]), r["saison"], r["date_debut"] or "9999", r["genre"], r["intitule"]))
    compteur = defaultdict(int)
    for r in lignes:
        saison = r["saison"].replace("20", "", 1).replace("-20", "") if "-" in r["saison"] else r["saison"]
        cle = (CODE[r["competition"]], saison)
        compteur[cle] += 1
        r["id"] = f"{cle[0]}-{saison}-{compteur[cle]:03d}"
        r["ouvert_pronos"] = "Oui"
    sortie = Path(__file__).parent / "epreuves.csv"
    with sortie.open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=CHAMPS, extrasaction="ignore")
        w.writeheader()
        w.writerows(lignes)
    print(f"{len(lignes)} lignes -> {sortie}")


if __name__ == "__main__":
    main(sys.argv[1])
