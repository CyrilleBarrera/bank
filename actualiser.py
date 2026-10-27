"""
Tout en une commande :  py actualiser.py
  1. ajoute tous les CSV de donnees/exports au classeur donnees/Comptes.xlsx (sans doublons) ;
  2. génère les graphiques et les ouvre dans le navigateur.

Options :
  --nouveau       recrée le classeur à partir de tous les CSV (applique les nouvelles règles
                  aux anciennes opérations, mais efface les corrections faites dans Excel)
  --sans-ouvrir   ne pas ouvrir le navigateur
Fermez Comptes.xlsx dans Excel avant de lancer.
"""
import _lanceur

_lanceur.dans_le_venv()

import sys  # noqa: E402

from comptes_bancaires import graphiques, releve  # noqa: E402

args = sys.argv[1:]
exports = _lanceur.RACINE / "donnees" / "exports"
if not list(exports.glob("*.csv")):
    sys.exit(f"Aucun fichier CSV dans {exports} : déposez-y vos exports BoursoBank.")

sys.argv = ["comptes", str(exports / "*.csv")] + (["--nouveau"] if "--nouveau" in args else [])
releve.main()
sys.argv = ["comptes-graphiques"] + ([] if "--sans-ouvrir" in args else ["--ouvrir"])
graphiques.main()
