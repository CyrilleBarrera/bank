"""
Lanceur commun aux commandes  py comptes.py / py graphiques.py / py actualiser.py.
Il relance automatiquement le script avec le Python de l'environnement virtuel (.venv) :
pas besoin d'activer l'environnement.
"""
import os
import subprocess
import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent
VENV = RACINE / ".venv"
PYTHON_VENV = VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def dans_le_venv() -> None:
    """Si on n'est pas déjà dans .venv, relance le script appelant avec le Python de .venv."""
    # le dossier donnees/ du projet est utilisé quel que soit le dossier courant
    os.environ.setdefault("COMPTES_RACINE", str(RACINE))
    if Path(sys.prefix).resolve() == VENV.resolve():
        return
    if not PYTHON_VENV.exists():
        sys.exit("Environnement .venv introuvable : lancez d'abord  py install.py")
    script = Path(sys.argv[0]).resolve()
    sys.exit(subprocess.call([str(PYTHON_VENV), str(script), *sys.argv[1:]]))
