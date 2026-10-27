"""
Importe des exports CSV dans le classeur (donnees/Comptes.xlsx).

    py comptes.py donnees/exports/*.csv
    py comptes.py donnees/exports/*.csv --nouveau          recrée le classeur
    py comptes.py donnees/exports/*.csv --analyse Nourriture
    py comptes.py --importer-config donnees/Comptes.xlsx   onglets Excel -> JSON
    py comptes.py --help                                   toutes les options
"""
import _lanceur

_lanceur.dans_le_venv()

from comptes_bancaires.releve import main  # noqa: E402

main()
