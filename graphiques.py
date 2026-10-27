"""
Graphiques interactifs (HTML) des dépenses dans le temps.

    py graphiques.py --ouvrir
    py graphiques.py --periode mois --compte Joint --ouvrir
    py graphiques.py --help
"""
import _lanceur

_lanceur.dans_le_venv()

from comptes_bancaires.graphiques import main  # noqa: E402

main()
