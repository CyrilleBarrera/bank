# Projet comptes-bancaires

Outil personnel de suivi des comptes de Cyrille, Marion et du compte Joint.
Il lit les exports CSV de BoursoBank, classe les opérations (type, libellé, moyen de
paiement, contexte), écrit un classeur Excel avec des formules et produit des graphiques Plotly.

## Structure
- `src/comptes_bancaires/releve.py` : cœur du projet.
  - Classes `Config`, `Contexte`, `Operation`.
  - Lecture du CSV : `lire_csv`. Calcul des soldes : `calculer_totaux`. Analyse : `analyser_contextes`.
  - Écriture Excel : `ecrire_excel`, avec les onglets Opérations, Paramètres, Synthèse, Évolution et les onglets de configuration.
  - Import de la config depuis Excel : `importer_config_excel`. Commande `comptes` : `main`.
- `src/comptes_bancaires/graphiques.py` : graphiques Plotly dans un seul fichier HTML (commande `comptes-graphiques`).
- `src/comptes_bancaires/config_defaut.json` : configuration modèle, copiée dans `donnees/` à l'installation.
- `src/comptes_bancaires/migration.py` : met à jour la config de l'utilisateur vers la version du modèle.
- `comptes.py`, `graphiques.py`, `actualiser.py`, `_lanceur.py` (racine) : commandes de l'utilisateur.
- `donnees/` : données RÉELLES de l'utilisateur (config, exports CSV, Comptes.xlsx). Exclu de Git.
- `tests/` : tests pytest sur des données FICTIVES (`tests/donnees/export_fictif.csv`).

## Commandes
- Installer ou mettre à jour : `py install.py` (Windows) ou `python3 install.py` ; options : `--mettre-a-jour`, `--sans-tests`.
- Tests : `.venv\Scripts\python -m pytest -q` (Windows) ou `.venv/bin/python -m pytest -q`.
- Commandes (relancent automatiquement le Python de .venv) : `py actualiser.py`, `py comptes.py ...`,
  `py graphiques.py ...`. L'utilisateur préfère les commandes Python aux fichiers .bat : ne pas en créer.

## Règles pour Claude
- L'utilisateur est sous Windows : chemins avec `pathlib`, tester les motifs `*.csv` (développés par le script).
- Noms de fonctions, messages et commentaires en français, comme le code existant.
- Ne jamais modifier, committer ni afficher en entier le contenu de `donnees/` (données bancaires
  personnelles). Pour un test, créer des données fictives dans `tests/donnees/`.
- Excel : les totaux et la synthèse sont des FORMULES (SUMIFS, INDEX/MATCH...), pas des valeurs
  calculées en Python. Ne pas utiliser de fonctions Excel 365 seulement (XLOOKUP, FILTER, UNIQUE...).
- Toute nouvelle règle de classement va dans la config (JSON ou onglets Excel), pas en dur dans le code.
- Modèle de config : `src/comptes_bancaires/config_defaut.json`. Si sa structure ou ses règles changent,
  incrémenter `version_config` et ajouter l'étape correspondante dans `migration.py` (la config de
  l'utilisateur dans `donnees/` n'est jamais écrasée, elle est migrée).
- Après chaque modification : lancer `pytest` et vérifier que tous les tests passent.
- Le format CSV BoursoBank a deux colonnes « Solde » : la 1re est le montant, la 2e le solde du compte.
