# Comptes bancaires

Suivi des comptes de Cyrille, Marion et du compte Joint à partir des exports CSV de BoursoBank :

- un classement automatique des opérations (type, libellé, moyen de paiement), réglable par des règles regex ;
- des **contextes** (vacances, événements) pour comparer le coût par jour, par exemple la nourriture en vacances et au quotidien ;
- un classeur Excel avec des formules (soldes, total général, synthèse, évolution mensuelle) et la configuration modifiable dans des onglets ;
- des graphiques interactifs Plotly (HTML) des dépenses dans le temps.

## Installation sous Windows

1. **Installer Python 3.10 ou plus récent** depuis <https://www.python.org/downloads/windows/>.
   Pendant l'installation, cochez **« Add python.exe to PATH »**.
2. **Placer `install.py`** dans le dossier du projet (par exemple `bank`).
3. **Lancer `install.py`** : double-clic, ou dans un terminal ouvert dans le dossier : `py install.py`
   (macOS / Linux : `python3 install.py`). Le script :
   - crée les dossiers et les fichiers du projet ;
   - crée un environnement Python privé (`.venv`) ;
   - installe openpyxl, pandas, plotly et pytest ;
   - crée `donnees\config_comptes.json` ;
   - lance les tests.

### Mettre à jour

Remplacez `install.py` par la nouvelle version et relancez `py install.py`.

- Les fichiers que vous n'avez pas modifiés sont mis à jour automatiquement.
- Ceux que vous avez modifiés sont conservés. `py install.py --mettre-a-jour` les remplace aussi.
- Vos relevés et votre classeur (`donnees\`) ne sont jamais modifiés.
- Votre configuration (`donnees\config_comptes.json`) reçoit les nouvelles règles, mais vos propres règles, types, comptes et contextes sont conservés. Une copie de l'ancienne version est gardée.

## Utilisation courante

1. Téléchargez l'export CSV depuis BoursoBank et déposez-le **tel quel** dans `donnees\exports\`.
   Ne l'ouvrez pas et ne l'enregistrez pas avec Excel avant : cela peut changer son format.
2. **Fermez `Comptes.xlsx`** s'il est ouvert dans Excel.
3. Dans un terminal ouvert dans le dossier du projet, lancez **`py actualiser.py`**.
   - Tous les CSV de `donnees\exports\` sont ajoutés à `donnees\Comptes.xlsx`, sans doublons. Les corrections que vous avez faites dans Excel sont conservées.
   - Les graphiques s'ouvrent ensuite dans le navigateur (`donnees\graphiques_comptes.html`).

Vous pouvez garder tous vos exports dans le dossier : les opérations déjà importées sont ignorées.

Après une mise à jour des règles, `py actualiser.py --nouveau` recrée le classeur à partir de tous les CSV. Les anciennes opérations sont alors reclassées avec les nouvelles règles, mais les corrections faites à la main dans Excel sont perdues.

### Commandes

Lancez-les dans un terminal (PowerShell ou invite de commandes) ouvert dans le dossier du projet. Inutile d'activer l'environnement `.venv` : les scripts s'en chargent. Sous macOS / Linux, remplacez `py` par `python3`.

| Commande | Effet |
|---|---|
| `py actualiser.py` | Importe tous les CSV de `donnees\exports\`, puis ouvre les graphiques |
| `py actualiser.py --nouveau` | Idem, mais recrée le classeur (reclasse tout avec les règles actuelles) |
| `py comptes.py donnees\exports\*.csv` | Importe les CSV dans `donnees\Comptes.xlsx` |
| `py comptes.py export.csv --compte Marion` | Force le compte, pour un numéro de compte inconnu de la config |
| `py comptes.py donnees\exports\*.csv --analyse Nourriture` | Affiche le coût par contexte |
| `py comptes.py --importer-config donnees\Comptes.xlsx` | Reporte les onglets de configuration d'Excel dans le JSON |
| `py graphiques.py --periode mois --ouvrir` | Graphiques regroupés par mois (ou `semaine`) |
| `py graphiques.py --compte Joint --hors-ligne` | Un seul compte, dans un fichier HTML utilisable sans Internet |
| `py comptes.py --help` | Toutes les options |

## Configuration

Les règles de libellés de la version 0.5 viennent du rapprochement entre un an d'exports BoursoBank et le classeur tenu à la main (oct. 2025 – sept. 2026). Pour chaque commerçant, c'est le nom et le type les plus récents saisis à la main qui ont été retenus. Les virements dont l'objet change à chaque fois (Luis Barrera, Lydia, chèques...) gardent un libellé générique et le type « À classer », à compléter dans Excel.

Tout se règle dans `donnees\config_comptes.json` :

- les numéros de compte et leur propriétaire ;
- les types ;
- les règles de libellés ;
- les moyens de paiement ;
- la correspondance avec les catégories de la banque ;
- les contextes.

Les mêmes réglages sont repris dans les onglets gris du classeur. Vous pouvez les modifier dans Excel, puis lancer `py comptes.py --importer-config donnees\Comptes.xlsx`. L'ancienne configuration est gardée en `.bak`.

L'onglet **« Sans règle »** liste les libellés qu'aucune règle ne reconnaît encore, avec une regex proposée.

## Développement avec VS Code

1. Ouvrez le dossier dans VS Code (**Fichier > Ouvrir le dossier**). Acceptez les extensions recommandées : Python et Claude Code.
2. **Tâches** : `Ctrl+Shift+P`, puis « Tasks: Run Task ». Vous pouvez y lancer l'actualisation, les graphiques, l'analyse par contexte ou les tests. `Ctrl+Shift+B` lance directement l'actualisation.
3. **Débogage** : `F5`, avec les configurations « import des CSV » et « graphiques ».
4. **Tests** : l'onglet Tests (icône en forme de fiole) lance `pytest`. Les tests utilisent uniquement les données fictives de `tests\donnees\`.

### Claude dans VS Code

Il faut VS Code 1.94 ou plus récent et un abonnement Claude payant (Pro, Max, Team ou Enterprise).

1. Installez l'extension **Claude Code** (éditeur : Anthropic) depuis l'onglet Extensions (`Ctrl+Shift+X`).
2. Ouvrez un fichier, cliquez sur l'icône ✱ en haut à droite et connectez-vous avec votre compte Claude.
3. Le fichier `CLAUDE.md` décrit le projet et ses règles, notamment ne jamais toucher à `donnees\`. Claude le lit automatiquement.

Exemples de demandes :

- « Ajoute une règle pour classer SPLASHHH en Loisirs »
- « Ajoute un graphique du cumul des dépenses par mois »
- « @releve.py explique-moi comment est calculé le solde initial »

## Structure

```
bank\
├─ install.py                   installation / mise à jour
├─ actualiser.py                CSV -> Excel -> graphiques, en une commande
├─ comptes.py, graphiques.py    commandes (utilisent .venv automatiquement)
├─ _lanceur.py                  relance les commandes avec le Python de .venv
├─ CLAUDE.md                    consignes pour Claude (VS Code)
├─ pyproject.toml               description du paquet et de ses dépendances
├─ src\comptes_bancaires\
│   ├─ releve.py                lecture CSV, règles, contextes, Excel
│   ├─ graphiques.py            graphiques Plotly
│   ├─ migration.py             mise à jour de votre configuration entre versions
│   └─ config_defaut.json       configuration modèle
├─ tests\                       tests pytest et CSV fictif
└─ donnees\                     VOS données (exclues de Git)
```

## Confidentialité

Le dossier `donnees\`, les fichiers `*.csv`, `*.xlsx` et `*.html` sont exclus de Git (`.gitignore`). Si vous publiez le code, par exemple sur GitHub, vos relevés ne partent pas avec.
