#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Installateur du projet « comptes-bancaires ».

Placez ce fichier à la racine du projet (ex. le dossier « bank ») puis lancez-le :
    Windows       :  double-clic sur install.py   ou   py install.py
    macOS / Linux :  python3 install.py

Il crée ou met à jour :
    - les dossiers et fichiers du projet (code, tests, réglages VS Code, commandes Python) ;
    - l'environnement virtuel .venv et y installe les dépendances (openpyxl, pandas, plotly, pytest) ;
    - donnees/exports/ et donnees/config_comptes.json (si absent) ;
    - la configuration existante, migrée vers la nouvelle version (vos réglages sont gardés) ;
puis lance les tests.

Mise à jour : les fichiers que vous n'avez pas modifiés sont remplacés automatiquement ;
ceux que vous avez modifiés sont conservés (un message les liste).

Options :
    --mettre-a-jour   remplace AUSSI les fichiers du projet que vous avez modifiés
    --sans-tests      ne lance pas les tests à la fin
    --sans-venv       crée seulement les fichiers (pas d'environnement virtuel)
    --git             initialise un dépôt Git (si Git est installé)
    --dossier CHEMIN  installe ailleurs que dans le dossier de ce fichier

Le dossier donnees/ (vos relevés, votre configuration, votre classeur) n'est JAMAIS modifié,
sauf pour y créer ce qui manque.
Les fichiers du projet sont intégrés plus bas, compressés (zlib + base64).
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import shutil
import subprocess
import sys
import venv
import zlib
from pathlib import Path

VERSION = "0.5.0"
VERSION_PYTHON_MIN = (3, 10)
MANIFESTE = ".installation.json"          # empreintes des fichiers écrits par l'installateur
OBSOLETES = ["actualiser.bat", "comptes.bat", "graphiques.bat"]   # remplacés par des commandes Python
WINDOWS = os.name == "nt"


def message(texte: str = "") -> None:
    try:
        print(texte, flush=True)
    except UnicodeEncodeError:                    # vieille console Windows
        print(texte.encode("ascii", "replace").decode(), flush=True)


def etape(n: int, total: int, texte: str) -> None:
    message(f"\n=== {n}/{total}  {texte} ===")


def python_du_venv(racine: Path) -> Path:
    return racine / ".venv" / ("Scripts/python.exe" if WINDOWS else "bin/python")


def executer(commande: list, racine: Path) -> bool:
    message("  > " + " ".join(str(c) for c in commande))
    return subprocess.run([str(c) for c in commande], cwd=racine).returncode == 0


def empreinte(contenu: bytes) -> str:
    return hashlib.sha256(contenu).hexdigest()


def empreintes_connues(racine: Path) -> dict[str, set[str]]:
    """Empreintes des fichiers tels qu'écrits par une installation précédente (non modifiés)."""
    connues: dict[str, set[str]] = {k: {v} for k, v in EMPREINTES_ANCIENNES.items()}
    manifeste = racine / MANIFESTE
    if manifeste.exists():
        try:
            for k, v in json.loads(manifeste.read_text(encoding="utf-8")).get("fichiers", {}).items():
                connues.setdefault(k, set()).add(v)
        except (ValueError, OSError):
            pass
    return connues


def ecrire_fichiers(racine: Path, forcer: bool) -> tuple[int, int, int, list[str]]:
    connues = empreintes_connues(racine)
    crees = remplaces = identiques = 0
    gardes: list[str] = []
    for chemin_rel, donnees in FICHIERS.items():
        cible = racine / chemin_rel
        contenu = zlib.decompress(base64.b64decode(donnees))
        if cible.exists():
            actuel = cible.read_bytes()
            if actuel == contenu:
                identiques += 1
                continue
            if not forcer and empreinte(actuel) not in connues.get(chemin_rel, set()):
                gardes.append(chemin_rel)          # modifié par l'utilisateur : on n'y touche pas
                continue
            remplaces += 1
        else:
            crees += 1
        cible.parent.mkdir(parents=True, exist_ok=True)
        cible.write_bytes(contenu)
    return crees, remplaces, identiques, gardes


def supprimer_obsoletes(racine: Path, forcer: bool) -> None:
    connues = empreintes_connues(racine)
    for nom in OBSOLETES:
        f = racine / nom
        if not f.exists():
            continue
        if forcer or empreinte(f.read_bytes()) in connues.get(nom, set()):
            f.unlink()
            message(f"  {nom} supprimé (remplacé par une commande Python).")
        else:
            message(f"  {nom} modifié par vous : conservé (il n'est plus utilisé).")


def ecrire_manifeste(racine: Path, gardes: list[str]) -> None:
    """Mémorise l'empreinte des fichiers écrits ; pour ceux que l'utilisateur a modifiés et qui ont
    été conservés, on garde l'ancienne empreinte (ils resteront détectés comme modifiés)."""
    precedent = {}
    if (racine / MANIFESTE).exists():
        try:
            precedent = json.loads((racine / MANIFESTE).read_text(encoding="utf-8")).get("fichiers", {})
        except (ValueError, OSError):
            pass
    fichiers = {}
    for k, v in FICHIERS.items():
        if k in gardes:
            ancien = precedent.get(k) or EMPREINTES_ANCIENNES.get(k)
            if ancien:
                fichiers[k] = ancien
        else:
            fichiers[k] = empreinte(zlib.decompress(base64.b64decode(v)))
    (racine / MANIFESTE).write_text(json.dumps({"version": VERSION, "fichiers": fichiers}, indent=1),
                                    encoding="utf-8")


def migrer_configuration(racine: Path) -> bool:
    """Met à jour donnees/config_comptes.json (nouvelles règles, vos réglages conservés)."""
    config = racine / "donnees" / "config_comptes.json"
    env = dict(os.environ, PYTHONPATH=str(racine / "src"))
    r = subprocess.run([sys.executable, "-m", "comptes_bancaires.migration", str(config)], cwd=racine,
                       env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
    sortie = (r.stdout + r.stderr).strip()
    if sortie:
        message("\n".join("  " + l.strip() for l in sortie.splitlines()))
    return "version" in sortie and "déjà à jour" not in sortie


def preparer_donnees(racine: Path) -> None:
    (racine / "donnees" / "exports").mkdir(parents=True, exist_ok=True)
    (racine / "donnees" / ".gitkeep").touch()
    config = racine / "donnees" / "config_comptes.json"
    if config.exists():
        message(f"  {config.relative_to(racine)} existe déjà : conservé.")
    else:
        shutil.copyfile(racine / "src" / "comptes_bancaires" / "config_defaut.json", config)
        message(f"  Configuration créée : {config.relative_to(racine)}")


def creer_venv(racine: Path) -> bool:
    py = python_du_venv(racine)
    if py.exists():
        message("  .venv existe déjà : réutilisé.")
    else:
        message("  Création de .venv (quelques secondes)...")
        try:
            venv.EnvBuilder(with_pip=True).create(racine / ".venv")
        except Exception as e:                    # ex. python3-venv absent sous Linux
            message(f"  Échec de la création de l'environnement virtuel : {e}")
            if not WINDOWS:
                message("  Sous Debian/Ubuntu : sudo apt install python3-venv")
            return False
    executer([py, "-m", "pip", "install", "--upgrade", "pip", "--quiet"], racine)   # non bloquant
    return executer([py, "-m", "pip", "install", "-e", ".[dev]"], racine)


def init_git(racine: Path) -> None:
    if (racine / ".git").exists():
        message("  Dépôt Git déjà présent.")
    elif shutil.which("git"):
        executer(["git", "init", "-q"], racine)
        message("  Dépôt Git créé (donnees/, *.csv, *.xlsx sont exclus par .gitignore).")
    else:
        message("  Git n'est pas installé : étape ignorée (https://git-scm.com/download/win).")


def main() -> int:
    p = argparse.ArgumentParser(description="Installe le projet comptes-bancaires")
    p.add_argument("--dossier", type=Path, default=Path(__file__).resolve().parent)
    p.add_argument("--mettre-a-jour", action="store_true")
    p.add_argument("--sans-tests", action="store_true")
    p.add_argument("--sans-venv", action="store_true")
    p.add_argument("--git", action="store_true")
    p.add_argument("--sans-pause", action="store_true", help=argparse.SUPPRESS)
    a = p.parse_args()
    racine = a.dossier.resolve()

    message(f"Installation de comptes-bancaires {VERSION}")
    message(f"Dossier du projet : {racine}")
    if sys.version_info < VERSION_PYTHON_MIN:
        message(f"Python {'.'.join(map(str, VERSION_PYTHON_MIN))} ou plus récent est nécessaire "
                f"(version actuelle : {sys.version.split()[0]}).")
        message("Téléchargez-le sur https://www.python.org/downloads/ (cochez « Add python.exe to PATH »).")
        return 1
    racine.mkdir(parents=True, exist_ok=True)

    total = 3 + a.git + (not a.sans_venv) + (not a.sans_venv and not a.sans_tests)
    n = 1
    etape(n, total, "Fichiers du projet")
    crees, remplaces, identiques, gardes = ecrire_fichiers(racine, a.mettre_a_jour)
    message(f"  {crees} créé(s), {remplaces} mis à jour, {identiques} déjà à jour.")
    if gardes:
        message(f"  {len(gardes)} fichier(s) que vous avez modifié(s), conservé(s) tel(s) quel(s) :")
        for g in gardes:
            message(f"    - {g}")
        message("  (py install.py --mettre-a-jour pour les remplacer aussi par cette version)")
    supprimer_obsoletes(racine, a.mettre_a_jour)
    ecrire_manifeste(racine, gardes)

    n += 1
    etape(n, total, "Dossier de données")
    preparer_donnees(racine)

    n += 1
    etape(n, total, "Configuration")
    config_migree = migrer_configuration(racine)

    if a.git:
        n += 1
        etape(n, total, "Dépôt Git")
        init_git(racine)

    ok = True
    if not a.sans_venv:
        n += 1
        etape(n, total, "Environnement virtuel et dépendances")
        ok = creer_venv(racine)
        if not ok:
            message("  L'installation des dépendances a échoué (voir ci-dessus). Vérifiez la connexion "
                    "Internet puis relancez install.py.")
        elif not a.sans_tests:
            n += 1
            etape(n, total, "Vérification (tests sur données fictives)")
            if not executer([python_du_venv(racine), "-m", "pytest", "-q"], racine):
                message("  ATTENTION : des tests échouent. L'outil peut fonctionner quand même.")

    if ok:
        py = "py" if WINDOWS else "python3"
        sep = "\\" if WINDOWS else "/"
        message("\n" + "=" * 64)
        message(f" Installation terminée (version {VERSION}).")
        message(f" 1. Déposez vos exports CSV BoursoBank dans  donnees{sep}exports")
        message(f" 2. Lancez :  {py} actualiser.py      (CSV -> Excel -> graphiques)")
        if config_migree and (racine / "donnees" / "Comptes.xlsx").exists():
            message(f"    Pour reclasser vos anciennes opérations avec les nouvelles règles :")
            message(f"    {py} actualiser.py --nouveau   (recrée Comptes.xlsx ; les corrections")
            message(f"    faites à la main dans Excel sont perdues, faites-en une copie avant)")
        message(f" Autres commandes :  {py} comptes.py --help     {py} graphiques.py --help")
        message(" VS Code : Fichier > Ouvrir le dossier, puis acceptez les extensions proposées.")
        message("=" * 64)
    return 0 if ok else 1


# Empreintes (SHA-256) des fichiers installés par les versions précédentes de l'installateur
EMPREINTES_ANCIENNES = {
    ".gitignore": "ce576bfd538761da00222a0fcdf8761c2bb10ed4c7014f919c33acfb79633bd7",
    ".vscode/extensions.json": "91d69e2c80a6130021d7eaf7b12de93a00867b601ec9b5009000a73c8bbeec89",
    ".vscode/launch.json": "b42be3bcbf1c3bc7f5447f889404aa40cb1568835e833284d645944f0ecd08f4",
    ".vscode/settings.json": "bfb4094c4cb29dd2b5ac1aa9033120a3de9db5ab5c49843ce848b04b40b6541e",
    ".vscode/tasks.json": "e490d3e377369b00c9d76314dc1d0b6d45dd98015acf7051bf8aabe0fd8b9582",
    "CLAUDE.md": "0bd4f4c0814e19c9b449ea6b40e7b66a01aad002630b24fb3297624c2fac82b1",
    "README.md": "4ad482d66ac4282bb9e2c2bd8ed818b4c8e6ea0066ad9195322982becaf63c05",
    "actualiser.bat": "ff523efeb96f3bfca563191876b610f80642cbe57800891eb0c843ae17882cfe",
    "comptes.bat": "f867a06bc085d3ab9a244846b040567e83eeeab8c04926b4ace2f8e9dd6446cc",
    "graphiques.bat": "1219635d088cca8ac3099f140d7d4882f23dfa7e240b2b1c3f65e134282fc304",
    "pyproject.toml": "8c4fe98a923e4beb8fe6dabe8a8ca9ecb5472f4b86e37a092e7bfe16ad7b6234",
    "src/comptes_bancaires/__init__.py": "e6b9f232308fbc782904a9c2852ce213ca2529194433f1be19808a5f253498e5",
    "src/comptes_bancaires/__main__.py": "697e63c194dba4b379fda379159a55c49dd225b1c131c210bb478a35129eb293",
    "src/comptes_bancaires/config_defaut.json": "81346258980c3883f8dde90902ef4d4a89716da86b5921065f55a64e73215322",
    "src/comptes_bancaires/graphiques.py": "6f70fe26b5ba6a5e4cbdc698044970c8b0e91d18e742e740fe6d457175748a45",
    "src/comptes_bancaires/releve.py": "654584e107443a99203dc871cc0efdeb2b59f4914ba589265d3d5b9464fb7c21",
    "tests/donnees/export_fictif.csv": "5a94dbbd5f315576cae3f6e3a66ce2d4f760ed192f4d5d4e7fe0ae004a261976",
    "tests/test_comptes.py": "8ac131f97a7b3597bae8b627163edd644afb34f68d702613c29f1309ac5b0fd7"
}

# Fichiers du projet : chemin relatif -> contenu compressé (zlib + base64)
FICHIERS = {
    ".gitignore": (
        "eNo9jzEOgzAMRXefIhUbEonUsXOXbr1BlIKB0BAQNlE5Us/Ri9WhotP387f+lwt1nWL8vJHUjAvJjCEIXFRENbjReVKZk5g+"
        "uws0WZBMCadj1J3nJ+IMpa4pwYmRmMz/cF+W+hXoJdLzGEQGKdMP9wQo1H3jfoqgE8ZkwNp5q13do7VGDrHrKh/byYCet5xs"
        "d1NwWdv2gMfqQ2Og8cQmR94isQvBMa4LaH+Q/LA3S65LLnKVzj/+Ap9QWmA="
    ),
    ".vscode/extensions.json": (
        "eNqr5lJQUCpKTc7PzU3NS0ksyczPK1ayUogGCgMlcot1CypLMvLz9CCUkg5EPDGvJKMovyAzWS85J7E0JVU3OT8lVQkoF8tV"
        "ywUASwgZsg=="
    ),
    ".vscode/launch.json": (
        "eNq9UstOwzAQvPcrrIgTatOKY68gfgDEBVXIsTepRew1a7ulivJBfAc/xtpBzQWJA4/LytqZ2Z0deVgIUR2AgkFXbUW1qa/q"
        "TbXMXYWuNV0iGRkLDD5yV4ihVMadtJAlN+9vDXYJSGyFsR4pCg1BXN89lDmFG0++cDU0qfOnGSB4SRBixnqZnNrPkEWd+qJS"
        "aH2E8NRIp6QhCDNJUjdbKx2NzgGENbxmK2F9WatwOAuYsFrJtjVqD1R9Nnfnceqo88KL4Yj0HLxUcIu9BhrnjRxLwMmXcRE6"
        "zgf0PZA1TvbTxHH5fVKs83uTj/+zlOqvdvzfhTljwy+F/Idc/NmZnpB92Ml6a3oYf/0krrvF+AH9a97p"
    ),
    ".vscode/settings.json": (
        "eNp9Tz1rwzAQ3fMrjOhkWrXZSrcOKXQrdIyMUaWTI6JIRjq5Ccb/vWc7ELtDFnG693Hv9ZuiYO0FD8FzDUZmh58eIbYR6P2S"
        "eGBvBXvof0M8plYq+AhOQxyE4B34TohvFW2LSYirCZyBPS5MyeVkvXRcKrSdRNj5zsbgT+CRrDFmWNMTWt9w+tK08/LHgb7D"
        "y97immmkS3cs32OTiLUnAlHGTWI0V5PCWAeJg1dBk2JsntG8siV2Vi5rIKifHcryua7bi5LqAHW9SDpjJYemebLehDV0TVNP"
        "un9QzMasANoPU4b9XKm6nQdtMUSSOIi3XkWx3b5MUzVqN8Mf28abFw=="
    ),
    ".vscode/tasks.json": (
        "eNrtV81uGjEQvvMU1ioHUnUh6pFDpah/qtT0UKJcuj0Y72RxMbYztskixLv02DwHL5bxQtglEMShEqnEhZ/xzvizv29mdmYt"
        "xpIJoJNGJz2WvOtcdC6St9HquRs5sv2kP4zNqk8yKz4AFR+9FD5wJR0g67EP/RuWvmefSgEq/iiQ26G8C+CqYJWnn1qIjhaN"
        "ANdYEGY85jqPa2eze4MjZ7mAz0blgPMs60xAT7KsL1Ba77LMTv3Q6A6UUIfgWNRYl5Y1vI6dJiv7r7WDkjqU5DGrPfbC6FYo"
        "ugOpu8v9n0LO1yGN+9cBrSda3LOg9y8E3Panix4oGF9xL4aA8Xrq8xdogt2MPJJLrIMgVb6+WlqQ7iPc8qA8rXoM8LRRq7Hd"
        "Xn2AZggCFw9ce6aACcWdg4CsrU2YgFLgGC7+FvR9/jr00jh9mlYgeTiJaCWi/bx/HVuDnliPrMa6kAPLjdYArgtlXDtaTaCg"
        "1oN7RvAzcN03HeEmJ7IPIvvLus6zNmUJSmQ51y4mueYTWXBPaX60nK670FZOL8GeWD6I5UvN1dQBsxyZMNpD6YG1I5eMByaG"
        "Rpbn/0NObyiALw/VNJ7NpLbB9+IJ5idtHKSNK/AeSQfst6GGrngUyK0sqOrbIKkBrOgpFZ30VUkkTeWqU6VLyLv0swH/JIhD"
        "BHEN7ngdPh03WSQfArNB+t2JxZdGgaS6rNbTq331QFIVxB1joKy2rsjdplqKUd+j1HVKJTm4isnVnHkde8fiD1sVYUx2nLRB"
        "63eqLSh9wI1y/YPw8oA0VDSt34x0El3TdE3PuJjrTWOf/BYPTcuNpNHEVAE39rkxU17AtnDy9VTUBFhfYGv+COOWSD8="
    ),
    "CLAUDE.md": (
        "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"
        "+Ad9NDrm"
    ),
    "README.md": (
        "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"
    ),
    "_lanceur.py": (
        "eNptU81qGzEQvuspBvXg3ZKu26vBhRIMDbSOqUMgBLPIu+NYQSsp+lnsPk2veY68WEdaJ96U6rCrGc3MN/N9Euec/RC6weig"
        "MV0XNYh4yFuhW/QA9pgsG9BXtJ0m+8EJu5dPceQSTYhCSY+OXBW7UuBQpbpULphOhBTeoQ6gEHzjpA0gemySuTqGvdHQIqgJ"
        "6l46o/UQS/sQUUFR9XRQwoxZ4WGL3kiKnxCo7NH9m1YxTlPJzhoXwPjXnY9b60yD/uw5erZzpgMrwl7JLZz8KzIZ+/Xt8mq5"
        "gHk2i7reSYV1XVaO4FWPRVlZ4QiO3S6WtxR2ip8Cz91ytrq7+X69rE/H+TeFgq/z9H5q89gVHpCD3FGjlRYdwnwOXAcOqDwC"
        "30p9CuQlY6zFHbRC+5paSSBFCZ++wtJonDGgRYOvJRCZeoI+QGKrfXl+fPmTsyA3dvEmzUgKa5Pvf5rknCoxmgA+pMPWeC+J"
        "9zYxjn4KbQSi9hEDJNQYJN2El2d4StrRB0ivME5sTHSElivS3Cf1Ko+BBhRRhYJfXv9c3SzW9cAqvwAfXDEYZZkTibOsDKlY"
        "WYc7eRhpk2hMjJ89A0FpOQzR6dca2pDgZ6VID+mDH8cnAPJSU4t3tzMzA1IHZ2IvtjTfDDKxv9Pd3BrX5tcjtQ9CKXoXfGj8"
        "xPn83L5wD/39582of/YO+Hx1q4ZKFfeJjFHT5UDPUJiMj29Fv8w2G+LrLwQoS6Q="
    ),
    "actualiser.py": (
        "eNptUs1u00AQvu9TjJaLA41LK05BPbRRkbiABIhLFEUTZ5wurHfd/bGSIt6FY/McfjFm/ZO0BV88Oz/ffPPNSCnFNxsDkIFo"
        "CApbVWg2BDOAeg9YhIhaeXJ5vRcAFzngD04nCDZ60ORh/vU7cP7GGkPkz2lXWxc8YIRCo/cU3TE2t1UdyOc77XeQeTSeQ3Gt"
        "rfETeM/wlzls24NpHx112FuH9Z26j2xS6Dw2NhzbpFJNYLBRWwzcIxficx0UI8GMgaZTw5nEJPrPUeHaQwI9sWr/QI0uKJfo"
        "Pxsnw7rWqW3n6ZB0slz7uOW/gH8+jDtAUyjiQZlk3R4cdmzOoELF7MsSix6usI7Z9FRLVKxHP87triA96bgnZaZpUuYGwFup"
        "sZ+c38+mFh/IVfQAz4Q9oQE2aEIaTzM53qGQvG5VpQ3BqvMxhhitPFWuNK0aMk02EWOm33uAV6zDPc7g9t3bSyFKZ6t0K6nr"
        "as3lqBwPMhSc1nbGymtq6GU9uq2Hq4Scs9ksLmZLMV7O1ZFa/uV6/vHTLZyDHE5IJntI5ElKBuW7UD5kgzPfarvO5Ou88I2c"
        "TGbdrlIb2qmQlfI6FtFAqYo7Ra6/3STXr6H8N9/9pj3U1tPDdA+N5dUNtG5sdN7eoPmZS9ZmpM50F3JQQp6BD26kkqgOPJbw"
        "BrKFPF6lXAJzf/IGZaDThLQnWCwnotct5+sxvIr/dZueZJZ9gxH1yf28ROboEOAeJ4Cxz18DH2Cx"
    ),
    "comptes.py": (
        "eNqVUbFOw0AM3e8rrLC0SJcixMTAUnWAoQyVWKPj4paTLnawc1H6SXwHP8ZFTRBDJOBNT9bz87NdFIV5bFqWDqFGBRxGrrA9"
        "vEDtSCEi+OhUMQmsaiZC1M2Wm7ZDLYeow7o0BjLaM/ipnOmsnPw216XX/q86sJY49egSfEPQy+cH/szzDztHLp4VYc9JJHRJ"
        "cKHZ2nC5hFjPdAwnWNo3tzGdIuYj7QaPEewDPB2e94uGbxhb+B0dp9ySd1PgtgtMaor8mEscqKIjP+5rZlaOr6kiVj1Sv1ob"
        "cxRu5snVaxa5IDmDYMQeYfJpXCCAKyB+d/ewu7u5NWasZYMvxM6pZw=="
    ),
    "graphiques.py": (
        "eNqFT0sKwlAM3OcUoW7sokXElXtRRHfuy7ONNtD3Me+14JE8hxczpShuxKxmJslkkmUZbMWElm89RWSXSEyd+BJxvjsdDzk2"
        "KjfPRyAXR2RcxI4wkQ2xBECtcMfrx6JUVhS+H4TlVzeQsG8IreeotPY2JMK91+t/d1vqAmSamm3wkrDqjKupF4A3KseMVUfV"
        "QG6Y5wAX8RanI7E665BhUbvr19uTlzXsEGfo/M2scbNaLAFGTU1eXrFkBg=="
    ),
    "pyproject.toml": (
        "eNp1UkFu2zAQvPMVBE8tEBFW2xRBAQlIUh8LFCnQiyEIjLSymVAkQ1Ku9aS+ox/rkrQtHVrdNDuzszvL3fMkVV/42QcYG+Lg"
        "bZIOPK3ojnkIkw3GKF9Xnz+xhmTus+heQfdIWTF4qrUjBMEI2VlnXqALDdFihMjszGgDeBTrTkQHRo7gvDQ6Vjf8lm8Y6cF3"
        "TtpwRp9AwfHPb097oGc9fff44yd9MJPz5kHo1/c0dqHbUwfqhgot1OyBWuFQoAOcAlAIdO+EPci3CfXflQlqZrio6PNkT9v7"
        "r9+2fOzZdfvCzuGQZ6irj7xMo1ncGXQncziE4scMgnY+qcRiNxm0QvcCI/uAK12g5FpXt7y8Q6xZEuImrStUsXZo0O+YboCT"
        "gA91hTLK3DQMdYVhsXWHnBlqLhktcbfXuLmLYcKXUUjNLsxiFcw/RUv9LCS7eG2+urvF1yD2yB2k7hvy6wAO8utxXRrzP4Ki"
        "F0FcZ148kxaPN8h928MgpsBfvNFpfRiwue6gPdePZa5dbXJaXGrZ5lwxlYhYEQ65c/zzKMgXjniC+dIjhtwQJTUUCvQ+Ecpy"
        "Q/4C0X0UVw=="
    ),
    "src/comptes_bancaires/__init__.py": (
        "eNp1kE1uwjAQhfc5xWhWILkRmy7bBZQFm9L/TVVZxhnArRO7thOFI/UcvVjtEBDlxxuP34yf3jeI+ESamt8fDwWBNKUN5GEh"
        "KimUi9Vg8vwGY1M7b8ai+hrC1S2YxScFDw+bsDYVg2krSQMFWDlh1+q7Jp8jYrZ0poTcJXsCVVrjAgweX+cvs7vZ9J7BxFRL"
        "teruQG0gBnNLTgSVTEUl9MaT47Lv+qhZq5P9oZjBuSOFlrWOg8EEUbcMSLqIwylFZX2YrU+MsJN1GpG+uWDqjS7Ic1WpoKIp"
        "L8jWyqcPwyzjvCHnY3TO4QZwlF/nI4yq0LpT3nFPjgxwy95XHUqq9/zpcbqBTj3dAf7Li0fs6dMhfXqf5U+N3QaOLC+T40f2"
        "B3jWvbc="
    ),
    "src/comptes_bancaires/__main__.py": (
        "eNo1y7ENg0AMRuH+pvjlCop4FjaIHOKIk/DdxRgktqHNHCwWUdC86ntENKibBtD2mGrBwzBWa6HL8yVllOy6gJmB7vx917zJ"
        "GjgPzHI5k/LWe+iZiNLHq4FdZ90U2Vr1gEkuKV3t+vQHRH4n7g=="
    ),
    "src/comptes_bancaires/config_defaut.json": (
        "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"
    ),
    "src/comptes_bancaires/graphiques.py": (
        "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"
        "6dSceNQ+oPd/DtO6OQ=="
    ),
    "src/comptes_bancaires/migration.py": (
        "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"
    ),
    "src/comptes_bancaires/reference_config_v1.json": (
        "eNqVV91u2kgUvs9THHFRSAVpWqntKtKqMo5TXJkE2QR1t1TVYCYwK+NxZ8YR2SZSb/ct9mrV7N7uG/hN9kn2zNgQg01aImLs"
        "4/P/nZ/hywE0IjahUURl4wQ+HAB8wX9oCDqjS6Q0Rq4P7nkwhL6FX9a5vvHdi/NG2/AVwoaTCbqgsYJrKiT0iWA8LrjUTWJY"
        "CiLS7tqPmLJ/8V3Pc6Br+b7jW4+Ysm8E07RNMyvqI3YCZ2BBMHQc33Pf9oZVEwGJCFqBQFEqIjabq/2MeJdu8AMBeCmT0CVC"
        "UEE2DcRpFO3QbY/HR28G+mLpS7+q36eLCU+FLLI0sPpb3gckVtn9Luf7vQEMfCcYWkOEOgCNu3M78BzXOnW+Z22RqpRWQXnc"
        "4sD3RgUoZ37VwhkXV4Sp0ss1DHMiZlTCGVtiCe/Q7jtD33KHcGp167xXQuumMsm+hahk08CIUQgxOnR/J9gupsfHvrB7zm1+"
        "Px5PqpZcVCEWRIRzTMSmlXO0IJhKxU4b9kV/YAVB59XL45e3geOdQe8F2E5NcdnoKovp3hacgWtf+m6NRidhYSoY2Vtl9+LS"
        "s87fOr7r3J5dXPrnrlfV3uVpROIZFWx/nwc9y6+p/wFWxYKEbL8itN2h+2tNPpliv0OLpIonRCgyo4dbaodYHTLhQu3SHJzb"
        "ZzUz5oH6w5rOPPd9TX9EbNlN5b7Kup6FH9uq6bhuRPBjE7F3qE7fHVpvqxqRCC1stQU8g0kq986hdQkj17l8D7bVHziXNT5b"
        "KWC3pkuwySKh6bbrHmeSiZ1D4tSxrWHPq1ttpzQkah5V9tl3NDrBwLIdGDn+0LWtmsp3ZEJCCiMqFAtJBC0q8YtMKwX2HUM9"
        "x/Mu4A2Oh4uqkR7ecEtKXr9fDuAj0g11dQRoWIkudDPOjVC5Ec1z0Ufm/gE481heyY3SKaDRJfHntKDbGCNJl8XDxggvSHyF"
        "n9GBPZc/rBJhHkb8Zv1iY0wbSvYVwohISUWjiBFBpDOOY0Z+muTOnJhMNj4RNjW5ulQsYjK7B0nTKF9nkgFJwzSmILJvs4hC"
        "c3VgagISpzzWVwo6gUdgRyj8MzRtgvkxtprQSvSSL5FgQASqps1DLRgRyL05yv22IqYNE6Uzd1LNvU+lIiZQ2UZJgdcpkyFX"
        "8+wbapH/ff2rXDJ5drK/czglUAUkxDUnOZ5BZGXDGXbnc8qSgt9Ay65yQbyTimkrFRN2lE6wszHhPGTGeWgZ2TbM+WRy09bS"
        "6NthRbLPJ5h1KtqQ3Uc0VIIvsvsYkUXKNLsPuTDqirgqtVnsbwkhkfP6eIaoOLtP5jzGV60rrDMdzULbpYfVU8RmWUv4nHLF"
        "pozGElroGjrY1hNsFcxWA5QEcWLMUqrhUSTGRm+RawwEAdMex7sUDLJ7DF7WvMFxPDFx1bzrZfcTikGYom3Ns38VjdBSiKOQ"
        "xbOVqXLLrJekTkZEJjrPuihKi7JIcOl4J+EKfZe6cGXOs8Ud8IhNsWWRAB2wUlVU2UbLe+wayW2wT5+djk4xm+w3ni7bgFeq"
        "6isY3VzjENN6nAsJHY/UVVtTpnkCJDyBB5xqsol+c+Tpc8Vr3ga1kVsTPQvyJD0BVaq5HSVWpCR3d521crIOzLhvhBxDXCqc"
        "W3RJF8n2j7X1+HLeO308pUP2J5ApSfC0eYRATVL17IrFOMbgiiPiCiz86/T7nXfvoMXiMMJtfARNXCdmqp1gVheL8qBrw0Jj"
        "nuAwRxf17yJ8oPEUg4dIg5C7Z3Sge7jTxCezULSyL1j6iBluHJ5eY1R3bWg+beKQVDzFFkFVhhWF0Sz6bIT0dipY2sBT+DBe"
        "LZZxow3jxjvOYjVufDwCDxWs85NPZHateyAhOLzQuTm9xgE+pwJa6GuCHuKsKQYvDntFYUZmMT08KvZjzBemukioe1ZCgK2q"
        "Oj5J5iT7JyqY1qukzFq8MxnX9BfHL151jl93nh8XbxCFDfqrlbY88GIrt1fL/QHnAumNcxb+Hh+Cb+EJ2CmOFlunS9MBte6X"
        "f97kA8FQ7/T148r8BpDFpixt4E3Zu1yKhLgq8JUSKd08q1RqFCe71OzXOCdaRu4Erkgk6SHCvyquNMbKM+u/rWtBI6wLBmWR"
        "Q+oq1AseGWlsqgiKw8IWluUjxBZ22R/oQPywTrbhe37cef66Bj5N/+mH4PteQp+WPazJpUlKcVI7uPsfM2ZLCA=="
    ),
    "src/comptes_bancaires/releve.py": (
        "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"
        "OBfMZzjEEzIc0nSGQ2RYhkOX587cy97/AzNud+Y="
    ),
    "tests/donnees/export_fictif.csv": (
        "eNqtlN9u2jAYxe/3FJYvpm0q1DHQtfJVSNM1EoQooVxwZ8AFa0nM7ESaerWn2cZz5E32JPsMoclYJzZpuUhOvvjPL+fYxre8"
        "EGiyrXaaF1LlmOF9ZcZTUWrMRnIh0rTaMXxUKCnXa2he7TDzeFHt1kpLwXCjUcS1yAuBWaLSlWCeyjJ451JDs7DMoLNCUNza"
        "Js24x8qhU6QkdFmLV5TQqw7pdXoOa0nsufHURz1ySXqX9AoF4dSPx27s3fvIG75zaK+PWQAQOuN6uYEfcFO5p7B/yfAMMD+V"
        "qpArKfLccnQouSAEM0YIceCicLGhKrVRQ55/ZM51j3QJC1X+jEQJa8kaybmukYaxmyR+HPjo9gFFk3jagOGh5sYI69VtiSKl"
        "C6jFwhS81DwvzAVacA33lTRLVWyq759KYX58+WrtUtJIbZAokIF+UhjL3h+cYR/8Hfv7/2dnj/wrkkNZS+IoHs1Q4kcuSu5i"
        "WBZ3McPTagdLZbtROUz45lF+FtaITC1kKt7CrO5CwfQWzKDXqGi1tkjOwaUCjEbL1qL8E+INHZwiNq6BxLMgPhB63ajrdsdd"
        "dOujEQogdh9bVzRKxJYjb8uzfcLZwg5fAz5qLg1aCWQgc7udkv2TYeeMdTf9U+vIgLVknSbp12lG924MbJ4fTmN31EoURxuu"
        "M/gCPBp2/LHCl3JvbMoXCs4FBSY1dB16ZrHd/LZRSJMsyL1tQZhM0diFhxtaEQeTsLYsyCGfMZe59QWEPpxL8Kk2rtpl0rxY"
        "6TgIqM7gDU5TJc3RArJJNZn6fjwKPtxP22EmhRA6letN8QuCFtW30rxcomhwhoqeUtEOvWYtWWdKYYdSm+l8PgcPvUkYPrTy"
        "nD89gX1L2AQlvMFoaHk8lo1dYS+Vzh0e/Weyny/Y2SU="
    ),
    "tests/test_comptes.py": (
        "eNq9Gl1z27jx3b8CAz8clZNpSU4cO1PfVFbknOb8VUn2zZ2q4UAUJDOhCAUkZfkymbnXPvcP3FMbtz+hb6fpH+kv6S4AUqAk"
        "2+k1rSbJkMRid7HfuwiltMvjJCZxKsmQx2QoomhxDw8nrUa3dd3sECdBgF1c4Dze5fOpkIk3CvwkGLl+PCu55JRFPpfkFZne"
        "ITCldCuYIBh5G4toayTFhAxZwpNgwolZwXe9MmXJTRgMsoVLeN3K9muEWxrQF5MpvHoDIMcCCUwaKOcPVxfd1utW87xMGiIa"
        "BeMyYREL72IuPV9ECZ/DvjLxWeinIXxLRMLSeZlwXwIej899Hpa3yOM/TUxjBBJmFwkRA8rhASZdyUM+y8/thZpiECcsSsrE"
        "i8RkIHlZCcRj/g2Djzzh0RDAJgKEHG9ttRotcqQk43gg+pB7XsmdMsmjZKvRuYY1hNgl1GiJ4vOapujW2cXr5mlTg5v9CBlL"
        "X+1Y4918Vccd8hFLExc1Sre2tn6vVeOOgnmSSr4Fy8QfjZ1kMvVQoyWy841RxislWbAK/UomYrj4FHLytREW0SzCexqRTF3k"
        "17+Ta+ajZcUEKZFf/0GGKam+IAz+PSQTJmMXTQ2Rg/AYHAuZc0PBhrGjjwrCZ0MPETo88sUwiMZHNE1GOwe0VMp39rKT0z4g"
        "+UArlUoVfjX40VeENu5kEIacfixsMFaltvQ+UFAjwhZYpmWQHqh1LGTA7VVcGPJBmuDHWqW2v1PZ26m+oI/ZIB0FUQH8UKE3"
        "jL8i5yICK6KSj0P13us/ik3yyTRkPjrD3VRt+EDbwDNLJdilYlbcsTGcGpAy1A9868qUf+wrtP4NnwQRnD1TuGUqhiljK0to"
        "91YGYORKHUpVw3QyjR2UKNh8BDEIPCD2g+DohIUxL+HHFaUpbJKDyUXGulxwGTnm0tE0Sg/YpgBCYJ8lbYwCGM8c1wEPKqPt"
        "auQrQcIR1pohLIAG4kQixn9jbySV18SOIcFiiD5J5t8O3amSg0qlXKnQEjk6IjtVeHMrG0Frv/7ycm+/vLevQWvw4u7tbwR9"
        "Xt1zDw80nH7eCKYBkJ7F+TLieD6DyLbC+XLZoY16u9sk1YPdyt5ubZ8ct+udTrPdapLG8bNqbe851fHLQeMsk70yqVVKimTx"
        "a/Wg9BCFy/bpNek0L+ukc9JeQ1etbURXK9nnCbmP6lY6BXUXTxPySH1ELIf2AnzsVfouYl6hUIM/B6VN3rNNEhks7jEeTcM0"
        "hmTjBzzCwKRe5eLex9C8mYqn4xvQorUDcrK4n0F0kARpamcJILRIjCqCjIQEWwU/g/0EQqRwIVNyCEZqewsBIQ76N4t72l89"
        "rUKjDQgYHBIWho5w0dvV5nORSnBHEBhVy8KFiBQk5BtSMe++5EP4gIazZERjtWl9EO5E3PEIokDAJ5hSbK4/km8wpDbQwMhx"
        "yFOOYesafA9B8fkSxBUuPs30Bx1k4xEKIMJA8ZQQ6gPMeYouWk6BNUCzPHBDxQmobII5ZjY8I65DsMLNmO8UmMo11ketk5iF"
        "GuIhpmin22y2T1tvvu1S9TljErPp+5QX+HIMugymnOFX3CqdObRjSGbpB0TVhkIiSmNzSNB8IKJHWDqrt8479fPPYUjjWvJj"
        "3i12mu/TAJYl5ISsyEGWzhRgUewP8dO4rJ9t5AXStEpGDnURJ+TmgorbfDIAa421lmPIT2DvtufrUO3xMQu9WITDDK8VBrbB"
        "J6DClRH6GsKHEM11rEdHFmDoEN0hXMIyFsMArTARNoLyBBwM6lT4yojGvOrbO9W+q7B6vlYXMl492FsJ8UXQMY+4BEY+AzTT"
        "9VoQzwoRO5rzxMvSO0rMkgKoD7EHjxlyHtyfNpscnZuXbaivYhG0FjsVkyjg6gHE+cU9Rh14q1U2o84d2BQkj5To22RZwmD9"
        "qXeQKVTT+CXkeXm5DLTeXvWxULMhL0Dm2auWCoLIMBXkkHcltr5MX7JsS1A7usTQKgqDcQSig6AZLstM2n9FQsVdiNytNzc5"
        "lnIhspc+FrKCQt1b0U+/R9+id2Ele0RefNaGIQeZxlxvQcu1dWCnJVWyv4SyYQPaXD6rCGsrNq46LY8Bl4kH3h97Q5EOQhGh"
        "oftCgh8n4ByxN2ZyyAuyyLsRLdt5GM9XStaGqVVxSYd6uynUqHDNIFRMcGkKVBNZljxA3IxSFbfiYIKRBWIOMEyaiExBqyZR"
        "wGGnd/MwawqxZfFuhXw3EOKdArsdYIVqf3aQC03xFs3jdtCjF9PFvWTq8Cb/g7tztD7S890ZC1OujMZHo7mNe9V+3yBwgZ/Q"
        "qZXNBjeIhnzuoDTyvFeChqyqWgywpjGocvELxr8JC7JwfztwYzbjFmePyK70RHsNcpws/jaBfAcNbRpBPEDeg9gU3nAUwKZC"
        "Mp6+2EY7SyqrVZDeuV726e9Yk60WAJvOWygqFA+9vC3UbrMs5hNs3lPlwQUFgSg7d1Fys/gUq2S++NNMhCmuqdS++ISNG9HR"
        "dnGv4Bt5lwkvXdWnFRw6p/W7IxLzxNlgMG58w3kSsQkIwfYpbXle3txP0yC29fZ/cx82w9Bc6N0B1DUd45P9+8apzJo9TKWy"
        "mt9IQ7kQ+DQ4kUNNOlQqUXVvnBe+dkMOj7qxNrKztYYn7gFCZTmKN/VWRiIFSwOVamCNy7sRMs4yCO0ro1YwCsVmGBufdebb"
        "ILnx4hRqG/B61ai7A/aOllzlUtDE2rYylmx6E0DyB+zJ5BEDeWACttyfBTxV+YDSEJ2ur+GrqkrslQ2E0ObGLi6ChKdQH4gh"
        "P6JQGYKXchOWTAI2OB9Try0cyx/J8CsMQD6DOlPPblD1Gi22DXQaiiS8W34Fm7rlsig0PbvzbiGyitvYWZUUmFSkDYqp2RzY"
        "yyB78N1kntiG41jHj0r2GIXO6VrIK04PnV6cSBsBfabolPq6pa7pUA72YKatrmRBDFm0cxcnfNKcB4nFydO4wfUQN4him+x8"
        "uR/ZBnwki5KzGnHAqKZSoD1jVwB5yFcFXoiSgAIeGE2tEF764uxkUyaodN7hMBVCbCKDn6CJGcg0KZO8l0K/9ECvPSXFbJhS"
        "O9itHOIwpdHq/oBzlJeHlaoK+gyqiRFUZKQRJHd0paArF5HUDJJOvUOa5z9e2IiakBt9NiSvyCnYcTrRUelhZJW93WpFcXTR"
        "6da/CKbj1kX3h5Or01Ny0r5qdTfjPJFpkDyBsbqfSevi7BL6k539F5UXpFaUW5RAECCNZv0hXO1mt10HNl7Xj0ntRSa6LjkD"
        "Gq1z8tpG1wZtMuxMIHAFQzO1wAZCdRfrmgAGDxS6q7OrS/LssvXjj3Vy9rJt4zxhcUIug59+YppCPm7NkV232uT09euOXoce"
        "LohjbGGyb6fBDNrUenEDdPldclnvtluN78jlt63T1qWuHmbY1jbqJ28u1F6opgKo81c5zwTROu8222cAf1q/IC/3yaHNemHW"
        "9LiqDjNJnDdOdq4vfqi/aV61OzYyXFElDRw+xoSAOCBm5MFTT7G9Wc1R0X+TQxXHe5jaskYVugKEcPQuSrO/epCxgsUO2TMz"
        "lQLC0G6zAbLAPB1RpDU9xlkCxKAMGi/NwE6+EoO3HDi5YREoDCIPPL1PdREL/SVUnIu/imCGwWrKYBNXHJAR1ppPHUWpORua"
        "4dbTq1aHHNfb7Wa7nh3wOmcISaXx8tB/XGaSHIqcQsVHjjHYSGWOi59N8JS0IBWZmylEeaizsgsNSyCPcW4iQnUX/oBRnF9c"
        "XTfrV6Td7HRVtKoaa5pxlqr2XdAnr+Is14ECYcAk/DsMYl9gdY11xr9+/otl8ZC1TDEAxg9NkBkq5d5X6M/BZ6SZh288YHHW"
        "6dD69QXoBkfj6nzff/+92/mu1W1fNb5zT5beX9IzSGsM+r/JkGfBWDccWRGjyuLUfJtViRNEcPBQ9aZmpDXjMsbVivv8y6fJ"
        "ZUGUcab7dA+9C/0cwhuU0MCCcD6vnMwRZdUkfsg7ANMvjorlfnYfuLz0BBAoNCOfZ53DrKovrEpPNwSz6m+7bAQSUzF1qBG4"
        "oUyzxd6yuTC3iiAfPsfLuLOLc8hSb+qd1jkatoFTK2rqMGZxEGVNh7q0vBFQBdGPfeji4ag26pyablCQlILIXr8mPXosA1+g"
        "aizw4iVptbanyOv578clVBYfYjMxVPDDwE8cRWXDepk8e/aB1sNA9eGqI8YTqJEvdLylL33hOKt+9nWjQobjsQjntEVLy28c"
        "N99Af253aWJLRgQbDH3HvGIl/WW1vrw/Y7ZiexUF89+ZzT//jLPA4pghunNkTkgzslqmqo5GYkezylXxakDk6DSTGhmmN335"
        "Z92iPIxRj5HVpRsmVtUUEDPyhrhWuJ2yA711hQKPSxsvWwfPZimaqPGJku77jFOpU+R8GZB1xZgFnOKrQ1oJdh3Ycq6HvKsI"
        "v+5GvaILGTUZL3pw6ObfiGBOVPwlGJiLxcgjxD4nB6+nWvvS6RFPVmOPnVktC8n5WMLav9EZkWCvT/7z3zaBKh/ySZLd3G7+"
        "7wXkv/ttEyynOJmxEMgVbq9xwp5KSInqycO5MeRID8x1lt3M5/+HSMMiuki8Z9BCPa9UobgDL5iIYYrjZ04043hdBgXQaMRV"
        "3TgQclgQYobLbdcbrfOm+1YEEerEySzSnd7Z8v831y7dCg=="
    ),
}


if __name__ == "__main__":
    code = main()
    # Fenêtre ouverte par double-clic sous Windows : attendre avant qu'elle se ferme
    if WINDOWS and "--sans-pause" not in sys.argv and sys.stdin and sys.stdin.isatty():
        try:
            input("\nAppuyez sur Entrée pour fermer...")
        except EOFError:
            pass
    sys.exit(code)
