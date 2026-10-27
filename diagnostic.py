"""
Diagnostic : pourquoi une opération n'est-elle pas classée comme prévu ?

    py diagnostic.py GEG "EAUX DE GRENOBLE"      (mots à chercher dans les libellés banque)

Pour chaque opération des CSV de donnees/exports dont le libellé contient un des mots :
  - affiche le libellé EXACT de la banque ;
  - indique quelle règle de la config s'applique (ou aucune) et le résultat obtenu ;
  - compare avec ce qui est enregistré dans donnees/Comptes.xlsx.
"""
import _lanceur

_lanceur.dans_le_venv()

import re  # noqa: E402
import sys  # noqa: E402

from comptes_bancaires.releve import (_lire_existant, charger_config, console_utf8,  # noqa: E402
                                      lire_csv)

console_utf8()
mots = sys.argv[1:] or ["GEG", "EAU", "GRENOBLE"]
cfg = charger_config()
donnees = cfg.chemin.parent
print(f"Configuration lue : {cfg.chemin}")
print(f"  {len(cfg.data.get('libelles', []))} règles de libellés")

# 1. La règle est-elle bien dans CE fichier de config, et la regex est-elle valide ?
for mot in mots:
    trouvees = [(i + 1, r) for i, r in enumerate(cfg.data.get("libelles", []))
                if mot.lower() in r.get("regex", "").lower()]
    for i, r in trouvees:
        print(f"  règle n°{i} contenant « {mot} » : {r}")
    if not trouvees:
        print(f"  ATTENTION : aucune règle ne contient « {mot} » dans ce fichier.")
if "Charges Fixes" not in cfg.types:
    print("  ATTENTION : le type « Charges Fixes » n'est pas dans la liste des types.")

# 2. Ce que donne chaque opération concernée
classeur = {}
xlsx = donnees / "Comptes.xlsx"
if xlsx.exists():
    try:
        for o in _lire_existant(xlsx, cfg)[0]:
            classeur[o.cle()] = o
    except PermissionError:
        print(f"  (impossible de lire {xlsx.name} : fermez-le dans Excel)")

motif = re.compile("|".join(re.escape(m) for m in mots), re.I)
csvs = sorted((donnees / "exports").glob("*.csv"))
if not csvs:
    print(f"\nAucun CSV dans {donnees / 'exports'}")
trouve = False
for f in csvs:
    for o in lire_csv(f, cfg):
        if not motif.search(o.libelle_banque):
            continue
        trouve = True
        regle = next(((i + 1, r) for i, r in enumerate(cfg.data["libelles"])
                      if re.search(r["regex"], o.libelle_banque, re.I)), None)
        print(f"\n{o.date_texte}  {o.montant:+.2f} €   ({f.name})")
        print(f"  libellé banque exact : {o.libelle_banque!r}")
        if regle:
            print(f"  règle appliquée      : n°{regle[0]}  {regle[1]['regex']!r}")
        else:
            print("  règle appliquée      : AUCUNE (le type vient de la catégorie de la banque)")
        print(f"  résultat avec la config actuelle : {o.libelle!r} / {o.type}")
        dans_classeur = classeur.get(o.cle())
        if dans_classeur is None:
            print("  Comptes.xlsx : opération absente (elle sera ajoutée au prochain import)")
        elif (dans_classeur.libelle, dans_classeur.type) != (o.libelle, o.type):
            print(f"  Comptes.xlsx : {dans_classeur.libelle!r} / {dans_classeur.type}   <-- ANCIEN CLASSEMENT")
            print("    L'opération était déjà dans le classeur : un import ne la reclasse pas.")
            print("    -> corrigez la ligne dans Excel, ou lancez  py actualiser.py --nouveau")
            print("       (recrée le classeur ; les corrections faites à la main sont perdues)")
        else:
            print("  Comptes.xlsx : identique, OK")
if csvs and not trouve:
    print(f"\nAucune opération ne contient {mots} dans les CSV de {donnees / 'exports'}.")
