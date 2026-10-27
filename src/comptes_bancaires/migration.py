"""
Mise à jour de la configuration de l'utilisateur (donnees/config_comptes.json) vers le
dernier modèle, SANS perdre ses propres réglages.

Version 2 (octobre 2026) : règles de libellés déduites du rapprochement entre l'export
BoursoBank et le classeur tenu à la main ; nouveaux types (Restaurant, Revenus, LivretA,
Vêtements, Perspectives, Chœurs) ; restaurants -> type Restaurant ; avoirs carte.

Ce qui est conservé :
  - comptes, noms de comptes, contextes ;
  - règles de libellés ajoutées ou modifiées par l'utilisateur (placées en tête, prioritaires) ;
  - types ajoutés par l'utilisateur ;
  - correspondances de catégories banque modifiées par l'utilisateur.
Une copie de l'ancienne configuration est gardée (config_comptes.avant-v2.json).

Utilisation : python -m comptes_bancaires.migration [chemin/config_comptes.json]
"""
from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

VERSION_CONFIG = 2
ICI = Path(__file__).parent
RENOMMER_TYPES = {"Choeur": "Chœurs"}


def _charger(p: Path) -> dict:
    return json.loads(p.read_text(encoding="utf-8"))


def version(data: dict) -> int:
    return int(data.get("version_config", 1))


def migrer_config(chemin: str | Path) -> list[str]:
    """Met à jour le fichier en place. Renvoie la liste des changements (vide si déjà à jour)."""
    chemin = Path(chemin)
    data = _charger(chemin)
    if version(data) >= VERSION_CONFIG:
        return []
    modele = _charger(ICI / "config_defaut.json")
    ref_v1 = _charger(ICI / "reference_config_v1.json")
    journal: list[str] = []

    sauvegarde = chemin.with_name(chemin.stem + ".avant-v2.json")
    if not sauvegarde.exists():
        shutil.copyfile(chemin, sauvegarde)
        journal.append(f"copie de l'ancienne configuration : {sauvegarde.name}")

    renomme = lambda t: RENOMMER_TYPES.get(t, t) if isinstance(t, str) else t

    # Règles de libellés : celles de l'utilisateur (absentes ou modifiées par rapport au
    # modèle v1) restent prioritaires, puis toutes les nouvelles règles.
    v1_par_regex = {r["regex"]: r for r in ref_v1["libelles"]}
    perso = []
    for r in data.get("libelles", []):
        origine = v1_par_regex.get(r.get("regex"))
        if origine is None or origine != r:
            perso.append({**r, "type": renomme(r.get("type"))})
    regex_perso = {r["regex"] for r in perso}
    data["libelles"] = perso + [r for r in modele["libelles"] if r["regex"] not in regex_perso]
    journal.append(f"règles de libellés : {len(modele['libelles'])} règles du modèle"
                   + (f" + {len(perso)} règle(s) personnelle(s) conservée(s) en tête" if perso else ""))

    # Types
    anciens = [renomme(t) for t in data.get("types", [])]
    ajoutes = [t for t in anciens if t not in modele["types"] and t not in ref_v1["types"]]
    types = [t for t in modele["types"] if t != "À classer"] + ajoutes + ["À classer"]
    nouveaux = [t for t in types if t not in anciens]
    data["types"] = types
    data["types_hors_analyse"] = list(dict.fromkeys(
        modele["types_hors_analyse"] + [renomme(t) for t in data.get("types_hors_analyse", [])]))
    if nouveaux:
        journal.append("types ajoutés : " + ", ".join(nouveaux))

    # Catégories banque : on remplace seulement les valeurs que l'utilisateur n'a pas modifiées
    cats = data.setdefault("categories_banque", {})
    for k, v in modele["categories_banque"].items():
        if k not in cats or cats[k] == ref_v1["categories_banque"].get(k):
            cats[k] = v
        else:
            cats[k] = renomme(cats[k])

    # Moyens de paiement : ajouts du modèle manquants
    moyens = data.setdefault("moyens_paiement", [])
    regex_moyens = {m["regex"] for m in moyens}
    for m in modele["moyens_paiement"]:
        if m["regex"] not in regex_moyens:
            moyens.insert(0, m)
            journal.append(f"moyen de paiement ajouté : {m['regex']} -> {m['moyen']}")

    # Contextes : conservés ; l'exemple d'origine non modifié est aligné sur le modèle
    exemples_v1 = {c["nom"]: c for c in ref_v1.get("contextes_exemples", [])}
    exemples_v2 = {c["nom"]: c for c in modele.get("contextes", [])}
    for ctx in data.get("contextes", []):
        if ctx.get("nom") in exemples_v1 and ctx == exemples_v1[ctx["nom"]] and ctx["nom"] in exemples_v2:
            ctx.update(exemples_v2[ctx["nom"]])
        ctx["remplacer_types"] = {renomme(k): renomme(v) for k, v in ctx.get("remplacer_types", {}).items()}

    data["version_config"] = VERSION_CONFIG
    chemin.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    journal.append(f"configuration passée en version {VERSION_CONFIG}")
    return journal


def main() -> None:
    chemin = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("donnees") / "config_comptes.json"
    if not chemin.exists():
        print(f"{chemin} introuvable : rien à migrer.")
        return
    journal = migrer_config(chemin)
    if not journal:
        print(f"{chemin.name} est déjà à jour (version {VERSION_CONFIG}).")
    for ligne in journal:
        print(f"  - {ligne}")


if __name__ == "__main__":
    main()
