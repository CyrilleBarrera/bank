#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Courbes interactives (Plotly) des dépenses dans le temps, par catégorie.

Utilisation :
    comptes-graphiques                                     # depuis donnees\Comptes.xlsx
    comptes-graphiques export.csv                          # ou directement depuis un/des CSV
    comptes-graphiques --periode semaine                   # mois | semaine | auto (défaut)
    comptes-graphiques --hors-ligne                        # HTML autonome (sans Internet, ~4 Mo)
    comptes-graphiques --compte Joint --ouvrir             # un seul compte, ouvre le navigateur

Produit UN fichier HTML avec :
  1. Évolution d'une catégorie (menu déroulant), moyenne en pointillés, périodes de
     contexte (vacances...) surlignées ;
  2. Toutes les catégories en petits graphiques (même échelle de temps) ;
  3. Les principales catégories superposées (+ « Autres ») ;
  4. Le solde de chaque compte et le total général ;
  5. Le coût par jour selon le contexte (Quotidien vs vacances...).

Dépendances : pip install plotly pandas openpyxl
Les fonctions (serie_par_periode, fig_categorie, ...) sont réutilisables en Python.
"""
from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from .releve import (Config, Operation, _lire_existant, analyser_contextes,
                     calculer_totaux, charger_config, console_utf8, etendre_motifs, lire_csv)

# Palette catégorielle validée (daltonisme : ordre fixe, jamais recyclé) + gris pour « Autres »
COULEURS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]
GRIS = "#8a8984"
SURFACE, TEXTE, TEXTE2, GRILLE = "#fcfcfb", "#0b0b0b", "#52514e", "#e7e6e2"
BANDE_CTX = "rgba(138,137,132,0.13)"
MOIS_COURTS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août",
               "sept.", "oct.", "nov.", "déc."]


# --------------------------------------------------------------------------- #
#  Données
# --------------------------------------------------------------------------- #
def charger_operations(sources: list[str], cfg: Config, compte: Optional[str] = None):
    """Charge depuis un classeur .xlsx (produit par la commande comptes) ou des CSV."""
    ops: list[Operation] = []
    soldes = None
    for s in sources:
        if s.lower().endswith((".xlsx", ".xlsm")):
            o, soldes = _lire_existant(Path(s), cfg)
            ops += o
        else:
            ops += lire_csv(s, cfg, compte=compte)
    ops.sort(key=lambda o: o.date)
    calculer_totaux(ops, cfg, soldes_initiaux=soldes)
    return ops


def en_dataframe(ops: list[Operation]) -> pd.DataFrame:
    return pd.DataFrame([{
        "date": pd.Timestamp(o.date), "compte": o.compte, "type": o.type, "contexte": o.contexte,
        "libelle": o.libelle, "depense": o.debit - o.credit, "debit": o.debit, "credit": o.credit,
        **{f"total_{k}": v for k, v in o.totaux.items()},
    } for o in ops])


def choisir_periode(df: pd.DataFrame, periode: str) -> str:
    if periode != "auto":
        return periode
    duree = (df["date"].max() - df["date"].min()).days
    return "semaine" if duree < 183 else "mois"     # < 6 mois : trop peu de points par mois


def serie_par_periode(df: pd.DataFrame, periode: str = "mois") -> pd.DataFrame:
    """Tableau (périodes × types) des dépenses nettes, périodes vides = 0."""
    freq = "MS" if periode == "mois" else "W-MON"
    cle = (df["date"].dt.to_period("M").dt.start_time if periode == "mois"
           else df["date"].dt.to_period("W-SUN").dt.start_time)
    t = df.assign(periode=cle).pivot_table(index="periode", columns="type", values="depense",
                                           aggfunc="sum", fill_value=0.0)
    plage = pd.date_range(t.index.min(), t.index.max(), freq=freq)
    return t.reindex(plage, fill_value=0.0).round(2)


def libelle_periode(ts: pd.Timestamp, periode: str) -> str:
    if periode == "mois":
        return f"{MOIS_COURTS[ts.month - 1]} {ts.year}"
    fin = ts + timedelta(days=6)
    return f"sem. du {ts.day} {MOIS_COURTS[ts.month - 1]} au {fin.day} {MOIS_COURTS[fin.month - 1]}"


def euros(x: float) -> str:
    return f"{x:,.2f} €".replace(",", " ").replace(".", ",")


# --------------------------------------------------------------------------- #
#  Mise en forme commune
# --------------------------------------------------------------------------- #
def _layout(fig: go.Figure, titre: str, hauteur: int = 420, legende: bool = False) -> go.Figure:
    fig.update_layout(
        title=dict(text=titre, x=0, xanchor="left", font=dict(size=17, color=TEXTE)),
        height=hauteur, paper_bgcolor=SURFACE, plot_bgcolor=SURFACE,
        font=dict(family="Arial, Helvetica, sans-serif", size=13, color=TEXTE2),
        margin=dict(l=70, r=30, t=70, b=50), separators=", ",
        hovermode="x unified", showlegend=legende,
        hoverlabel=dict(bgcolor="white", font_size=13, bordercolor=GRILLE),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left", title=None),
    )
    fig.update_xaxes(showgrid=False, linecolor=GRILLE, ticks="outside", tickcolor=GRILLE)
    fig.update_yaxes(gridcolor=GRILLE, zeroline=True, zerolinecolor="#c9c8c3",
                     ticksuffix=" €", tickformat=",.0f")
    return fig


def _axe_fr(fig: go.Figure, index: pd.DatetimeIndex, periode: str, max_ticks: int = 9, **kw) -> None:
    """Graduations en français (Plotly afficherait 'Jun', 'Aug'...)."""
    pas = max(1, -(-len(index) // max_ticks))
    vals = list(index[::pas])
    if periode == "mois":
        txt = [f"{MOIS_COURTS[t.month - 1]} {t.year}" for t in vals]
    else:
        txt = [f"{t.day} {MOIS_COURTS[t.month - 1]}" + (f"<br>{t.year}" if i == 0 or t.month == 1 else "")
               for i, t in enumerate(vals)]
    fig.update_xaxes(tickvals=vals, ticktext=txt, **kw)


def _etiquettes_sans_chevauchement(valeurs: list[float], ecart: float) -> list[float]:
    """Écarte verticalement les étiquettes de fin de courbe trop proches."""
    ordre = sorted(range(len(valeurs)), key=lambda i: valeurs[i])
    pos = list(valeurs)
    for a, b in zip(ordre, ordre[1:]):
        if pos[b] - pos[a] < ecart:
            pos[b] = pos[a] + ecart
    return pos


def _bandes_contextes(fig: go.Figure, cfg: Config, df: pd.DataFrame, **kw) -> None:
    """Surligne les périodes de contexte (vacances...) visibles dans les données."""
    d0, d1 = df["date"].min().date(), df["date"].max().date()
    for c in cfg.contextes:
        if c.actif and c.fin >= d0 and c.debut <= d1:
            fig.add_vrect(x0=pd.Timestamp(c.debut), x1=pd.Timestamp(c.fin) + timedelta(days=1),
                          fillcolor=BANDE_CTX, line_width=0, layer="below",
                          annotation_text=c.nom, annotation_position="top left",
                          annotation_font=dict(size=11, color=TEXTE2), **kw)


# --------------------------------------------------------------------------- #
#  Figures
# --------------------------------------------------------------------------- #
def fig_categorie(df: pd.DataFrame, cfg: Config, periode: str, defaut: str = "Nourriture") -> go.Figure:
    """Une courbe par catégorie, une seule visible à la fois (menu déroulant)."""
    t = serie_par_periode(df, periode)
    exclus = set(cfg.data.get("types_hors_analyse", []))
    types = [c for c in t.columns if t[c].abs().sum() > 0]
    courants = [c for c in types if c not in exclus]
    t["Toutes dépenses courantes"] = t[courants].sum(axis=1)
    types = ["Toutes dépenses courantes"] + types
    if defaut not in types:
        defaut = types[0]
    x_txt = [libelle_periode(ts, periode) for ts in t.index]
    fig = go.Figure()
    for typ in types:
        y = t[typ]
        visible = typ == defaut
        fig.add_trace(go.Scatter(
            x=t.index, y=y, name=typ, mode="lines+markers", visible=visible,
            line=dict(color=COULEURS[0], width=2.5, shape="linear"),
            marker=dict(size=8, color=COULEURS[0], line=dict(color=SURFACE, width=2)),
            customdata=x_txt, hovertemplate="%{customdata}<br><b>%{y:,.2f} €</b><extra>" + typ + "</extra>"))
        fig.add_trace(go.Scatter(
            x=[t.index.min(), t.index.max()], y=[y.mean()] * 2, mode="lines", visible=visible,
            name=f"Moyenne {typ}", line=dict(color=GRIS, width=1.5, dash="dash"), hoverinfo="skip"))
    boutons = []
    for i, typ in enumerate(types):
        vis = [False] * (2 * len(types))
        vis[2 * i] = vis[2 * i + 1] = True
        moy = t[typ].mean()
        boutons.append(dict(label=typ, method="update",
                            args=[{"visible": vis},
                                  {"title.text": f"{typ} — dépense nette par {periode} "
                                                 f"(moyenne {euros(moy)})"}]))
    fig.update_layout(updatemenus=[dict(buttons=boutons, direction="down", x=1, xanchor="right",
                                        y=1.18, yanchor="top", active=types.index(defaut),
                                        bgcolor="white", bordercolor=GRILLE)])
    _bandes_contextes(fig, cfg, df)
    _axe_fr(fig, t.index, periode)
    return _layout(fig, f"{defaut} — dépense nette par {periode} (moyenne {euros(t[defaut].mean())})", 460)


def fig_petits_multiples(df: pd.DataFrame, cfg: Config, periode: str) -> go.Figure:
    """Un petit graphique par catégorie : même axe du temps, échelle verticale propre."""
    t = serie_par_periode(df, periode)
    exclus = set(cfg.data.get("types_hors_analyse", []))
    types = [c for c in t.columns if c not in exclus and t[c].abs().sum() > 0]
    types.sort(key=lambda c: -t[c].sum())
    ncol = 3
    nlig = max(1, -(-len(types) // ncol))
    fig = make_subplots(rows=nlig, cols=ncol, shared_xaxes=True, vertical_spacing=0.09,
                        horizontal_spacing=0.07,
                        subplot_titles=[f"{c} · total {euros(t[c].sum())}" for c in types])
    x_txt = [libelle_periode(ts, periode) for ts in t.index]
    for i, typ in enumerate(types):
        r, c = i // ncol + 1, i % ncol + 1
        fig.add_trace(go.Scatter(
            x=t.index, y=t[typ], mode="lines+markers", name=typ, showlegend=False,
            line=dict(color=COULEURS[0], width=2), marker=dict(size=6, color=COULEURS[0]),
            fill="tozeroy", fillcolor="rgba(42,120,214,0.08)",
            customdata=x_txt, hovertemplate="%{customdata}<br><b>%{y:,.2f} €</b><extra>" + typ + "</extra>"),
            row=r, col=c)
    for c in cfg.contextes:
        if c.actif:
            fig.add_vrect(x0=pd.Timestamp(c.debut), x1=pd.Timestamp(c.fin) + timedelta(days=1),
                          fillcolor=BANDE_CTX, line_width=0, layer="below", row="all", col="all")
    fig.update_annotations(font=dict(size=12, color=TEXTE))
    _layout(fig, "Toutes les catégories (hors virements et charges fixes) — bandes grises = contextes",
            hauteur=230 * nlig + 90)
    fig.update_layout(hovermode="closest")
    fig.update_yaxes(tickfont=dict(size=10))
    _axe_fr(fig, t.index, periode, max_ticks=4, tickfont=dict(size=10))
    return fig


def fig_principales(df: pd.DataFrame, cfg: Config, periode: str, n: int = 4) -> go.Figure:
    """Les n catégories les plus coûteuses superposées + « Autres »."""
    t = serie_par_periode(df, periode)
    exclus = set(cfg.data.get("types_hors_analyse", []))
    types = [c for c in t.columns if c not in exclus]
    top = sorted(types, key=lambda c: -t[c].sum())[:n]
    autres = [c for c in types if c not in top]
    series = [(c, t[c], COULEURS[i]) for i, c in enumerate(top)]
    if autres:
        series.append(("Autres", t[autres].sum(axis=1), GRIS))
    x_txt = [libelle_periode(ts, periode) for ts in t.index]
    fig = go.Figure()
    for nom, y, coul in series:
        fig.add_trace(go.Scatter(
            x=t.index, y=y, name=nom, mode="lines+markers",
            line=dict(color=coul, width=2.5), marker=dict(size=8, color=coul, line=dict(color=SURFACE, width=2)),
            customdata=x_txt, hovertemplate="%{y:,.2f} €"))
    # étiquettes directes en bout de courbe (encre neutre), écartées si elles se chevauchent
    fins = [float(y.iloc[-1]) for _, y, _ in series]
    etendue = max(float(max(y.max() for _, y, _ in series) - min(y.min() for _, y, _ in series)), 1.0)
    for (nom, y, _), pos in zip(series, _etiquettes_sans_chevauchement(fins, etendue / 16)):
        fig.add_annotation(x=t.index[-1], y=pos, text=nom, showarrow=False, xanchor="left",
                           xshift=10, font=dict(size=12, color=TEXTE))
    _layout(fig, f"Principales catégories par {periode}", 460, legende=True)
    _axe_fr(fig, t.index, periode)
    fig.update_layout(margin=dict(r=130))
    _bandes_contextes(fig, cfg, df)
    return fig


def fig_soldes(df: pd.DataFrame, cfg: Config) -> go.Figure:
    """Solde de chaque compte (actif) et total général, après chaque jour."""
    jour = df.groupby("date").last()
    comptes = [c for c in cfg.noms_comptes if (df["compte"] == c).any()]
    fig = go.Figure()
    if len(comptes) == 1:              # un seul compte : sa courbe EST le total
        c = comptes[0]
        fig.add_trace(go.Scatter(x=jour.index, y=jour[f"total_{c}"], name=c, mode="lines",
                                 line=dict(color=COULEURS[0], width=2.5, shape="hv"),
                                 hovertemplate="%{y:,.2f} €"))
        _layout(fig, f"Solde du compte {c}", 400)
        fig.update_xaxes(tickformat="%d/%m/%y")
        _bandes_contextes(fig, cfg, df)
        return fig
    for i, c in enumerate(comptes):
        fig.add_trace(go.Scatter(x=jour.index, y=jour[f"total_{c}"], name=c, mode="lines",
                                 line=dict(color=COULEURS[(i + 1) % 4], width=2, shape="hv"),
                                 hovertemplate="%{y:,.2f} €"))
    if "total_Total" in jour:
        fig.add_trace(go.Scatter(x=jour.index, y=jour["total_Total"], name="Total général", mode="lines",
                                 line=dict(color=COULEURS[0], width=3, shape="hv"),
                                 hovertemplate="%{y:,.2f} €"))
    _layout(fig, "Solde des comptes", 400, legende=True)
    fig.update_xaxes(tickformat="%d/%m/%y")
    _bandes_contextes(fig, cfg, df)
    return fig


def fig_contextes(ops: list[Operation], cfg: Config, types: tuple = ("Nourriture", "Loisirs", None)) -> go.Figure:
    """Coût par jour selon le contexte, pour quelques types (None = toutes dépenses courantes)."""
    fig = go.Figure()
    for i, typ in enumerate(types):
        lignes = analyser_contextes(ops, cfg, typ)
        nom = typ or "Toutes dépenses courantes"
        fig.add_trace(go.Bar(
            y=[l["contexte"] for l in lignes], x=[l["par_jour"] or 0 for l in lignes], name=nom,
            orientation="h", marker=dict(color=COULEURS[i], cornerradius=4, line=dict(color=SURFACE, width=2)),
            customdata=[[l["jours"], l["depense"]] for l in lignes],
            hovertemplate="<b>%{x:,.2f} € / jour</b><br>%{customdata[1]:,.2f} € sur %{customdata[0]} jours"
                          "<extra>" + nom + "</extra>",
            text=[euros(l["par_jour"] or 0) for l in lignes], textposition="outside",
            textfont=dict(color=TEXTE, size=11)))
    _layout(fig, "Coût par jour selon le contexte", 140 + 90 * (len(cfg.contextes) + 1), legende=True)
    fig.update_layout(barmode="group", hovermode="closest", bargap=0.3)
    fig.update_xaxes(ticksuffix=" €", tickformat=",.0f", gridcolor=GRILLE, showgrid=True)
    fig.update_yaxes(ticksuffix="", autorange="reversed", gridcolor=SURFACE)
    return fig


# --------------------------------------------------------------------------- #
#  Page HTML
# --------------------------------------------------------------------------- #
def generer_html(ops: list[Operation], cfg: Config, sortie: str | Path = "graphiques_comptes.html",
                 periode: str = "auto", hors_ligne: bool = False) -> Path:
    df = en_dataframe(ops)
    if df.empty:
        raise ValueError("Aucune opération à représenter")
    periode = choisir_periode(df, periode)
    figures = [
        ("Évolution d'une catégorie", "Choisissez la catégorie dans le menu en haut à droite. "
         "Pointillés = moyenne de la période. Bandes grises = contextes (vacances...).",
         fig_categorie(df, cfg, periode)),
        ("Principales catégories", "Cliquez sur un nom de la légende pour masquer / afficher une courbe ; "
         "double-cliquez pour l'isoler.", fig_principales(df, cfg, periode)),
        ("Toutes les catégories", "Chaque petit graphique a sa propre échelle verticale : "
         "comparez les formes, pas les hauteurs.", fig_petits_multiples(df, cfg, periode)),
        ("Coût selon le contexte", "Dépense nette ramenée au nombre de jours de chaque contexte. "
         "« Quotidien » = jours hors de tout contexte.", fig_contextes(ops, cfg)),
        ("Soldes", "Solde après chaque jour d'opérations.", fig_soldes(df, cfg)),
    ]
    d0, d1 = df["date"].min(), df["date"].max()
    morceaux = []
    for i, (titre, aide, fig) in enumerate(figures):
        html = fig.to_html(full_html=False, include_plotlyjs=(True if hors_ligne else "cdn") if i == 0 else False,
                           config={"displaylogo": False, "responsive": True,
                                   "toImageButtonOptions": {"format": "png", "scale": 2}})
        morceaux.append(f'<section><h2>{titre}</h2><p class="aide">{aide}</p>{html}</section>')
    page = f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dépenses dans le temps</title>
<style>
  body {{ background:{SURFACE}; color:{TEXTE}; font-family:Arial, Helvetica, sans-serif;
         margin:0 auto; max-width:1180px; padding:24px 16px 60px; }}
  h1 {{ font-size:24px; margin:0 0 4px; }}
  h2 {{ font-size:18px; margin:36px 0 2px; }}
  .sous-titre, .aide {{ color:{TEXTE2}; font-size:14px; margin:0 0 8px; }}
  section {{ border-top:1px solid {GRILLE}; }}
</style></head><body>
<h1>Dépenses dans le temps</h1>
<p class="sous-titre">{len(ops)} opérations du {d0.day}/{d0.month:02d}/{d0.year} au {d1.day}/{d1.month:02d}/{d1.year}
 · regroupement par {periode} · montants nets (débits − crédits) · généré le {date.today().strftime('%d/%m/%Y')}</p>
{''.join(morceaux)}
</body></html>"""
    sortie = Path(sortie)
    sortie.write_text(page, encoding="utf-8")
    print(f"Graphiques -> {sortie}  (regroupement par {periode})")
    return sortie


def main() -> None:
    p = argparse.ArgumentParser(description="Courbes interactives des dépenses (Plotly)")
    p.add_argument("sources", nargs="*", help="classeur .xlsx et/ou CSV (défaut : Comptes.xlsx à côté de la config)")
    p.add_argument("-o", "--sortie", help="fichier HTML (défaut : graphiques_comptes.html à côté de la config)")
    p.add_argument("-c", "--config", help="fichier de configuration (défaut : donnees\\config_comptes.json)")
    p.add_argument("--periode", choices=["auto", "mois", "semaine"], default="auto")
    p.add_argument("--compte", help="ne garder qu'un compte (Cyrille, Marion, Joint)")
    p.add_argument("--hors-ligne", action="store_true", help="intègre plotly.js (fonctionne sans Internet)")
    p.add_argument("--ouvrir", action="store_true", help="ouvre le résultat dans le navigateur")
    a = p.parse_args()
    console_utf8()
    cfg = charger_config(a.config)
    dossier = cfg.chemin.parent
    sources = etendre_motifs(a.sources) if a.sources else [str(dossier / "Comptes.xlsx")]
    if not a.sources and not Path(sources[0]).exists():
        raise SystemExit(f"{sources[0]} introuvable : lancez d'abord  comptes <vos CSV>")
    ops = charger_operations(sources, cfg)
    if a.compte:
        ops = [o for o in ops if o.compte == a.compte]
    sortie = generer_html(ops, cfg, a.sortie or dossier / "graphiques_comptes.html", a.periode, a.hors_ligne)
    if a.ouvrir:
        import webbrowser
        webbrowser.open(Path(sortie).resolve().as_uri())


if __name__ == "__main__":
    main()
