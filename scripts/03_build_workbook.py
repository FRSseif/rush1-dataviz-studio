"""Phase 7 — Construction du classeur client (et uniquement du classeur client).

Le fichier est construit DEPUIS ZÉRO à partir des données déjà filtrées : seules
les colonnes autorisées (voir COLONNES_CLIENT) entrent dans le classeur. Les
chiffres affichés sont des formules Excel vivantes sur le tableau « Donnees » ;
leur valeur, calculée ici en Python, est enregistrée comme résultat en cache
(aperçus, mode protégé) et recoupée avec le calcul d'Excel par 05_verify_numbers.py.

Entrées : work/videos_clean.csv, work/resultats.json (privés)
Sortie  : deliverable/Dataviz_Studio_TikTok_Client.xlsx

Usage : .venv/bin/python scripts/03_build_workbook.py
"""
import json
import math
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from scipy import stats

import ooxml_extras as ox

ROOT = Path(__file__).resolve().parent.parent
WORK = ROOT / "work"
SORTIE = ROOT / "deliverable" / "Dataviz_Studio_TikTok_Client.xlsx"

T = "Donnees"
ECH = ["Tendances", "Likées par @tiktok", "Humour"]
ECH_COURT = {"Tendances": "Tendances", "Likées par @tiktok": "Likées", "Humour": "Humour"}
THEMES = ["Animaux", "Famille & enfants", "Humour & sketchs",
          "Lifestyle, musique & sport", "Sujet non déclaré"]
TRANCHES = ["1. < 1 M", "2. 1-3 M", "3. 3-8 M", "4. ≥ 8 M"]
CLASSES = ["1. ≤ 10 s", "2. 11-15 s", "3. 16-30 s", "4. 31-60 s", "5. > 60 s"]
JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
CRENEAUX = ["0-5 h", "6-11 h", "12-17 h", "18-23 h"]

# Charte : une couleur par sujet (reprise par le thème du classeur → TreeMap), bleu nuit.
NUIT, GRIS, GRIS_CLAIR, ORANGE = "#1F3A5F", "#6B7280", "#F3F4F6", "#E76F51"
COULEURS_THEMES = {"Animaux": "2A9D8F", "Famille & enfants": "E9A23B",
                   "Humour & sketchs": "E76F51", "Lifestyle, musique & sport": "5B6BBF",
                   "Sujet non déclaré": "9AA0A6"}
COULEURS_ECH = {"Tendances": NUIT, "Likées par @tiktok": "#9AA0A6", "Humour": "#E76F51"}

FMT_M = '#,##0.0,," M"'
FMT_PTS = '+0.0" pts";-0.0" pts";0.0" pt"'
FMT_RHO = "+0.00;-0.00;0.00"
FMT_P = '[<0.001]"< 0,001";0.000'

# Colonnes du tableau client, dans l'ordre. f = colonne calculée par formule.
COLONNES_CLIENT = [
    ("Échantillon", None), ("Créateur", None), ("Thème", None), ("Âge (jours)", None),
    ("Jour", None), ("Week-end", "f"), ("Heure UTC", None), ("Créneau UTC", "f"),
    ("Durée (s)", None), ("Classe de durée", "f"), ("Plus de 15 s", "f"),
    ("Nb hashtags", None), ("Hashtags génériques", None), ("Longueur du texte", None),
    ("Légende vide", None), ("Mention duo stitch", None), ("Question ou appel", None),
    ("Son original", None), ("Stitch autorisé", None), ("Abonnés", None),
    ("Tranche abonnés", "f"), ("Vidéos publiées", None), ("Vues", None), ("Likes", None),
    ("Commentaires", None), ("Partages", None), ("Engagement", "f"), ("Top performer", "f"),
    ("Vues par abonné", "f"), ("Compté dans les cumuls", None),
    ("Rang engagement", "f"), ("Rang vues", "f"), ("Rang durée", "f"),
    ("Rang hashtags", "f"), ("Rang longueur", "f"), ("Rang abonnés", "f"),
    ("Rang vidéos publiées", "f"),
]
COLS = [c for c, _ in COLONNES_CLIENT]
RANGS = {"Rang engagement": "Engagement", "Rang vues": "Vues", "Rang durée": "Durée (s)",
         "Rang hashtags": "Nb hashtags", "Rang longueur": "Longueur du texte",
         "Rang abonnés": "Abonnés", "Rang vidéos publiées": "Vidéos publiées"}

# Plage des seuils de top performers (onglet Méthode & limites), utilisée par la colonne
# « Top performer » et par l'indice d'engagement des leaders.
SEUILS_LIGNE0 = 8          # ligne Excel (0-based) de « Tendances » dans le tableau des seuils
SEUILS_PLAGE = "'Méthode & limites'!$B$9:$D$11"


def fr(x, dec=1, signe=False):
    s = f"{x:+,.{dec}f}" if signe else f"{x:,.{dec}f}"
    return s.replace(",", " ").replace(".", ",").replace("-", "−")


def pct(x, dec=1):
    return f"{fr(x * 100, dec)} %"


def pval(p):
    return "< 0,001" if p < 0.001 else fr(p, 2 if p >= 0.01 else 3)


def q(s):
    return '"' + str(s).replace('"', '""') + '"'


def ref(col):
    return f"{T}[{col}]"


def this(col):
    return f"{T}[[#This Row],[{col}]]"


# --------------------------------------------------------------------------- #
# Données client
# --------------------------------------------------------------------------- #
def oui_non(s):
    return np.where(s.astype(bool), "Oui", "Non")


def donnees_client(df, res):
    nommes, pseudo = set(res["nommes"]), res["pseudonymes"]
    d = pd.DataFrame()
    d["Échantillon"] = df["echantillon"]
    d["Créateur"] = df["createur"].map(lambda c: "@" + c if c in nommes else pseudo[c])
    d["Thème"] = df["theme"]
    d["Âge (jours)"] = df["age_jours"].astype(int)
    d["Jour"] = df["jour_nom"]
    d["Week-end"] = np.where(d["Jour"].isin(["samedi", "dimanche"]), "Oui", "Non")
    d["Heure UTC"] = df["heure_utc"].astype(int)
    h = d["Heure UTC"]
    d["Créneau UTC"] = np.select([h < 6, h < 12, h < 18], CRENEAUX[:3], CRENEAUX[3])
    d["Durée (s)"] = df["duree_s"].astype(int)
    s = d["Durée (s)"]
    d["Classe de durée"] = np.select([s <= 10, s <= 15, s <= 30, s <= 60], CLASSES[:4],
                                     CLASSES[4])
    d["Plus de 15 s"] = np.where(s > 15, "Oui", "Non")
    d["Nb hashtags"] = df["nb_hashtags"].astype(int)
    d["Hashtags génériques"] = oui_non(df["hashtag_generique"])
    d["Longueur du texte"] = df["longueur_texte"].astype(int)
    d["Légende vide"] = oui_non(df["legende_vide"])
    d["Mention duo stitch"] = oui_non(df["format_collab"])
    d["Question ou appel"] = oui_non(df["question_ou_appel"])
    humour = df["echantillon"] == "Humour"
    d["Son original"] = np.where(humour, np.where(df["son_original"] == True, "Oui", "Non"), "")
    d["Stitch autorisé"] = np.where(humour, np.where(df["stitch_autorise"] == True, "Oui", "Non"),
                                    "")
    d["Abonnés"] = df["abonnes"]
    a = d["Abonnés"]
    d["Tranche abonnés"] = np.select(
        [a.isna(), a < 1e6, a < 3e6, a < 8e6], [""] + TRANCHES[:3], TRANCHES[3])
    d["Vidéos publiées"] = df["nb_videos_compte"]
    for c, src in (("Vues", "vues"), ("Likes", "likes"), ("Commentaires", "commentaires"),
                   ("Partages", "partages")):
        d[c] = df[src].astype("int64")
    d["Engagement"] = (d["Likes"] + d["Commentaires"] + d["Partages"]) / d["Vues"]
    seuil = d.groupby("Échantillon")["Engagement"].transform(lambda x: x.quantile(0.75))
    d["Top performer"] = (d["Engagement"] >= seuil).astype(int)
    d["Vues par abonné"] = d["Vues"] / d["Abonnés"]
    # Une même vidéo figure dans Tendances et Humour : comptée une seule fois dans les cumuls.
    doublon = df.duplicated("video_id", keep="first") & df["echantillon"].isin(
        ["Tendances", "Humour"]) & df["video_id"].isin(
        df.loc[df["echantillon"] == "Tendances", "video_id"])
    d["Compté dans les cumuls"] = np.where(doublon, "Non", "Oui")
    for rang, col in RANGS.items():
        d[rang] = d.groupby("Échantillon")[col].rank(method="average")
    ordre = {e: i for i, e in enumerate(ECH)}
    d = d.sort_values(["Échantillon", "Vues"], key=lambda c: c.map(ordre) if c.name ==
                      "Échantillon" else -c).reset_index(drop=True)
    assert list(d.columns) == COLS
    return d


FORMULES = {
    "Week-end": lambda: f'=IF(OR({this("Jour")}="samedi",{this("Jour")}="dimanche"),"Oui","Non")',
    "Créneau UTC": lambda: (f'=IF({this("Heure UTC")}<6,"0-5 h",IF({this("Heure UTC")}<12,'
                            f'"6-11 h",IF({this("Heure UTC")}<18,"12-17 h","18-23 h")))'),
    "Classe de durée": lambda: (
        f'=IF({this("Durée (s)")}<=10,"1. ≤ 10 s",IF({this("Durée (s)")}<=15,"2. 11-15 s",'
        f'IF({this("Durée (s)")}<=30,"3. 16-30 s",IF({this("Durée (s)")}<=60,"4. 31-60 s",'
        f'"5. > 60 s"))))'),
    "Plus de 15 s": lambda: f'=IF({this("Durée (s)")}>15,"Oui","Non")',
    "Tranche abonnés": lambda: (
        f'=IF({this("Abonnés")}="","",IF({this("Abonnés")}<1000000,"1. < 1 M",'
        f'IF({this("Abonnés")}<3000000,"2. 1-3 M",IF({this("Abonnés")}<8000000,"3. 3-8 M",'
        f'"4. ≥ 8 M"))))'),
    "Engagement": lambda: (f'=({this("Likes")}+{this("Commentaires")}+{this("Partages")})'
                           f'/{this("Vues")}'),
    "Top performer": lambda: (f'=IF({this("Engagement")}>=VLOOKUP({this("Échantillon")},'
                              f'{SEUILS_PLAGE},2,FALSE),1,0)'),
    "Vues par abonné": lambda: f'=IF({this("Abonnés")}="","",{this("Vues")}/{this("Abonnés")})',
}


def formule_rang(col):
    base = (f'SUMPRODUCT(({ref("Échantillon")}={this("Échantillon")})*({ref(col)}<{this(col)}))'
            f'+(SUMPRODUCT(({ref("Échantillon")}={this("Échantillon")})*({ref(col)}='
            f'{this(col)}))+1)/2')
    if col in ("Abonnés", "Vidéos publiées"):
        return f'=IF({this(col)}="","",{base})'
    return "=" + base


# --------------------------------------------------------------------------- #
# Formules + valeurs en cache
# --------------------------------------------------------------------------- #
class Calc:
    """Construit une formule Excel et calcule la même valeur en Python."""

    def __init__(self, d):
        self.d = d

    def _mask(self, conds):
        m = pd.Series(True, index=self.d.index)
        for c, op, v in conds:
            s = self.d[c]
            m &= {"=": s == v, "<>": s != v, ">": s > v, "<": s < v}[op]
        return m

    @staticmethod
    def _expr(conds):
        return "*".join(f"({ref(c)}{op}{q(v) if isinstance(v, str) else v})"
                        for c, op, v in conds)

    def median(self, conds, target):
        f = f"{{=MEDIAN(IF({self._expr(conds)},{ref(target)}))}}"
        return f, float(self.d.loc[self._mask(conds), target].median())

    def ecart_pts(self, conds, col):
        f1, v1 = self.median(conds + [(col, "=", "Oui")], "Engagement")
        f2, v2 = self.median(conds + [(col, "=", "Non")], "Engagement")
        f = f"{{=(MEDIAN(IF({self._expr(conds + [(col, '=', 'Oui')])},{ref('Engagement')}))" \
            f"-MEDIAN(IF({self._expr(conds + [(col, '=', 'Non')])},{ref('Engagement')})))*100}}"
        return f, (v1 - v2) * 100

    def correl(self, ech, rang1, rang2):
        e = q(ech)
        f = (f"{{=CORREL(IF({ref('Échantillon')}={e},{ref(rang1)}),"
             f"IF({ref('Échantillon')}={e},{ref(rang2)}))}}")
        sub = self.d[self.d["Échantillon"] == ech]
        x, y = pd.to_numeric(sub[rang1], errors="coerce"), pd.to_numeric(sub[rang2],
                                                                         errors="coerce")
        m = x.notna() & y.notna()
        return f, float(np.corrcoef(x[m], y[m])[0, 1]), int(m.sum())

    def countifs(self, conds):
        args = ",".join(f"{ref(c)},{q(('<>' if op == '<>' else '') + v) if isinstance(v, str) else v}"
                        for c, op, v in conds)
        return f"=COUNTIFS({args})", int(self._mask(conds).sum())

    def averageifs(self, target, conds):
        args = ",".join(f"{ref(c)},{q(('<>' if op == '<>' else '') + v)}" for c, op, v in conds)
        return f"=AVERAGEIFS({ref(target)},{args})", float(self.d.loc[self._mask(conds),
                                                                      target].mean())

    def sumifs(self, target, conds):
        args = ",".join(f"{ref(c)},{q(('<>' if op == '<>' else '') + v)}" for c, op, v in conds)
        return f"=SUMIFS({ref(target)},{args})", float(self.d.loc[self._mask(conds),
                                                                  target].sum())


def p_correl(r, n):
    t = abs(r) * math.sqrt((n - 2) / (1 - r * r))
    return float(2 * stats.t.sf(t, n - 2))


# --------------------------------------------------------------------------- #
# Classeur
# --------------------------------------------------------------------------- #
class Classeur:
    def __init__(self, path):
        self.wb = xlsxwriter.Workbook(path, {"use_future_functions": False})
        self.wb.set_properties({
            "title": "Dataviz Studio — ce qui fait performer un contenu TikTok",
            "subject": "Étude des facteurs de performance TikTok (données publiques 2021)",
            "author": "", "manager": "", "company": "", "category": "Version client",
            "keywords": "TikTok, influence, engagement", "comments": "",
        })
        self._f = {}

    def fmt(self, **kw):
        key = tuple(sorted(kw.items()))
        if key not in self._f:
            base = {"font_name": "Calibri", "font_size": 10, "valign": "vcenter"}
            base.update(kw)
            self._f[key] = self.wb.add_format(base)
        return self._f[key]


def styles(C):
    s = {}
    s["titre"] = C.fmt(bold=True, font_size=18, font_color=NUIT)
    s["sous_titre"] = C.fmt(font_size=10, font_color=GRIS, text_wrap=True, valign="top")
    s["section"] = C.fmt(bold=True, font_size=11, font_color="#FFFFFF", bg_color=NUIT)
    s["th"] = C.fmt(bold=True, bg_color="#E5E7EB", text_wrap=True, border=1,
                    border_color="#D1D5DB", align="center")
    s["th_g"] = C.fmt(bold=True, bg_color="#E5E7EB", text_wrap=True, border=1,
                      border_color="#D1D5DB")
    s["txt"] = C.fmt(text_wrap=True, valign="top", border=1, border_color="#E5E7EB")
    s["txt_nb"] = C.fmt(text_wrap=True, valign="top")
    s["txt_b"] = C.fmt(text_wrap=True, valign="top", bold=True, border=1,
                       border_color="#E5E7EB")
    s["puce"] = C.fmt(text_wrap=True, valign="top", indent=1)
    s["note"] = C.fmt(italic=True, font_color=GRIS, text_wrap=True, valign="top", font_size=9)
    s["lien"] = C.fmt(font_color="#1D4ED8", underline=1, valign="top", text_wrap=True)
    for nom, nf in (("int", "#,##0"), ("m", FMT_M), ("pct", "0.0%"), ("pct0", "0%"),
                    ("pts", FMT_PTS), ("rho", FMT_RHO), ("p", FMT_P), ("dec1", "0.0"),
                    ("x", '0.0"×"')):
        s[nom] = C.fmt(num_format=nf, border=1, border_color="#E5E7EB", align="center")
        s[nom + "_big"] = C.fmt(num_format=nf, bold=True, font_size=16, font_color=NUIT,
                                align="center", border=1, border_color="#E5E7EB")
    s["verdict"] = {
        "Effet confirmé": C.fmt(bold=True, bg_color="#D1FAE5", border=1,
                                border_color="#E5E7EB", align="center", text_wrap=True),
        "Effet isolé": C.fmt(bold=True, bg_color="#FEF3C7", border=1,
                             border_color="#E5E7EB", align="center", text_wrap=True),
        "Contradictoire": C.fmt(bold=True, bg_color="#EDE9FE", border=1,
                                border_color="#E5E7EB", align="center", text_wrap=True),
        "Aucun effet mesurable": C.fmt(bold=True, bg_color="#F3F4F6", font_color="#4B5563",
                                       border=1, border_color="#E5E7EB", align="center",
                                       text_wrap=True),
    }
    s["centre"] = C.fmt(align="center", border=1, border_color="#E5E7EB", text_wrap=True)
    return s


def hauteur(*cellules, minimum=16):
    """Hauteur de ligne (points) pour que des textes renvoyés à la ligne restent visibles.

    cellules : (texte, largeur de la zone en caractères, taille de police).
    """
    h = minimum
    for texte, largeur, taille in cellules:
        par_ligne = max(8, largeur * (1.12 if taille <= 10 else 0.98))
        lignes = sum(max(1, math.ceil(len(par) / par_ligne)) for par in str(texte).split("\n"))
        h = max(h, lignes * (taille + 3.2) + 5)
    return h


def entete(ws, s, titre, sous_titre, largeur):
    ws.hide_gridlines(2)
    ws.set_row(0, 30)
    ws.write(0, 1, titre, s["titre"])
    ws.set_row(1, 30)
    ws.merge_range(1, 1, 1, largeur, sous_titre, s["sous_titre"])


def section(ws, s, row, texte, c1, c2):
    ws.set_row(row, 20)
    ws.merge_range(row, c1, row, c2, texte, s["section"])


def lien(ws, s, row, col, feuille, cellule, texte):
    ws.write_url(row, col, f"internal:'{feuille}'!{cellule}", s["lien"], string=texte)


def ecrire_formule(ws, r, c, f, v, fmt):
    if f.startswith("{"):
        ws.write_array_formula(r, c, r, c, f, fmt, v)
    else:
        ws.write_formula(r, c, f, fmt, v)


def py(v):
    """Valeur Python native (XlsxWriter n'écrit pas les entiers numpy comme nombres)."""
    if v is None:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating, float)):
        return None if math.isnan(v) else float(v)
    return v


def cache_val(v):
    v = py(v)
    return "" if v is None else v


# --------------------------------------------------------------------------- #
# Onglet « Données »
# --------------------------------------------------------------------------- #
def feuille_donnees(C, s, ws, d):
    n = len(d)
    fmts = {}
    for c in COLS:
        if c in ("Vues", "Likes", "Commentaires", "Partages", "Abonnés", "Vidéos publiées"):
            fmts[c] = C.fmt(num_format="#,##0")
        elif c == "Engagement":
            fmts[c] = C.fmt(num_format="0.0%")
        elif c == "Vues par abonné" or c.startswith("Rang"):
            fmts[c] = C.fmt(num_format="0.0")
        else:
            fmts[c] = C.fmt()
    entete_fmt = C.fmt(bold=True, text_wrap=True, font_color="#FFFFFF", bg_color=NUIT,
                       valign="vcenter")
    formules = {c: (FORMULES[c]() if c in FORMULES else formule_rang(RANGS[c]))
                for c, kind in COLONNES_CLIENT if kind == "f"}
    columns = []
    for c, kind in COLONNES_CLIENT:
        spec = {"header": c, "format": fmts[c], "header_format": entete_fmt}
        if kind == "f":
            spec["formula"] = formules[c]
        columns.append(spec)
    data = [[py(v) for v in row] for row in d.itertuples(index=False)]
    ws.add_table(0, 0, n, len(COLS) - 1, {"name": T, "style": "Table Style Light 1",
                                          "columns": columns, "data": data})
    for j, (c, kind) in enumerate(COLONNES_CLIENT):
        if kind == "f":
            for i in range(n):
                ws.write_formula(i + 1, j, formules[c], fmts[c], cache_val(d.iat[i, j]))
    ws.set_row(0, 32)
    largeurs = {"Échantillon": 17, "Créateur": 21, "Thème": 23, "Classe de durée": 12,
                "Compté dans les cumuls": 11, "Tranche abonnés": 11}
    for j, c in enumerate(COLS):
        ws.set_column(j, j, largeurs.get(c, 10.5))
    ws.freeze_panes(1, 2)
    ws.set_zoom(90)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0)


# --------------------------------------------------------------------------- #
# Onglet « Méthode & limites » — le tableau des seuils est en B9:D11 (SEUILS_PLAGE)
# --------------------------------------------------------------------------- #
def feuille_methode(C, s, ws, d, res, calc):
    ws.set_column(0, 0, 1.5)
    ws.set_column(1, 1, 30)
    ws.set_column(2, 6, 15)
    ws.set_column(7, 7, 55)
    entete(ws, s, "Méthode, sources et limites",
           "Comment les chiffres de ce classeur ont été obtenus, ce qu'ils valent, et comment "
           "les données personnelles ont été traitées. Toutes les valeurs sont recalculées par "
           "des formules à partir de l'onglet Données, sauf mention contraire.", 7)
    e = res["echantillons"]
    t, h = e["Tendances"], e["Humour"]

    def para(row, texte, haut, fmt=None):
        ws.set_row(row, max(haut, hauteur((texte, 160, 10))))
        ws.merge_range(row, 1, row, 7, texte, fmt or s["txt_nb"])

    section(ws, s, 3, "1. Définition du succès", 1, 7)
    para(4, "Question du client : qu'est-ce qui distingue les contenus et les créateurs TikTok "
            "les plus performants ? Mesure retenue : le TAUX D'ENGAGEMENT = (likes + "
            "commentaires + partages) ÷ vues, c'est-à-dire la part des personnes touchées qui "
            "réagissent activement.", 32)
    para(5, "Pourquoi : (1) c'est la réaction du public qu'achète une marque, pas seulement "
            "l'affichage ; (2) c'est un ratio : il ne favorise pas mécaniquement les gros comptes "
            "ni les vidéos en ligne depuis plus longtemps (dans les tendances, l'âge d'une vidéo "
            f"est lié à ses vues, ρ = {fr(t['age_vues']['rho'], 2, True)}, mais pas à son "
            f"engagement, ρ = {fr(t['age_engagement']['rho'], 2, True)}, non significatif) ; "
            "(3) dans les tendances, l'engagement va de pair avec la portée (ρ vues–engagement "
            f"= {fr(t['vues_engagement']['rho'], 2, True)}). La portée (vues) est toujours "
            "affichée à côté.", 58)
    para(6, "TOP PERFORMERS = le quart supérieur (75e centile) du taux d'engagement, calculé "
            "séparément dans chaque échantillon, car les niveaux d'engagement diffèrent selon la "
            "façon dont l'échantillon a été constitué. Ces seuils alimentent la colonne « Top "
            "performer » de l'onglet Données.", 32)
    hdr = ["Échantillon", "Seuil top performers (75e centile)", "Engagement médian",
           "Vidéos", "Vues médianes", "Créateurs distincts"]
    ws.set_row(7, 30)
    for j, x in enumerate(hdr):
        ws.write(7, 1 + j, x, s["th"] if j else s["th_g"])
    for i, ech in enumerate(ECH):
        r = SEUILS_LIGNE0 + i
        sub = d[d["Échantillon"] == ech]
        ws.write(r, 1, ech, s["txt_b"])
        cond = f"{ref('Échantillon')}={q(ech)}"
        ws.write_array_formula(r, 2, r, 2, f"{{=PERCENTILE(IF({cond},{ref('Engagement')}),0.75)}}",
                               s["pct"], float(sub["Engagement"].quantile(0.75)))
        f, v = calc.median([("Échantillon", "=", ech)], "Engagement")
        ecrire_formule(ws, r, 3, f, v, s["pct"])
        f, v = calc.countifs([("Échantillon", "=", ech)])
        ecrire_formule(ws, r, 4, f, v, s["int"])
        f, v = calc.median([("Échantillon", "=", ech)], "Vues")
        ecrire_formule(ws, r, 5, f, v, s["m"])
        ws.write_array_formula(
            r, 6, r, 6, f"{{=SUM(IF({cond},1/COUNTIFS({ref('Créateur')},{ref('Créateur')},"
                        f"{ref('Échantillon')},{q(ech)})))}}", s["int"],
            int(sub["Créateur"].nunique()))
    para(12, "Limite de cette définition : l'échantillon Humour ne contient que des vidéos à plus "
             f"de {fr(res['gros_comptes']['likes_min'] / 1e6)} M de likes (top 100 du hashtag, "
             "trié par likes). À likes presque constants, l'engagement y baisse mécaniquement "
             "quand les vues montent (ρ vues–engagement = "
             f"{fr(res['gros_comptes']['vues_engagement']['rho'], 2, True)}). Cet échantillon "
             "sert donc surtout à répondre à la question des gros comptes, avec les vues et les "
             "vues par abonné.", 45, s["note"])

    r = 14
    section(ws, s, r, "2. Extraits utilisés et écartés", 1, 7)
    ws.set_row(r + 1, 30)
    for j, x in enumerate(["Extrait", "Contenu", "Publications", "Collecte (estimée)",
                           "Rôle"]):
        ws.write(r + 1, 1 + j, x, s["th"] if j else s["th_g"])
    ws.merge_range(r + 1, 6, r + 1, 7, "Pourquoi", s["th"])
    extraits = [
        ("trending_videos.csv", "100 vidéos, 84 créateurs", "27/05 → 24/08/2021", "≈ 25/08/2021",
         "Principal", "Vidéos mises en avant par l'algorithme, les plus récentes : la meilleure "
         "image de « ce qui performe ». Contient les 4 métriques et les leviers (durée, légende, "
         "date). Ne contient pas les abonnés."),
        ("tiktok_funny_hashtag_videos.csv", "100 vidéos, 79 créateurs", "04/10/2019 → 31/07/2021",
         "≈ 25/08/2021", "Contrôle + gros comptes", "Seul extrait avec les abonnés et la musique. "
         "Biais : un seul registre (#funny) et top 100 trié par likes (≥ 6,8 M)."),
        ("tiktok_collected_liked_videos.csv", "100 vidéos, 92 créateurs", "10/02 → 21/08/2021",
         "≈ 22/08/2021", "Contrôle", "Même format que les tendances, mais sélection éditoriale "
         "du compte officiel (célébrités, marques) : sert à vérifier qu'un effet se répète."),
        ("trending_authors.csv", "100 lignes, 92 auteurs (8 doublons)", "—", "≈ 23/08/2021",
         "Écarté", "Aucune métrique de performance, surtout des données personnelles (bios avec "
         "e-mails, téléphones, comptes de paiement). Seul apport possible : le badge vérifié, "
         "conséquence du succès et non levier. Le garder contredirait la minimisation (RGPD)."),
        ("tiktok_collected_videos.csv", "100 vidéos, 1 compte", "09/09/2019 → 10/08/2021",
         "≥ 10/08/2021", "Écarté", "Vidéos du compte officiel @tiktok : un acteur institutionnel "
         "unique qui promeut la plateforme, pas un créateur qu'une marque peut recruter."),
    ]
    for i, x in enumerate(extraits):
        rr = r + 2 + i
        ws.set_row(rr, hauteur((x[5], 70, 10), (x[1], 15, 10), (x[0], 30, 10)))
        for j in range(5):
            ws.write(rr, 1 + j, x[j], s["txt_b"] if j == 0 else s["txt"])
        ws.merge_range(rr, 6, rr, 7, x[5], s["txt"])

    r = 22
    blocs = [
        ("3. Préparation des données", [
            "Les 3 extraits retenus sont fusionnés en une table de 300 vidéos (onglet Données), "
            "colonnes renommées et uniformisées. Les CSV bruts ne sont jamais modifiés ; tout est "
            "reproductible avec les scripts du projet.",
            "Doublons : aucune vidéo en double dans un même extrait. Une vidéo figure à la fois "
            "dans Tendances et Humour (chiffres quasi identiques) : elle reste dans chaque "
            "échantillon mais n'est comptée qu'une fois dans les vues cumulées (colonne « Compté "
            "dans les cumuls »).",
            "Valeurs manquantes : 19 légendes vides (13 tendances, 6 likées), traitées comme une "
            "variable (« Légende vide ») ; 3 vidéos à 0 commentaire (commentaires sans doute "
            "désactivés), conservées.",
            "Dates : la date de publication est encodée dans l'identifiant de chaque vidéo "
            "(vérifié sur les extraits qui la fournissent). Jour et heure sont en temps universel "
            "(UTC) : l'heure locale des créateurs n'est pas connue.",
            "Les compteurs sont arrondis par la plateforme (2 à 3 chiffres significatifs, par "
            "exemple 8 300 000) : les petits écarts ne sont pas interprétables. Dans l'extrait "
            "Humour, la colonne « video_stats » correspond aux likes (vérifié sur la vidéo "
            "commune : 8,3 M de likes dans les deux extraits).",
            "Variables dérivées : taux d'engagement, classe de durée, nombre de hashtags, "
            "hashtags génériques (#fyp, #foryou, #viral, #xyzbca, #parati…), longueur du texte "
            "hors hashtags et mentions, mention/duo/stitch/réponse, question ou appel à l'action, "
            "week-end, créneau horaire, tranche d'abonnés, vues par abonné.",
            "Sujet (thème) : déduit de mots-clés présents dans les hashtags, la légende et le "
            "pseudo (dog, cat, kids, mom, dance, food, funny, prank…), par ordre de priorité "
            "Animaux > Famille & enfants > Lifestyle, musique & sport > Humour & sketchs. Chaque "
            "créateur reçoit le sujet le plus fréquent de ses vidéos ; sans mot-clé : « Sujet non "
            "déclaré ». Classement automatique et approximatif, relu à la main.",
        ]),
        ("4. Méthodes statistiques", [
            "Médianes plutôt que moyennes : vues et likes sont très asymétriques (quelques vidéos "
            "dépassent 100 M de vues).",
            "Facteur oui/non : écart de taux d'engagement médian (en points) et test de "
            "Mann-Whitney. Facteur continu : corrélation de rang ρ de Spearman (de −1 à +1 ; "
            "|ρ| < 0,3 = lien faible), recalculée dans Excel par CORREL sur les colonnes « Rang … » "
            "de l'onglet Données. Facteur à plusieurs modalités : écart entre la meilleure et la "
            "moins bonne modalité, test de Kruskal-Wallis.",
            "Seuil : p < 0,05. Verdict : « Effet confirmé » si significatif dans au moins deux "
            "échantillons, dans le même sens ; « Effet isolé » si dans un seul ; "
            "« Contradictoire » si les sens s'opposent ; sinon « Aucun effet mesurable ».",
            "Une quarantaine de tests ont été menés : au seuil de 5 %, environ 2 « faux positifs » "
            "sont attendus par hasard. Un effet isolé est donc une piste à tester, pas une preuve.",
            "Les p-values des tests de Mann-Whitney et de Kruskal-Wallis sont calculées par le "
            "script Python du projet (Excel ne propose pas ces tests), ainsi que deux corrélations "
            "signalées dans l'onglet Gros comptes. Tous les autres chiffres sont des formules "
            "Excel, recoupées avec Python.",
        ]),
        ("5. Biais et limites : ce que ces données ne peuvent pas dire", [
            "Biais de sélection : on n'observe que des contenus déjà populaires, choisis par "
            "l'algorithme, par un hashtag trié par likes ou par le compte officiel. Aucun échec "
            "n'est visible : impossible de dire ce qui fait entrer une vidéo dans les tendances, "
            "et les écarts entre « bons » et « très bons » contenus sont mécaniquement réduits.",
            "Date : collectes d'août 2021, publications de 2019 à 2021. L'algorithme, les durées "
            "possibles (jusqu'à 10 minutes depuis 2022), les formats et les audiences ont changé "
            "depuis : les résultats décrivent 2021, ils sont à revalider avant d'être appliqués.",
            "Trois échantillons de 100 vidéos, collectés à des moments et selon des règles "
            "différents : ils ne sont pas représentatifs de TikTok dans son ensemble. On les "
            "compare entre eux sans les additionner (sauf pour les vues cumulées par créateur).",
            "Corrélation n'est pas causalité. Les abonnés sont mesurés en août 2021, après la "
            "publication des vidéos : ils peuvent être la conséquence du succès autant que sa "
            "cause.",
            "Absent des données : l'audience (âge, pays), le temps de visionnage, les conversions "
            "et les ventes, le coût des créateurs, et le contenu réel des vidéos (seule la "
            "légende est disponible).",
        ]),
        ("6. Données personnelles (RGPD) : ce que contient cette version et pourquoi", [
            "Ces contenus sont publics, mais ce sont des données personnelles : le RGPD "
            "s'applique. L'agence les traite sur la base de son intérêt légitime (art. 6.1.f) : "
            "étudier des contenus rendus publics par leurs auteurs pour conseiller une marque, "
            "sans profilage individuel ni décision automatisée.",
            "Minimisation : cette version ne contient que ce qui sert la question. Exclus : "
            "identifiants internes, liens des vidéos, photos de profil, noms affichés, biographies "
            "(e-mails, téléphones, adresses, comptes de paiement, mentions de santé), légendes "
            "(qui citent des tiers), titres et auteurs des musiques, badges et indicateurs de "
            "compte. Ces éléments ne figurent nulle part dans le fichier.",
            f"Pseudonymisation : seuls les {len(res['nommes'])} créateurs recommandés (leaders "
            "par catégorie) sont nommés, par leur pseudo public. Les autres apparaissent sous un "
            "code (« Créateur 001 ») sans table de correspondance dans ce fichier. Un compte dont "
            "le pseudo révèle un état de santé n'est jamais nommé (donnée sensible, art. 9).",
            "Usage par le client : limité au choix de créateurs pour cette campagne ; pas de "
            "réutilisation, de diffusion ni de croisement avec d'autres fichiers ; suppression à "
            "la fin de la campagne, au plus tard 6 mois après réception.",
            "Contact des créateurs : par leurs canaux professionnels publics, en les informant dès "
            "le premier contact de l'origine des données et de leur droit de s'y opposer "
            "(art. 14 et 21). Toute demande d'effacement est transmise à l'agence. Les chiffres "
            "datent de 2021 : à revérifier avant tout contact (exactitude, art. 5).",
        ]),
    ]
    for titre, puces in blocs:
        section(ws, s, r, titre, 1, 7)
        r += 1
        for p in puces:
            ws.set_row(r, hauteur(("• " + p, 160, 10)))
            ws.merge_range(r, 1, r, 7, "• " + p, s["puce"])
            r += 1
        r += 1
    ws.set_zoom(90)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.print_area(0, 0, r, 7)


# --------------------------------------------------------------------------- #
# Onglet « Facteurs »
# --------------------------------------------------------------------------- #
BIN_COLS = {"duree_15": "Plus de 15 s", "hashtag_generique": "Hashtags génériques",
            "legende_vide": "Légende vide", "format_collab": "Mention duo stitch",
            "question_ou_appel": "Question ou appel", "week_end": "Week-end",
            "son_original": "Son original", "stitch_autorise": "Stitch autorisé"}
NUM_COLS = {"duree_s": "Rang durée", "nb_hashtags": "Rang hashtags",
            "longueur_texte": "Rang longueur", "nb_videos_compte": "Rang vidéos publiées"}
ORDRE_FACTEURS = ["duree_15", "duree_s", "hashtag_generique", "nb_hashtags",
                  "longueur_texte", "legende_vide", "format_collab", "question_ou_appel",
                  "week_end", "jour_nom", "creneau_utc", "theme", "son_original",
                  "stitch_autorise", "nb_videos_compte"]
# Positions (0-based) des blocs de l'onglet Facteurs.
FA_TABLE = 5            # première ligne du tableau des facteurs
FA_GRAPH = 22           # section graphique
FA_CONSEQ = 40          # conséquences du succès
FA_DUREE = 47           # détail durée
FA_JOURS = 62           # détail jours / créneaux
FA_TCD = 80             # tableau croisé dynamique (ligne de la zone du TCD)
# Catégories & TreeMap
CA_THEMES = 5
CA_LEADERS = 15
CA_PORTEE = 33
CA_TREEMAP = 42
CA_THEME_ECH = 72
CA_TCD = 98


def phrases_facteurs(res):
    f = {x["cle"]: x["resultats"] for x in res["facteurs"]}
    T, L, H = "Tendances", "Likées par @tiktok", "Humour"
    d15 = f["duree_15"]
    return {
        "duree_15": f"Plus d'engagement dans les tendances ({pct(d15[T]['med_oui'])} contre "
                    f"{pct(d15[T]['med_non'])} ; {pct(d15[T]['top_oui'], 0)} de top performers "
                    f"contre {pct(d15[T]['top_non'], 0)}), mais rien de tel dans les deux autres "
                    "échantillons : une piste à tester, pas une règle.",
        "duree_s": f"Lien faible : nul dans les tendances (ρ = {fr(f['duree_s'][T]['rho'], 2, True)}) "
                   f"et les likées, positif seulement dans l'humour (ρ = "
                   f"{fr(f['duree_s'][H]['rho'], 2, True)}), où les vidéos longues ont aussi "
                   "moins de vues.",
        "hashtag_generique": "Aucun effet : #fyp, #foryou ou #viral ne changent pas l'engagement "
                             "(écarts faibles, jamais significatifs, de signe variable).",
        "nb_hashtags": "Aucun effet dans les tendances ni dans les likées ; négatif seulement "
                       f"dans l'humour (ρ = {fr(f['nb_hashtags'][H]['rho'], 2, True)}), échantillon "
                       "sélectionné sur les likes : à ne pas généraliser.",
        "longueur_texte": "Résultats contradictoires : légèrement positif dans les tendances, "
                          f"négatif dans les likées (ρ = {fr(f['longueur_texte'][L]['rho'], 2, True)}) "
                          ": pas de règle sur la longueur de la légende.",
        "legende_vide": "Aucun effet : une vidéo sans légende ne fait pas moins bien.",
        "format_collab": "Aucun effet mesurable des mentions, duos, stitchs ou réponses "
                         "visibles dans la légende.",
        "question_ou_appel": "Pas d'effet significatif : léger plus dans les tendances "
                             f"({fr(f['question_ou_appel'][T]['ecart'] * 100, 1, True)} pts, p = "
                             f"{pval(f['question_ou_appel'][T]['p'])}), moins dans l'humour.",
        "week_end": "Aucun effet : publier le week-end ou en semaine revient au même.",
        "jour_nom": "Aucun jour de publication ne se détache durablement (écarts non "
                    "significatifs dans les trois échantillons).",
        "creneau_utc": "Aucun créneau horaire ne se détache (écarts non significatifs).",
        "theme": "Aucun sujet ne gagne systématiquement : le classement des sujets change d'un "
                 "échantillon à l'autre et les écarts ne sont pas significatifs.",
        "son_original": "Léger avantage aux musiques existantes sur les sons originaux "
                        f"({fr(f['son_original'][H]['ecart'] * 100, 1, True)} pts pour le son "
                        f"original), non significatif (p = {pval(f['son_original'][H]['p'])}).",
        "stitch_autorise": "Aucun effet du fait d'autoriser les stitchs.",
        "nb_videos_compte": "Publier davantage (nombre de vidéos du compte) n'est pas associé à "
                            "plus d'engagement (lien faible et non significatif).",
    }


def feuille_facteurs(C, s, ws, d, res, calc, refs):
    ws.set_column(0, 0, 1.5)
    ws.set_column(1, 1, 33)
    ws.set_column(2, 2, 21)
    for c in (3, 5, 7):
        ws.set_column(c, c, 11)
    for c in (4, 6, 8):
        ws.set_column(c, c, 8.5)
    ws.set_column(9, 9, 15)
    ws.set_column(10, 10, 62)
    entete(ws, s, "Facteurs de succès : ce que le créateur contrôle… et ce qui n'y change rien",
           "Lecture : pour chaque levier, écart de taux d'engagement médian entre les vidéos qui "
           "ont la caractéristique et les autres (en points), ou corrélation de rang ρ (de −1 à "
           "+1). « Tendances » est l'échantillon principal ; « Likées » et « Humour » servent de "
           "contrôle. Un effet n'est retenu que si p < 0,05.", 10)
    section(ws, s, 3, "A. Leviers que le créateur contrôle : effet sur le taux d'engagement",
            1, 10)
    ws.set_row(4, 32)
    for j, x in enumerate(["Levier testé", "Mesure", "Tendances", "p", "Likées", "p",
                           "Humour", "p", "Verdict", "En une phrase"]):
        ws.write(4, 1 + j, x, s["th"] if j else s["th_g"])
    fac = {x["cle"]: x for x in res["facteurs"]}
    phrases = phrases_facteurs(res)
    graph_rows = {}
    for i, cle in enumerate(ORDRE_FACTEURS):
        r = FA_TABLE + i
        x = fac[cle]
        ws.set_row(r, hauteur((phrases[cle], 62, 10), (x["libelle"], 33, 10)))
        ws.write(r, 1, x["libelle"], s["txt_b"])
        for j, ech in enumerate(ECH):
            c = 3 + 2 * j
            if ech not in x["echantillons"] or (cle == "legende_vide" and ech == "Humour"):
                ws.write(r, c, "—", s["centre"])
                ws.write(r, c + 1, "", s["centre"])
                continue
            rr = x["resultats"][ech]
            if x["type"] == "bin":
                f, v = calc.ecart_pts([("Échantillon", "=", ech)], BIN_COLS[cle])
                ecrire_formule(ws, r, c, f, v, s["pts"])
                ws.write_number(r, c + 1, rr["p"], s["p"])
                mesure = "Écart d'engagement médian (oui − non)"
                graph_rows.setdefault(cle, {})[ech] = (r, c, v)
            elif x["type"] == "num":
                f, v, n = calc.correl(ech, NUM_COLS[cle], "Rang engagement")
                ecrire_formule(ws, r, c, f, v, s["rho"])
                cell = xlsxwriter.utility.xl_rowcol_to_cell(r, c)
                ws.write_formula(r, c + 1, f"=TDIST(ABS({cell}*SQRT(({n}-2)/(1-{cell}^2))),"
                                           f"{n}-2,2)", s["p"], p_correl(v, n))
                mesure = "Corrélation de rang ρ (Spearman)"
            else:
                plage, valeurs = refs[cle][ech]
                f = f"=(MAX({plage})-MIN({plage}))*100"
                v = (max(valeurs) - min(valeurs)) * 100
                ecrire_formule(ws, r, c, f, v, s["pts"])
                ws.write_number(r, c + 1, rr["p"], s["p"])
                mesure = "Écart max entre modalités"
                graph_rows.setdefault(cle, {})[ech] = (r, c, v)
        ws.write(r, 2, mesure, s["txt"])
        ws.write(r, 9, x["verdict"], s["verdict"][x["verdict"]])
        ws.write(r, 10, phrases[cle], s["txt"])
    rn = FA_TABLE + len(ORDRE_FACTEURS)
    ws.set_row(rn, 30)
    ws.merge_range(rn, 1, rn, 10,
                   "p : probabilité d'observer un tel écart par hasard. Écarts et ρ sont des "
                   "formules Excel sur l'onglet Données ; les p des écarts (tests de Mann-Whitney "
                   "et de Kruskal-Wallis) sont calculées par le script Python du projet. Les "
                   "facteurs « son original », « stitch » et « vidéos publiées » n'existent que "
                   "dans l'extrait Humour.", s["note"])

    # Graphique des écarts (facteurs oui/non communs aux trois échantillons).
    section(ws, s, FA_GRAPH, "Seule la durée se détache, et seulement dans les tendances", 1, 10)
    ws.write(FA_GRAPH + 1, 7, "Données du graphique (renvoient au tableau A)", s["note"])
    hdr_r = FA_GRAPH + 2
    for j, x in enumerate(["Levier", "Tendances", "Likées", "Humour"]):
        ws.write(hdr_r, 7 + j, x, s["th"])
    cles = ["duree_15", "hashtag_generique", "legende_vide", "format_collab",
            "question_ou_appel", "week_end"]
    courts = {"duree_15": "Durée > 15 s", "hashtag_generique": "Hashtags #fyp…",
              "legende_vide": "Légende vide", "format_collab": "Mention / duo",
              "question_ou_appel": "Question / appel", "week_end": "Week-end"}
    for i, cle in enumerate(cles):
        r = hdr_r + 1 + i
        ws.write(r, 7, courts[cle], s["txt"])
        for j, ech in enumerate(ECH):
            if ech in graph_rows.get(cle, {}):
                rr, cc, rv = graph_rows[cle][ech]
                cell = xlsxwriter.utility.xl_rowcol_to_cell(rr, cc)
                ws.write_formula(r, 8 + j, f"={cell}", s["pts"], rv)
            else:
                ws.write(r, 8 + j, "", s["pts"])
    ch = C.wb.add_chart({"type": "bar"})
    for j, ech in enumerate(ECH):
        ch.add_series({
            "name": ECH_COURT[ech],
            "categories": ["Facteurs", hdr_r + 1, 7, hdr_r + len(cles), 7],
            "values": ["Facteurs", hdr_r + 1, 8 + j, hdr_r + len(cles), 8 + j],
            "fill": {"color": COULEURS_ECH[ech]}, "gap": 60,
        })
    ch.set_title({"name": "Écart d'engagement médian (avec − sans), en points",
                  "name_font": {"size": 11, "bold": True, "color": NUIT}})
    ch.set_x_axis({"num_format": "+0;-0;0", "major_gridlines": {"visible": True,
                   "line": {"color": "#E5E7EB"}}, "name": "points d'engagement"})
    ch.set_y_axis({"reverse": True})
    ch.set_legend({"position": "bottom"})
    ch.set_size({"width": 700, "height": 330})
    ws.insert_chart(FA_GRAPH + 1, 1, ch, {"description":
                    "Barres horizontales : écart de taux d'engagement médian entre vidéos avec et "
                    "sans chaque caractéristique, pour les trois échantillons. Seule la durée "
                    "supérieure à 15 secondes montre un écart net (+5,5 points), dans les "
                    "tendances uniquement."})

    # B. Conséquences du succès.
    g = res["gros_comptes"]
    section(ws, s, FA_CONSEQ, "B. Conséquences du succès : à connaître, mais pas des leviers",
            1, 10)
    ws.set_row(FA_CONSEQ + 1, 30)
    for j, x in enumerate(["Variable", "Mesure", "Valeur", "p", "", "", "", "", "Statut",
                           "Pourquoi ce n'est pas un levier"]):
        if x:
            ws.write(FA_CONSEQ + 1, 1 + j, x, s["th"] if j else s["th_g"])
    lignes = []
    f, v, n = calc.correl("Humour", "Rang abonnés", "Rang vues")
    lignes.append(("Nombre d'abonnés", "ρ avec les vues (Humour)", (f, v, n), "rho",
                   "Le résultat d'années de succès : on ne peut pas l'« activer » pour une "
                   "campagne. Lien faible avec les vues, voir l'onglet Gros comptes."))
    f, v, n = calc.correl("Humour", "Rang abonnés", "Rang engagement")
    lignes.append(("Nombre d'abonnés", "ρ avec l'engagement (Humour)", (f, v, n), "rho",
                   "Plus le compte est gros, moins son audience réagit en proportion."))
    lignes.append(("Likes cumulés du compte", "ρ avec les vues (Humour)",
                   g["likes_cumules_vues"], "static",
                   "Somme des succès passés. Variable non transmise dans cette version "
                   "(non actionnable, minimisation)."))
    lignes.append(("Badge « vérifié »", "Écart d'engagement (vérifié − non)",
                   g["verifie"], "static_pts",
                   "Attribué par la plateforme aux comptes notoires. Indicateur sur une personne, "
                   "non transmis dans cette version (minimisation)."))
    for i, (lib, mesure, val, kind, why) in enumerate(lignes):
        r = FA_CONSEQ + 2 + i
        ws.set_row(r, hauteur((why, 62, 10), (mesure, 21, 10)))
        ws.write(r, 1, lib, s["txt_b"])
        ws.write(r, 2, mesure, s["txt"])
        if kind == "rho":
            f, v, n = val
            ecrire_formule(ws, r, 3, f, v, s["rho"])
            cell = xlsxwriter.utility.xl_rowcol_to_cell(r, 3)
            ws.write_formula(r, 4, f"=TDIST(ABS({cell}*SQRT(({n}-2)/(1-{cell}^2))),{n}-2,2)",
                             s["p"], p_correl(v, n))
        elif kind == "static":
            ws.write_number(r, 3, val["rho"], s["rho"])
            ws.write_number(r, 4, val["p"], s["p"])
        else:
            ws.write_number(r, 3, val["ecart"] * 100, s["pts"])
            ws.write_number(r, 4, val["p"], s["p"])
        ws.write(r, 9, "Conséquence", s["centre"])
        ws.write(r, 10, why, s["txt"])

    # C. Détail de la durée.
    section(ws, s, FA_DUREE, "C. Détail : la durée", 1, 10)
    for k, (titre, cible) in enumerate((("Part de top performers", "top"),
                                        ("Taux d'engagement médian", "med"))):
        r0 = FA_DUREE + 1 + k * 7
        ws.write(r0, 1, titre, s["th_g"])
        for j, ech in enumerate(ECH):
            ws.write(r0, 3 + 2 * j, ECH_COURT[ech], s["th"])
        for i, cl in enumerate(CLASSES):
            r = r0 + 1 + i
            ws.write(r, 1, cl, s["txt"])
            for j, ech in enumerate(ECH):
                conds = [("Échantillon", "=", ech), ("Classe de durée", "=", cl)]
                if cible == "top":
                    f, v = calc.averageifs("Top performer", conds)
                    ecrire_formule(ws, r, 3 + 2 * j, f, v, s["pct0"])
                else:
                    f, v = calc.median(conds, "Engagement")
                    ecrire_formule(ws, r, 3 + 2 * j, f, v, s["pct"])
    ch = C.wb.add_chart({"type": "column"})
    r0 = FA_DUREE + 1
    for j, ech in enumerate(ECH):
        ch.add_series({"name": ECH_COURT[ech],
                       "categories": ["Facteurs", r0 + 1, 1, r0 + 5, 1],
                       "values": ["Facteurs", r0 + 1, 3 + 2 * j, r0 + 5, 3 + 2 * j],
                       "fill": {"color": COULEURS_ECH[ech]}, "gap": 80})
    ch.set_title({"name": "Part de top performers selon la durée",
                  "name_font": {"size": 11, "bold": True, "color": NUIT}})
    ch.set_y_axis({"num_format": "0%", "major_gridlines": {"visible": True,
                   "line": {"color": "#E5E7EB"}}})
    ch.set_legend({"position": "bottom"})
    ch.set_size({"width": 520, "height": 270})
    ws.insert_chart(FA_DUREE + 1, 10, ch, {"x_offset": 10, "description":
                    "Colonnes : part de top performers par classe de durée, pour les trois "
                    "échantillons. Dans les tendances, les vidéos de 16 à 30 secondes ont la plus "
                    "forte part de top performers ; le schéma ne se répète pas ailleurs."})

    # D. Jours et créneaux (sources des écarts « max − min » du tableau A).
    section(ws, s, FA_JOURS, "D. Détail : jour et créneau de publication (taux d'engagement "
                             "médian, heure UTC)", 1, 10)
    for k, (col, modalites, titre) in enumerate((("Jour", JOURS, "Jour"),
                                                 ("Créneau UTC", CRENEAUX, "Créneau (UTC)"))):
        r0 = FA_JOURS + 1 + k * 9
        ws.write(r0, 1, titre, s["th_g"])
        for j, ech in enumerate(ECH):
            ws.write(r0, 3 + 2 * j, ECH_COURT[ech], s["th"])
        for i, m in enumerate(modalites):
            r = r0 + 1 + i
            ws.write(r, 1, m, s["txt"])
            for j, ech in enumerate(ECH):
                conds = [("Échantillon", "=", ech), (col, "=", m)]
                if calc._mask(conds).sum() == 0:
                    ws.write(r, 3 + 2 * j, "—", s["centre"])
                    continue
                f, v = calc.median(conds, "Engagement")
                ecrire_formule(ws, r, 3 + 2 * j, f, v, s["pct"])
    # E. TCD.
    section(ws, s, FA_TCD - 2, "E. Tableau croisé dynamique : part de top performers par "
                               "classe de durée et par échantillon (se met à jour à "
                               "l'ouverture)", 1, 10)
    ws.set_zoom(90)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.freeze_panes(5, 2)
    return {"hashtag": (FA_TABLE + ORDRE_FACTEURS.index("hashtag_generique"), 3,
                        graph_rows["hashtag_generique"]["Tendances"][2]),
            "duree": (FA_TABLE + ORDRE_FACTEURS.index("duree_15"), 3,
                      graph_rows["duree_15"]["Tendances"][2]),
            "theme": (FA_TABLE + ORDRE_FACTEURS.index("theme"), 3,
                      graph_rows["theme"]["Tendances"][2])}


def refs_categoriels(d, calc):
    """Plages (et valeurs) des médianes par modalité utilisées pour les écarts « max − min »."""
    out = {"jour_nom": {}, "creneau_utc": {}, "theme": {}}
    for k, (cle, col, modalites) in enumerate((("jour_nom", "Jour", JOURS),
                                              ("creneau_utc", "Créneau UTC", CRENEAUX))):
        r0 = FA_JOURS + 1 + k * 9
        for j, ech in enumerate(ECH):
            c = 3 + 2 * j
            plage = (f"{xlsxwriter.utility.xl_rowcol_to_cell(r0 + 1, c)}:"
                     f"{xlsxwriter.utility.xl_rowcol_to_cell(r0 + len(modalites), c)}")
            vals = [calc.median([("Échantillon", "=", ech), (col, "=", m)], "Engagement")[1]
                    for m in modalites if calc._mask([("Échantillon", "=", ech),
                                                      (col, "=", m)]).sum()]
            out[cle][ech] = (plage, vals)
    for j, ech in enumerate(("Tendances", "Humour")):
        c = 2 + j
        plage = (f"'Catégories & TreeMap'!{xlsxwriter.utility.xl_rowcol_to_cell(CA_THEME_ECH, c)}:"
                 f"{xlsxwriter.utility.xl_rowcol_to_cell(CA_THEME_ECH + len(THEMES) - 1, c)}")
        vals = [calc.median([("Échantillon", "=", ech), ("Thème", "=", t)], "Engagement")[1]
                for t in THEMES if calc._mask([("Échantillon", "=", ech),
                                               ("Thème", "=", t)]).sum()]
        out["theme"][ech] = (plage, vals)
    return out


# --------------------------------------------------------------------------- #
# Onglet « Gros comptes »
# --------------------------------------------------------------------------- #
GC_TRANCHES = 10        # première ligne du tableau par tranche
GC_CORREL = 18          # première ligne du tableau des corrélations
GC_TCD = 55             # ligne de la zone du TCD (filtre deux lignes au-dessus)


def p_partielle(r, n):
    t = abs(r) * math.sqrt((n - 3) / (1 - r * r))
    return float(2 * stats.t.sf(t, n - 3))


def feuille_gros_comptes(C, s, ws, d, res, calc):
    g = res["gros_comptes"]
    tr = {x["tranche"]: x for x in g["tranches"]}
    t1, t4 = tr[TRANCHES[0]], tr[TRANCHES[-1]]
    ws.set_column(0, 0, 1.5)
    ws.set_column(1, 1, 36)
    ws.set_column(2, 8, 14)
    ws.set_column(9, 9, 34)
    entete(ws, s, "« Les gros comptes gagneraient de toute façon » : ce que disent les données",
           "Seul l'extrait Humour (top 100 des vidéos #funny) contient le nombre d'abonnés : "
           "c'est donc sur lui que porte la réponse. Les tendances ne permettent pas de répondre.",
           9)
    section(ws, s, 3, "La réponse courte", 1, 9)
    ws.write(3, 8, "Chiffre", s["section"])
    ws.write(3, 9, "Ce que mesure le chiffre", s["section"])
    stmts = [
        ("Pour la portée : NON. Les comptes de plus de 8 M d'abonnés en ont "
         f"{fr(t4['abonnes_med'] / t1['abonnes_med'], 0)} fois plus que ceux de moins de 1 M, "
         f"mais n'obtiennent que {fr(t4['vues_med'] / t1['vues_med'], 1)} fois plus de vues "
         f"médianes ({fr(t4['vues_med'] / 1e6)} M contre {fr(t1['vues_med'] / 1e6)} M).",
         "rho", "corrélation de rang abonnés ↔ vues (0 = aucun lien, 1 = lien parfait)"),
        ("Pour l'engagement : plutôt l'INVERSE. Sous 1 M d'abonnés, "
         f"{pct(t1['eng_med'])} d'engagement médian et {pct(t1['part_top'], 0)} de top "
         f"performers ; au-delà de 8 M, {pct(t4['eng_med'])} et {pct(t4['part_top'], 0)}. "
         "À vues égales, l'écart se réduit : tendance à confirmer.",
         "rho2", "corrélation de rang abonnés ↔ engagement (négative = moins d'engagement "
                 "quand le compte grossit)"),
        ("Par abonné, les petits comptes rapportent bien plus : "
         f"{fr(t1['vues_par_abonne_med'], 0)} vues par abonné (médiane) sous 1 M, contre "
         f"{fr(t4['vues_par_abonne_med'], 1)} au-delà de 8 M. Intéressant si la rémunération "
         "suit le nombre d'abonnés.",
         "ratio", "rapport des vues par abonné : moins de 1 M / plus de 8 M"),
    ]
    rho_cells = {}
    for i, (txt, kind, lab) in enumerate(stmts):
        r = 4 + i
        ws.set_row(r, hauteur((txt, 120, 10), (lab, 34, 9), minimum=30))
        ws.merge_range(r, 1, r, 7, txt, s["txt"])
        ws.write(r, 9, lab, s["note"])
        rho_cells[kind] = r
    # Tableau par tranche.
    section(ws, s, 8, "Par taille de compte (top 100 #funny, 100 vidéos, 79 créateurs)", 1, 9)
    hdr = ["Tranche d'abonnés", "Vidéos", "Créateurs", "Abonnés (médiane)",
           "Vues (médiane)", "Vues par abonné (médiane)", "Engagement (médiane)",
           "Part de top performers"]
    ws.set_row(9, 32)
    for j, x in enumerate(hdr):
        ws.write(9, 1 + j, x, s["th"] if j else s["th_g"])
    lignes = [(t, [("Échantillon", "=", "Humour"), ("Tranche abonnés", "=", t)])
              for t in TRANCHES] + [("Ensemble (Humour)", [("Échantillon", "=", "Humour")])]
    for i, (lib, conds) in enumerate(lignes):
        r = GC_TRANCHES + i
        st = s["txt_b"] if i == len(lignes) - 1 else s["txt"]
        ws.write(r, 1, lib, st)
        f, v = calc.countifs(conds)
        ecrire_formule(ws, r, 2, f, v, s["int"])
        expr = Calc._expr(conds)
        ws.write_array_formula(
            r, 3, r, 3, f"{{=SUM(IF({expr},1/COUNTIFS({ref('Créateur')},{ref('Créateur')},"
                        f"{ref('Échantillon')},\"Humour\")))}}", s["int"],
            int(d.loc[calc._mask(conds), "Créateur"].nunique()))
        for c, cible, fm in ((4, "Abonnés", "m"), (5, "Vues", "m"), (6, "Vues par abonné", "dec1"),
                             (7, "Engagement", "pct")):
            f, v = calc.median(conds, cible)
            ecrire_formule(ws, r, c, f, v, s[fm])
        f, v = calc.averageifs("Top performer", conds)
        ecrire_formule(ws, r, 8, f, v, s["pct0"])
    # Corrélations.
    section(ws, s, GC_CORREL - 2, "Corrélations (ρ de Spearman)", 1, 9)
    ws.set_row(GC_CORREL - 1, 30)
    for j, x in enumerate(["Relation", "ρ", "p", "n", "Lecture"]):
        ws.write(GC_CORREL - 1, 1 + j, x, s["th"] if j else s["th_g"])
    ws.merge_range(GC_CORREL - 1, 5, GC_CORREL - 1, 9, "Lecture", s["th"])
    part = g["partielle_abonnes_engagement_a_vues_egales"]
    correls = [
        ("Abonnés ↔ vues (par vidéo)", ("Rang abonnés", "Rang vues"),
         "Lien positif mais faible : la taille du compte n'explique qu'une petite partie de la "
         "portée."),
        ("Abonnés ↔ engagement (par vidéo)", ("Rang abonnés", "Rang engagement"),
         "Plus le compte est gros, moins son audience réagit en proportion."),
        ("Abonnés ↔ vues médianes (par créateur)", g["abonnes_vues_par_createur"],
         "Même conclusion en ne comptant qu'une fois chaque créateur (calcul Python : médiane "
         "par créateur)."),
        ("Abonnés ↔ engagement, à vues égales (partielle)", {"rho": part, "n": 100,
                                                            "p": p_partielle(part, 100)},
         "Une fois les vues neutralisées, le lien s'affaiblit et n'est plus significatif "
         "(calcul Python)."),
        ("Vues ↔ engagement (dans cet extrait)", ("Rang vues", "Rang engagement"),
         "Très négatif : effet de la sélection sur les likes (≥ 6,8 M), pas un comportement du "
         "public. D'où l'usage des vues pour juger les gros comptes."),
    ]
    exact = {}
    for i, (lib, src, lect) in enumerate(correls):
        r = GC_CORREL + i
        ws.set_row(r, hauteur((lect, 90, 10), (lib, 36, 10)))
        ws.write(r, 1, lib, s["txt_b"])
        if isinstance(src, tuple):
            f, v, n = calc.correl("Humour", *src)
            exact[i] = v
            ecrire_formule(ws, r, 2, f, v, s["rho"])
            cell = xlsxwriter.utility.xl_rowcol_to_cell(r, 2)
            ws.write_formula(r, 3, f"=TDIST(ABS({cell}*SQRT(({n}-2)/(1-{cell}^2))),{n}-2,2)",
                             s["p"], p_correl(v, n))
        else:
            ws.write_number(r, 2, src["rho"], s["rho"])
            ws.write_number(r, 3, src["p"], s["p"])
            n = src["n"]
        ws.write_number(r, 4, n, s["int"])
        ws.merge_range(r, 5, r, 9, lect, s["txt"])
    # Chiffres de la réponse courte (renvoient aux tableaux ci-dessous).
    c_rho = xlsxwriter.utility.xl_rowcol_to_cell(GC_CORREL, 2)
    c_rho2 = xlsxwriter.utility.xl_rowcol_to_cell(GC_CORREL + 1, 2)
    c_vpa1 = xlsxwriter.utility.xl_rowcol_to_cell(GC_TRANCHES, 6)
    c_vpa4 = xlsxwriter.utility.xl_rowcol_to_cell(GC_TRANCHES + 3, 6)
    vpa = [calc.median([("Échantillon", "=", "Humour"), ("Tranche abonnés", "=", t)],
                       "Vues par abonné")[1] for t in (TRANCHES[0], TRANCHES[-1])]
    eng1 = calc.median([("Échantillon", "=", "Humour"), ("Tranche abonnés", "=", TRANCHES[0])],
                       "Engagement")[1]
    ws.write_formula(rho_cells["rho"], 8, f"={c_rho}", s["rho_big"], exact[0])
    ws.write_formula(rho_cells["rho2"], 8, f"={c_rho2}", s["rho_big"], exact[1])
    ws.write_formula(rho_cells["ratio"], 8, f"={c_vpa1}/{c_vpa4}", s["x_big"], vpa[0] / vpa[1])
    # Permet / ne permet pas.
    r = GC_CORREL + len(correls) + 1
    ws.set_row(r, 20)
    ws.merge_range(r, 1, r, 4, "Ce que les données permettent de répondre", s["section"])
    ws.merge_range(r, 5, r, 9, "Ce qu'elles ne permettent pas de dire", s["section"])
    permet = [
        "Parmi les vidéos qui ont percé, la taille du compte n'explique qu'une petite partie "
        f"de la portée (ρ = {fr(g['abonnes_vues']['rho'], 2, True)}).",
        "Des comptes de moins de 1 M d'abonnés atteignent des dizaines de millions de vues : "
        "la portée ne dépend pas que du nombre d'abonnés.",
        f"L'engagement relatif baisse avec la taille du compte ({pct(t1['eng_med'])} → "
        f"{pct(t4['eng_med'])}).",
        "Le résultat tient en comptant chaque créateur une seule fois "
        f"(ρ = {fr(g['abonnes_vues_par_createur']['rho'], 2, True)} sur "
        f"{g['abonnes_vues_par_createur']['n']} créateurs).",
    ]
    nepermet = [
        "La probabilité de percer : les petits comptes qui n'ont pas percé ne sont pas dans les "
        "données (biais du survivant).",
        "La causalité : les abonnés sont comptés en août 2021, après les vidéos ; le succès a pu "
        "créer les abonnés.",
        "La situation dans les tendances (abonnés absents de l'extrait) ou hors du registre "
        "humoristique.",
        f"Si l'écart d'engagement tient à vues égales : il s'affaiblit (ρ partiel = "
        f"{fr(part, 2, True)}) et n'est plus significatif.",
    ]
    for i, (a, b) in enumerate(zip(permet, nepermet)):
        rr = r + 1 + i
        ws.set_row(rr, hauteur(("• " + a, 78, 10), ("• " + b, 90, 10)))
        ws.merge_range(rr, 1, rr, 4, "• " + a, s["txt"])
        ws.merge_range(rr, 5, rr, 9, "• " + b, s["txt"])
    # Graphique.
    rg = r + 6
    section(ws, s, rg, f"{fr(t4['abonnes_med'] / t1['abonnes_med'], 0)} fois plus d'abonnés, "
                       f"{fr(t4['vues_med'] / t1['vues_med'], 1)} fois plus de vues", 1, 9)
    col = C.wb.add_chart({"type": "column"})
    col.add_series({"name": "Vues médianes",
                    "categories": ["Gros comptes", GC_TRANCHES, 1, GC_TRANCHES + 3, 1],
                    "values": ["Gros comptes", GC_TRANCHES, 5, GC_TRANCHES + 3, 5],
                    "fill": {"color": NUIT}, "gap": 70,
                    "data_labels": {"value": True, "num_format": '0,," M"'}})
    line = C.wb.add_chart({"type": "line"})
    line.add_series({"name": "Engagement médian (axe de droite)",
                     "categories": ["Gros comptes", GC_TRANCHES, 1, GC_TRANCHES + 3, 1],
                     "values": ["Gros comptes", GC_TRANCHES, 7, GC_TRANCHES + 3, 7],
                     "y2_axis": True, "line": {"color": ORANGE, "width": 2.5},
                     "marker": {"type": "circle", "size": 7, "fill": {"color": ORANGE},
                                "border": {"color": ORANGE}},
                     "data_labels": {"value": True, "num_format": "0.0%", "position": "above"}})
    col.combine(line)
    col.set_title({"name": "Vues et engagement médians par tranche d'abonnés",
                   "name_font": {"size": 11, "bold": True, "color": NUIT}})
    col.set_y_axis({"num_format": '0,," M"', "min": 0,
                    "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}}})
    line.set_y2_axis({"num_format": "0%", "min": 0})
    col.set_legend({"position": "bottom"})
    col.set_size({"width": 720, "height": 330})
    ws.insert_chart(rg + 1, 1, col, {"description":
                    "Colonnes : vues médianes par tranche d'abonnés ; courbe : engagement médian. "
                    "Les vues varient peu d'une tranche à l'autre alors que l'engagement baisse "
                    "quand le compte grossit."})
    section(ws, s, GC_TCD - 4, "Tableau croisé dynamique : vidéos, top performers et engagement "
                               "moyen par tranche (filtre : Humour ; se met à jour à "
                               "l'ouverture)", 1, 9)
    ws.set_zoom(90)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return {"rho": (GC_CORREL, 2, exact[0]), "eng1": (GC_TRANCHES, 7, eng1)}


# --------------------------------------------------------------------------- #
# Onglet « Catégories & TreeMap »
# --------------------------------------------------------------------------- #
def feuille_categories(C, s, ws, d, res, calc):
    ws.set_column(0, 0, 1.5)
    ws.set_column(1, 1, 26)
    ws.set_column(2, 2, 24)
    ws.set_column(3, 3, 11)
    ws.set_column(4, 9, 13)
    ws.set_column(10, 10, 62)
    entete(ws, s, "Qui se distingue ? Sujets de contenu, leaders et TreeMap",
           "Catégories : 5 sujets déduits des hashtags et légendes, attribués à chaque créateur. "
           "Leader = meilleur taux d'engagement, rapporté à la médiane de son échantillon "
           "(indice), parmi les créateurs dont la portée médiane atteint au moins la médiane de "
           "leur échantillon. Périmètre : Tendances + Humour.", 10)
    hors_likees = ("Échantillon", "<>", "Likées par @tiktok")
    compte = ("Compté dans les cumuls", "=", "Oui")
    section(ws, s, CA_THEMES - 2, "A. Les 5 sujets de contenu", 1, 10)
    ws.set_row(CA_THEMES - 1, 32)
    for j, x in enumerate(["Sujet", "Créateurs", "Vidéos", "Vues cumulées", "Part des vues",
                           "Engagement médian", "Part de top performers"]):
        ws.write(CA_THEMES - 1, 1 + j, x, s["th"] if j else s["th_g"])
    ws.merge_range(CA_THEMES - 1, 8, CA_THEMES - 1, 10, "Lecture", s["th"])
    them = {x["theme"]: x for x in res["themes"]}
    lectures = {
        "Animaux": "Contenus d'animaux : portée solide, engagement moyen.",
        "Famille & enfants": "Engagement élevé ; présence fréquente de mineurs (voir "
                             "vigilance).",
        "Humour & sketchs": "Le plus gros volume de vues, porté par l'échantillon #funny.",
        "Lifestyle, musique & sport": "Sujet hétérogène (musique, sport, cuisine, beauté…).",
        "Sujet non déclaré": "Légende vide ou sans mot-clé : souvent des comptes de "
                             "personnalités.",
    }
    total_vues = calc.sumifs("Vues", [hors_likees, compte])[1]
    for i, t in enumerate(THEMES):
        r = CA_THEMES + i
        ws.set_row(r, 22)
        ws.write(r, 1, t, C.fmt(bold=True, font_color="#FFFFFF", bg_color="#" +
                                COULEURS_THEMES[t], border=1, border_color="#FFFFFF"))
        conds = [("Thème", "=", t), hors_likees]
        ws.write_array_formula(
            r, 2, r, 2, f"{{=SUM(IF({Calc._expr(conds)},1/COUNTIFS({ref('Créateur')},"
                        f"{ref('Créateur')},{ref('Échantillon')},\"<>Likées par @tiktok\")))}}",
            s["int"], them[t]["createurs"])
        f, v = calc.countifs(conds + [compte])
        ecrire_formule(ws, r, 3, f, v, s["int"])
        f, v = calc.sumifs("Vues", conds + [compte])
        ecrire_formule(ws, r, 4, f, v, s["m"])
        c1 = xlsxwriter.utility.xl_rowcol_to_cell(r, 4)
        tot = xlsxwriter.utility.xl_range_abs(CA_THEMES, 4, CA_THEMES + len(THEMES) - 1, 4)
        ws.write_formula(r, 5, f"={c1}/SUM({tot})", s["pct"], v / total_vues)
        f, v = calc.median(conds + [compte], "Engagement")
        ecrire_formule(ws, r, 6, f, v, s["pct"])
        f, v = calc.averageifs("Top performer", conds + [compte])
        ecrire_formule(ws, r, 7, f, v, s["pct0"])
        ws.merge_range(r, 8, r, 10, lectures[t], s["txt"])
    conc = res["concentration"]
    rn = CA_THEMES + len(THEMES)
    ws.set_row(rn, 30)
    ws.merge_range(rn, 1, rn, 10,
                   f"Concentration : {pct(conc['top10_n'] / conc['createurs'], 0)} des créateurs "
                   f"({conc['top10_n']} sur {conc['createurs']}) captent "
                   f"{pct(conc['top10_part_vues'], 0)} des vues cumulées (calcul du script "
                   "Python, sur les vues cumulées par créateur). Aucun sujet ne se détache "
                   "nettement en engagement : voir le tableau E et l'onglet Facteurs.", s["note"])

    # B. Leaders par sujet.
    section(ws, s, CA_LEADERS - 2, "B. Leaders par sujet : meilleur engagement, à portée au "
                                   "moins médiane", 1, 10)
    ws.set_row(CA_LEADERS - 1, 32)
    for j, x in enumerate(["Sujet", "Créateur", "Rang", "Échantillon", "Vidéos",
                           "Vues médianes", "Engagement médian", "Indice d'engagement",
                           "Abonnés", "Pourquoi"]):
        ws.write(CA_LEADERS - 1, 1 + j, x, s["th"] if j else s["th_g"])
    for i, x in enumerate(res["leaders"]):
        r = CA_LEADERS + i
        ws.set_row(r, hauteur((x["pourquoi"], 62, 10)))
        cr, ech = "@" + x["createur"], x["echantillon"]
        conds = [("Créateur", "=", cr), ("Échantillon", "=", ech)]
        ws.write(r, 1, x["theme"], s["txt"])
        ws.write(r, 2, cr, s["txt_b"])
        ws.write_number(r, 3, x["rang"], s["int"])
        ws.write(r, 4, ech, s["centre"])
        f, v = calc.countifs(conds)
        ecrire_formule(ws, r, 5, f, v, s["int"])
        f, v = calc.median(conds, "Vues")
        ecrire_formule(ws, r, 6, f, v, s["m"])
        f, v = calc.median(conds, "Engagement")
        ecrire_formule(ws, r, 7, f, v, s["pct"])
        c_eng = xlsxwriter.utility.xl_rowcol_to_cell(r, 7)
        c_ech = xlsxwriter.utility.xl_rowcol_to_cell(r, 4)
        med_ech = calc.median([("Échantillon", "=", ech)], "Engagement")[1]
        ws.write_formula(r, 8, f"={c_eng}/VLOOKUP({c_ech},{SEUILS_PLAGE},3,FALSE)", s["x"],
                         v / med_ech)
        if ech == "Humour":
            f, v = calc.averageifs("Abonnés", conds)
            ecrire_formule(ws, r, 9, f, v, s["m"])
        else:
            ws.write(r, 9, "—", s["centre"])
        ws.write(r, 10, x["pourquoi"], s["txt"])

    # C. Leaders de portée.
    section(ws, s, CA_PORTEE - 2, "C. Leaders de portée : le plus de vues cumulées par sujet",
            1, 10)
    ws.set_row(CA_PORTEE - 1, 30)
    for j, x in enumerate(["Sujet", "Créateur", "Vidéos", "Vues cumulées"]):
        ws.write(CA_PORTEE - 1, 1 + j, x, s["th"] if j else s["th_g"])
    ws.merge_range(CA_PORTEE - 1, 5, CA_PORTEE - 1, 10, "À noter", s["th"])
    for i, x in enumerate(res["leaders_portee"]):
        r = CA_PORTEE + i
        cr = "@" + x["createur"]
        conds = [("Créateur", "=", cr), hors_likees, compte]
        ws.write(r, 1, x["theme"], s["txt"])
        ws.write(r, 2, cr, s["txt_b"])
        f, v = calc.countifs(conds)
        ecrire_formule(ws, r, 3, f, v, s["int"])
        f, v = calc.sumifs("Vues", conds)
        ecrire_formule(ws, r, 4, f, v, s["m"])
        eng = float(d.loc[calc._mask(conds), "Engagement"].median())
        ws.merge_range(r, 5, r, 10, f"Portée la plus forte du sujet ; engagement médian "
                                    f"{pct(eng)} (à comparer avec les leaders d'engagement).",
                       s["txt"])

    # D. TreeMap.
    section(ws, s, CA_TREEMAP - 3, "D. TreeMap : qui capte les vues ?", 1, 10)
    ws.set_row(CA_TREEMAP - 2, 32)
    ws.merge_range(CA_TREEMAP - 2, 1, CA_TREEMAP - 2, 10,
                   "Lecture : chaque tuile est un créateur ; plus elle est grande, plus il a "
                   "cumulé de vues (Tendances + Humour). La couleur indique le sujet du créateur. "
                   "Les créateurs non retenus sont regroupés par sujet (« Autres créateurs »). "
                   "La taille dit la portée, la couleur dit le sujet : deux informations "
                   "différentes.", s["txt_nb"])
    for j, x in enumerate(["Sujet", "Tuile", "Vues cumulées"]):
        ws.write(CA_TREEMAP - 1, 1 + j, x, s["th"] if j else s["th_g"])
    debut = {}
    for i, x in enumerate(res["treemap"]):
        r = CA_TREEMAP + i
        t = x["theme"]
        debut.setdefault(t, r)
        ws.write(r, 1, t, s["txt"])
        ws.write(r, 2, x["tuile"], s["txt_b"] if x["tuile"].startswith("@") else s["txt"])
        if x["tuile"].startswith("@"):
            f, v = calc.sumifs("Vues", [("Créateur", "=", x["tuile"]), hors_likees, compte])
            ecrire_formule(ws, r, 3, f, v, s["m"])
        else:
            f_tot, v_tot = calc.sumifs("Vues", [("Thème", "=", t), hors_likees, compte])
            nommes = (f"{xlsxwriter.utility.xl_rowcol_to_cell(debut[t], 3)}:"
                      f"{xlsxwriter.utility.xl_rowcol_to_cell(r - 1, 3)}")
            v = x["vues"]
            ws.write_formula(r, 3, f"{f_tot}-SUM({nommes})", s["m"], v)
    fin = CA_TREEMAP + len(res["treemap"])
    ws.set_row(fin, 30)
    ws.merge_range(fin, 1, fin, 3, "Ces 3 colonnes alimentent le TreeMap (Insertion > "
                                   "Graphique hiérarchique > Carte proportionnelle).", s["note"])

    # E. Engagement par sujet et par échantillon (source de l'écart « sujet » des Facteurs).
    section(ws, s, CA_THEME_ECH - 2, "E. Taux d'engagement médian par sujet et par échantillon",
            1, 10)
    for j, x in enumerate(["Sujet", "Tendances", "Humour"]):
        ws.write(CA_THEME_ECH - 1, 1 + j, x, s["th"] if j else s["th_g"])
    for i, t in enumerate(THEMES):
        r = CA_THEME_ECH + i
        ws.write(r, 1, t, s["txt"])
        for j, ech in enumerate(("Tendances", "Humour")):
            conds = [("Échantillon", "=", ech), ("Thème", "=", t)]
            if calc._mask(conds).sum() == 0:
                ws.write(r, 2 + j, "—", s["centre"])
                continue
            f, v = calc.median(conds, "Engagement")
            ecrire_formule(ws, r, 2 + j, f, v, s["pct"])
    ch = C.wb.add_chart({"type": "bar"})
    for j, ech in enumerate(("Tendances", "Humour")):
        ch.add_series({"name": ech,
                       "categories": ["Catégories & TreeMap", CA_THEME_ECH, 1,
                                      CA_THEME_ECH + len(THEMES) - 1, 1],
                       "values": ["Catégories & TreeMap", CA_THEME_ECH, 2 + j,
                                  CA_THEME_ECH + len(THEMES) - 1, 2 + j],
                       "fill": {"color": COULEURS_ECH[ech]}, "gap": 60})
    ch.set_title({"name": "Aucun sujet ne gagne partout (engagement médian)",
                  "name_font": {"size": 11, "bold": True, "color": NUIT}})
    ch.set_x_axis({"num_format": "0%", "min": 0,
                   "major_gridlines": {"visible": True, "line": {"color": "#E5E7EB"}}})
    ch.set_y_axis({"reverse": True})
    ch.set_legend({"position": "bottom"})
    ch.set_size({"width": 560, "height": 300})
    ws.insert_chart(CA_THEME_ECH - 1, 4, ch, {"x_offset": 20, "description":
                    "Barres : taux d'engagement médian par sujet, tendances et humour. Le "
                    "classement des sujets change d'un échantillon à l'autre."})

    # F. Vigilance.
    rv = CA_THEME_ECH + 16
    section(ws, s, rv, "F. Points de vigilance avant de contacter un créateur", 1, 10)
    vig = [
        "Famille & enfants : des mineurs apparaissent dans plusieurs vidéos. Toute collaboration "
        "doit respecter la loi n° 2020-1266 du 19 octobre 2020 (image des enfants de moins de "
        "16 ans sur les plateformes) et la loi n° 2023-451 du 9 juin 2023 (influence "
        "commerciale) : validation juridique préalable.",
        "Plusieurs leaders sont des célébrités ou des institutions (groupe de musique, club de "
        "football, artistes) plutôt que des créateurs : conditions et tarifs sans rapport avec "
        "ceux des créateurs.",
        "Les chiffres datent de 2021 : vérifier l'activité et l'audience actuelles des comptes "
        "avant tout contact.",
        "Le sujet est déduit des légendes : le confirmer en regardant les vidéos.",
    ]
    for i, txt in enumerate(vig):
        ws.set_row(rv + 1 + i, hauteur(("• " + txt, 200, 10)))
        ws.merge_range(rv + 1 + i, 1, rv + 1 + i, 10, "• " + txt, s["puce"])
    section(ws, s, CA_TCD - 2, "G. Tableau croisé dynamique : vues par sujet et par échantillon "
                               "(se met à jour à l'ouverture)", 1, 10)
    ws.set_zoom(90)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


# --------------------------------------------------------------------------- #
# Onglet « Synthèse »
# --------------------------------------------------------------------------- #
def feuille_synthese(C, s, ws, d, res, pos_f, pos_g):
    fac = {x["cle"]: x["resultats"] for x in res["facteurs"]}
    g = res["gros_comptes"]
    tr = {x["tranche"]: x for x in g["tranches"]}
    t1, t4 = tr[TRANCHES[0]], tr[TRANCHES[-1]]
    e = res["echantillons"]
    largeurs = [1.5, 3.5, 44, 12, 54, 13, 2, 23, 21, 21]
    for j, w in enumerate(largeurs):
        ws.set_column(j, j, w)
    ws.hide_gridlines(2)
    ws.set_row(0, 30)
    ws.write(0, 1, "TikTok : ce qui fait vraiment performer un contenu", s["titre"])
    ws.set_row(1, 34)
    n_tot = sum(x["videos"] for x in e.values())
    ws.merge_range(
        1, 1, 1, 9,
        f"Synthèse pour la campagne · {n_tot} vidéos populaires collectées en août 2021 "
        "(tendances, top #funny, vidéos likées par @tiktok). Succès = taux d'engagement = "
        "(likes + commentaires + partages) ÷ vues ; top performers = le quart supérieur de "
        "chaque échantillon. Chaque conclusion renvoie à l'onglet qui la prouve.",
        s["sous_titre"])
    ws.set_row(2, 6)
    ws.set_row(3, 20)
    ws.merge_range(3, 1, 3, 5, "LES CONCLUSIONS", s["section"])
    ws.merge_range(3, 7, 3, 9, "QUI SE DISTINGUE : LEADERS PAR SUJET", s["section"])
    ws.set_row(4, 30)
    for j, x in enumerate(["", "Conclusion", "Chiffre clé", "Ce que dit le chiffre", "Preuve"]):
        ws.write(4, 1 + j, x, s["th"] if j != 1 else s["th_g"])
    for j, x in enumerate(["Sujet", "Leader d'engagement", "Leader de portée"]):
        ws.write(4, 7 + j, x, s["th"] if j else s["th_g"])

    hf, dur, th = fac["hashtag_generique"], fac["duree_15"], fac["theme"]
    T = "Tendances"
    rc = lambda sh, pos: f"='{sh}'!{xlsxwriter.utility.xl_rowcol_to_cell(pos[0], pos[1])}"
    conclusions = [
        ("La taille du compte ne fait pas la portée.",
         rc("Gros comptes", pos_g["rho"]), pos_g["rho"][2], "rho_big",
         f"Corrélation abonnés ↔ vues, faible. Dans le top 100 #funny, "
         f"{fr(t4['abonnes_med'] / t1['abonnes_med'], 0)} fois plus d'abonnés (plus de 8 M "
         f"contre moins de 1 M) ne donnent que {fr(t4['vues_med'] / t1['vues_med'], 1)} fois "
         f"plus de vues médianes.", ("Gros comptes", "B4")),
        ("Les petits comptes font davantage réagir leur audience.",
         rc("Gros comptes", pos_g["eng1"]), pos_g["eng1"][2], "pct_big",
         f"Engagement médian sous 1 M d'abonnés, contre {pct(t4['eng_med'])} au-delà de 8 M "
         f"({pct(t1['part_top'], 0)} de top performers contre {pct(t4['part_top'], 0)}). "
         "Tendance à confirmer : l'écart se réduit à vues égales.", ("Gros comptes", "B10")),
        ("Les « recettes » de publication n'y changent rien.",
         rc("Facteurs", pos_f["hashtag"]), pos_f["hashtag"][2], "pts_big",
         f"Écart d'engagement avec ou sans #fyp/#foryou dans les tendances (p = "
         f"{pval(hf[T]['p'])}). Jour, heure, appel à l'action, mentions, son original : aucun "
         "effet significatif dans les 3 échantillons.", ("Facteurs", "B6")),
        ("Seul levier qui ressort : des vidéos un peu plus longues.",
         rc("Facteurs", pos_f["duree"]), pos_f["duree"][2], "pts_big",
         f"Vidéos de plus de 15 s : {pct(dur[T]['med_oui'])} d'engagement contre "
         f"{pct(dur[T]['med_non'])} dans les tendances ({pct(dur[T]['top_oui'], 0)} de top "
         f"performers contre {pct(dur[T]['top_non'], 0)}, p = {pval(dur[T]['p'])}). Non "
         "confirmé ailleurs : à tester, pas à imposer.", ("Facteurs", "B6")),
        ("Aucun sujet ne gagne à coup sûr : choisir le sujet pour la marque, puis le "
         "créateur pour son engagement.",
         rc("Facteurs", pos_f["theme"]), pos_f["theme"][2], "pts_big",
         f"Écart maximal d'engagement médian entre sujets dans les tendances, non significatif "
         f"(p = {pval(th[T]['p'])}). Des leaders existent dans chaque sujet (ci-contre).",
         ("Catégories & TreeMap", "B13")),
    ]
    for i, (txt, f, v, fm, detail, cible) in enumerate(conclusions):
        r = 5 + i
        ws.set_row(r, hauteur((txt, 44, 11), (detail, 54, 10), minimum=42))
        ws.write(r, 1, i + 1, C.fmt(bold=True, font_size=14, font_color="#FFFFFF",
                                    bg_color=ORANGE, align="center"))
        ws.write(r, 2, txt, C.fmt(bold=True, font_size=11, text_wrap=True, border=1,
                                  border_color="#E5E7EB", font_color=NUIT))
        ws.write_formula(r, 3, f, s[fm], v)
        ws.write(r, 4, detail, s["txt"])
        lien(ws, s, r, 5, cible[0], cible[1], f"→ {cible[0].split(' &')[0]}")
    lp = {x["theme"]: "@" + x["createur"] for x in res["leaders_portee"]}
    le = {x["theme"]: "@" + x["createur"] for x in res["leaders"] if x["rang"] == 1}
    for i, t in enumerate(THEMES):
        r = 5 + i
        ws.write(r, 7, t, C.fmt(bold=True, font_color="#FFFFFF", bg_color="#" +
                                COULEURS_THEMES[t], text_wrap=True, border=1,
                                border_color="#FFFFFF"))
        ws.write(r, 8, le[t], s["centre"])
        ws.write(r, 9, lp[t], s["centre"])
    ws.set_row(10, 18)
    lien(ws, s, 10, 7, "Catégories & TreeMap", "B13",
         "→ 3 leaders par sujet, leurs chiffres et le TreeMap")
    ws.set_row(11, 8)
    ws.set_row(12, 20)
    ws.merge_range(12, 1, 12, 5, "À RETENIR POUR LA CAMPAGNE (leviers que la marque et le "
                                 "créateur contrôlent)", s["section"])
    ws.merge_range(12, 7, 12, 9, "CE QUE CES DONNÉES NE DISENT PAS", s["section"])
    recos = [
        "Sélectionner les créateurs sur leur taux d'engagement mesuré sur plusieurs vidéos, pas "
        "sur leur nombre d'abonnés.",
        "Inclure des comptes de moins de 3 M d'abonnés : sous 1 M, "
        f"{fr(t1['vues_par_abonne_med'] / t4['vues_par_abonne_med'], 0)} fois plus de vues par "
        "abonné qu'au-delà de 8 M. Négocier les tarifs en conséquence.",
        "Laisser la main au créateur sur le format : ne pas imposer de hashtags « algorithme », "
        "d'horaire ni de jour de publication dans le brief.",
        "Tester en A/B des formats de 15 à 60 s contre des formats très courts, en suivant le "
        "taux d'engagement.",
    ]
    limites = [
        "Ce qui fait entrer une vidéo dans les tendances : on ne voit que des succès, aucun échec "
        "à comparer.",
        "Ce qui marche aujourd'hui : données d'août 2021 ; algorithme, formats et audiences ont "
        "changé depuis.",
        "Les causes : ce sont des corrélations ; les abonnés, mesurés après coup, peuvent être "
        "une conséquence du succès.",
        "Le retour sur investissement : ni audience (âge, pays), ni ventes, ni coût des "
        "créateurs dans les données.",
    ]
    for i, (a, b) in enumerate(zip(recos, limites)):
        r = 13 + i
        ws.set_row(r, hauteur((a, 123, 10), (b, 65, 10), minimum=24))
        ws.write(r, 1, "✓", C.fmt(bold=True, font_color="#047857", align="center",
                                  valign="top", font_size=12))
        ws.merge_range(r, 2, r, 5, a, s["txt"])
        ws.merge_range(r, 7, r, 9, b, s["txt"])
    ws.set_row(17, 8)
    ws.set_row(18, 18)
    ws.write(18, 1, "→", s["txt_nb"])
    for j, (nom, cell) in enumerate((("Facteurs", "A1"), ("Gros comptes", "A1"),
                                     ("Catégories & TreeMap", "A1"),
                                     ("Méthode & limites", "A1"), ("Données", "A1"))):
        lien(ws, s, 18, [2, 3, 4, 5, 7][j] if j < 4 else 8, nom, cell, nom)
    ws.set_zoom(80)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.print_area(0, 0, 18, 9)


# --------------------------------------------------------------------------- #
# Tableaux croisés dynamiques
# --------------------------------------------------------------------------- #
def source_table(d):
    rows = [tuple(py(v) if not (isinstance(v, str) and v == "") else None for v in row)
            for row in d.itertuples(index=False)]
    return ox.SourceTable(T, COLS, rows, orders={
        "Échantillon": ECH, "Classe de durée": CLASSES, "Tranche abonnés": TRANCHES,
        "Thème": THEMES, "Jour": JOURS, "Créneau UTC": CRENEAUX})


def ecrire_rendu_tcd(C, s, ws, cache, spec, formats):
    table = ox.pivot_values(cache, spec)
    r0, c0 = spec.top_left
    if spec.page:
        ws.write(r0 - 2, c0, spec.page[0], s["th_g"])
        ws.write(r0 - 2, c0 + 1, spec.page[1], s["centre"])
    for i, row in enumerate(table):
        for j, v in enumerate(row):
            if v is None:
                continue
            if i <= (1 if spec.cols else 0) or j == 0:
                ws.write(r0 + i, c0 + j, v, s["th_g"] if j == 0 else s["th"])
            else:
                fm = formats[0] if spec.cols else formats[j - 1]
                ws.write_number(r0 + i, c0 + j, v, s[fm])


def main():
    df = pd.read_csv(WORK / "videos_clean.csv")
    res = json.loads((WORK / "resultats.json").read_text(encoding="utf-8"))
    d = donnees_client(df, res)
    SORTIE.parent.mkdir(exist_ok=True)
    tmp = WORK / "_classeur_sans_tcd.xlsx"
    C = Classeur(str(tmp))
    s = styles(C)
    noms = ["Synthèse", "Facteurs", "Gros comptes", "Catégories & TreeMap",
            "Méthode & limites", "Données"]
    ws = {n: C.wb.add_worksheet(n) for n in noms}
    calc = Calc(d)
    refs = refs_categoriels(d, calc)
    feuille_donnees(C, s, ws["Données"], d)
    feuille_methode(C, s, ws["Méthode & limites"], d, res, calc)
    pos_f = feuille_facteurs(C, s, ws["Facteurs"], d, res, calc, refs)
    pos_g = feuille_gros_comptes(C, s, ws["Gros comptes"], d, res, calc)
    feuille_categories(C, s, ws["Catégories & TreeMap"], d, res, calc)
    feuille_synthese(C, s, ws["Synthèse"], d, res, pos_f, pos_g)

    src = source_table(d)
    cache = ox.Cache(src)
    specs = [
        (ox.PivotSpec("Facteurs", "TCD_Duree", (FA_TCD, 1), rows="Classe de durée",
                      cols="Échantillon", data=[("Top performer", "average",
                                                 "Part de top performers", 9)],
                      row_caption="Classe de durée", col_caption="Échantillon"), ["pct0"]),
        (ox.PivotSpec("Gros comptes", "TCD_Abonnes", (GC_TCD, 1), rows="Tranche abonnés",
                      data=[("Vues", "count", "Vidéos", 3),
                            ("Top performer", "average", "Part de top performers", 9),
                            ("Engagement", "average", "Engagement moyen", 10)],
                      page=("Échantillon", "Humour"), row_caption="Tranche d'abonnés"),
         ["int", "pct0", "pct"]),
        (ox.PivotSpec("Catégories & TreeMap", "TCD_Themes", (CA_TCD, 1), rows="Thème",
                      cols="Échantillon", data=[("Vues", "sum", "Vues (somme)", 0)],
                      row_caption="Sujet", col_caption="Échantillon"), ["m"]),
    ]
    for spec, formats in specs:
        ecrire_rendu_tcd(C, s, ws[spec.sheet], cache, spec, formats)
    ws["Synthèse"].activate()
    ws["Synthèse"].set_first_sheet()
    C.wb.close()

    pkg = ox.Package(tmp)
    vues = specs[2][0]
    vues.data = [("Vues", "sum", "Vues (somme)", ox.numfmt_id(pkg, FMT_M))]
    ox.inject_pivots(pkg, src, [sp for sp, _ in specs])
    ox.set_theme_colors(pkg, [COULEURS_THEMES[t] for t in THEMES] + [NUIT.lstrip("#")])
    pkg.save(SORTIE)
    tmp.unlink()
    print(f"-> {SORTIE.relative_to(ROOT)} ({len(d)} vidéos, {len(res['nommes'])} créateurs "
          f"nommés, {len(specs)} TCD)")


if __name__ == "__main__":
    main()
