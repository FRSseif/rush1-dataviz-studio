"""Phases 3 à 6 — Statistiques, facteurs, objection « gros comptes », catégories,
leaders et données du TreeMap.

Lit work/videos_clean.csv (produit par 01_clean.py) et écrit work/resultats.json,
source unique de tous les chiffres repris dans le classeur et vérifiés par
05_verify_numbers.py.

Définition du succès : taux d'engagement = (likes + commentaires + partages) / vues.
Top performer = quart supérieur (≥ 75e centile) du taux d'engagement de son échantillon.

Usage : .venv/bin/python scripts/02_analyse.py
"""
import json
import random
from pathlib import Path

import numpy as np
import pandas as pd
from scipy import stats

ROOT = Path(__file__).resolve().parent.parent
WORK = ROOT / "work"

PRINCIPAL = "Tendances"
ECHANTILLONS = ["Tendances", "Likées par @tiktok", "Humour"]
CREATEURS = ["Tendances", "Humour"]  # échantillons utilisés pour catégories et leaders
THEMES = ["Animaux", "Famille & enfants", "Humour & sketchs",
          "Lifestyle, musique & sport", "Sujet non déclaré"]
TRANCHES = ["1. < 1 M", "2. 1-3 M", "3. 3-8 M", "4. ≥ 8 M"]
SEUIL = 0.05

# Comptes jamais nommés : plateforme elle-même, ou pseudo révélant une donnée de
# santé (catégorie particulière, art. 9 RGPD) — mots-clés génériques uniquement.
NON_NOMMABLES = r"^tiktok$|tourette|autis|adhd|cancer|diabet|sober|epilep"

# Facteurs que le créateur contrôle. type : bin (oui/non), num (continu), cat.
FACTEURS = [
    ("duree_15", "Durée supérieure à 15 s", "bin", lambda d: d["duree_s"] > 15, ECHANTILLONS),
    ("duree_s", "Durée (secondes)", "num", "duree_s", ECHANTILLONS),
    ("hashtag_generique", "Hashtags génériques (#fyp, #foryou, #viral…)", "bin",
     "hashtag_generique", ECHANTILLONS),
    ("nb_hashtags", "Nombre de hashtags", "num", "nb_hashtags", ECHANTILLONS),
    ("longueur_texte", "Longueur du texte de la légende", "num", "longueur_texte", ECHANTILLONS),
    ("legende_vide", "Légende vide", "bin", "legende_vide", ["Tendances", "Likées par @tiktok"]),
    ("format_collab", "Mention, duo, stitch ou réponse", "bin", "format_collab", ECHANTILLONS),
    ("question_ou_appel", "Question ou appel à l'action", "bin", "question_ou_appel", ECHANTILLONS),
    ("week_end", "Publication le week-end", "bin", "week_end", ECHANTILLONS),
    ("jour_nom", "Jour de publication", "cat", "jour_nom", ECHANTILLONS),
    ("creneau_utc", "Créneau horaire (UTC)", "cat", "creneau_utc", ECHANTILLONS),
    ("theme", "Sujet du contenu (thème du créateur)", "cat", "theme", CREATEURS),
    ("son_original", "Son original (vs musique existante)", "bin", "son_original", ["Humour"]),
    ("stitch_autorise", "Stitch autorisé", "bin", "stitch_autorise", ["Humour"]),
    ("nb_videos_compte", "Vidéos publiées par le compte (régularité)", "num",
     "nb_videos_compte", ["Humour"]),
]


def r6(x):
    return None if x is None or pd.isna(x) else float(round(x, 6))


def spearman(x, y):
    m = x.notna() & y.notna()
    rho, p = stats.spearmanr(x[m], y[m])
    return {"n": int(m.sum()), "rho": r6(rho), "p": r6(p)}


def binaire(d, mask):
    a, b = d.loc[mask, "engagement"], d.loc[~mask, "engagement"]
    res = {"n_oui": int(len(a)), "n_non": int(len(b))}
    if len(a) < 3 or len(b) < 3:
        return res | {"ecart": None, "p": None}
    return res | {
        "med_oui": r6(a.median()), "med_non": r6(b.median()),
        "ecart": r6(a.median() - b.median()),
        "top_oui": r6(d.loc[mask, "top_performer"].mean()),
        "top_non": r6(d.loc[~mask, "top_performer"].mean()),
        "p": r6(stats.mannwhitneyu(a, b, alternative="two-sided").pvalue),
        "vues_oui": r6(d.loc[mask, "vues"].median()),
        "vues_non": r6(d.loc[~mask, "vues"].median()),
    }


def categoriel(d, col):
    groupes = {k: g["engagement"] for k, g in d.groupby(col)}
    meds = {k: r6(v.median()) for k, v in groupes.items()}
    return {"n": int(len(d)), "medianes": meds,
            "ecart": r6(max(meds.values()) - min(meds.values())),
            "p": r6(stats.kruskal(*groupes.values()).pvalue)}


def verdict(effets):
    """effets : liste de (signe, p) — un élément par échantillon testé.

    Confirmé : significatif (p < 0,05) dans au moins 2 échantillons, même sens.
    Contradictoire : significatif dans un sens, et au moins tendanciel (p < 0,10)
    dans le sens opposé ailleurs. Isolé : significatif dans un seul échantillon.
    """
    sig = [s for s, p in effets if p is not None and p < SEUIL]
    tendanciel = [s for s, p in effets if p is not None and p < 0.10]
    if len(sig) >= 2 and len(set(sig)) == 1:
        return "Effet confirmé"
    if sig and len(set(tendanciel) - {0.0}) > 1:
        return "Contradictoire"
    if len(sig) == 1:
        return "Effet isolé"
    return "Aucun effet mesurable"


def analyse_facteurs(df):
    out = []
    for key, label, typ, col, echs in FACTEURS:
        par_ech, effets = {}, []
        for e in echs:
            d = df[df["echantillon"] == e]
            if typ == "bin":
                mask = col(d) if callable(col) else d[col].astype(bool)
                res = binaire(d, mask)
                signe = None if res["ecart"] is None else np.sign(res["ecart"])
            elif typ == "num":
                res = spearman(d[col], d["engagement"]) | {
                    "rho_vues": spearman(d[col], d["vues"])["rho"]}
                signe = np.sign(res["rho"])
            else:
                res = categoriel(d, col)
                signe = 0  # pas de sens pour une variable à plusieurs modalités
            par_ech[e] = res
            if res.get("p") is not None:
                effets.append((float(signe), res["p"]))
        out.append({"cle": key, "libelle": label, "type": typ, "echantillons": echs,
                    "resultats": par_ech, "verdict": verdict(effets)})
    return out


def table_classes(d, col):
    g = d.groupby(col).agg(n=("engagement", "size"), eng_med=("engagement", "median"),
                           vues_med=("vues", "median"), part_top=("top_performer", "mean"))
    return [{"classe": k, **{c: r6(v) for c, v in row.items()}} for k, row in g.iterrows()]


def gros_comptes(df):
    h = df[df["echantillon"] == "Humour"].copy()
    tranches = []
    for t in TRANCHES:
        g = h[h["tranche_abonnes"] == t]
        tranches.append({
            "tranche": t, "videos": int(len(g)), "createurs": int(g["createur"].nunique()),
            "abonnes_med": r6(g["abonnes"].median()), "vues_med": r6(g["vues"].median()),
            "vues_par_abonne_med": r6(g["vues_par_abonne"].median()),
            "eng_med": r6(g["engagement"].median()), "part_top": r6(g["top_performer"].mean()),
        })
    c = h.groupby("createur").agg(ab=("abonnes", "first"), vues=("vues", "median"),
                                  eng=("engagement", "median"))
    # Corrélation partielle (sur les rangs) abonnés ~ engagement à vues égales.
    rk = h[["abonnes", "engagement", "vues"]].rank()
    z = np.c_[np.ones(len(rk)), rk["vues"]]
    res = [rk[v] - z @ np.linalg.lstsq(z, rk[v], rcond=None)[0] for v in ("abonnes", "engagement")]
    partielle = float(np.corrcoef(*res)[0, 1])
    return {
        "tranches": tranches,
        "abonnes_vues": spearman(h["abonnes"], h["vues"]),
        "abonnes_engagement": spearman(h["abonnes"], h["engagement"]),
        "abonnes_vues_par_createur": spearman(c["ab"], c["vues"]),
        "abonnes_engagement_par_createur": spearman(c["ab"], c["eng"]),
        "partielle_abonnes_engagement_a_vues_egales": r6(partielle),
        "likes_cumules_vues": spearman(h["likes_cumules_compte"], h["vues"]),
        "likes_cumules_engagement": spearman(h["likes_cumules_compte"], h["engagement"]),
        "verifie": binaire(h, h["verifie"].astype(bool)),
        "likes_min": r6(h["likes"].min()),
        "vues_engagement": spearman(h["vues"], h["engagement"]),
        "ratio_vues_par_abonne": r6(tranches[0]["vues_par_abonne_med"]
                                    / tranches[-1]["vues_par_abonne_med"]),
    }


def createurs(df):
    """Une ligne par (échantillon, créateur) pour Tendances et Humour."""
    d = df[df["echantillon"].isin(CREATEURS)]
    ref = d.groupby("echantillon").agg(vues_ech=("vues", "median"),
                                       eng_ech=("engagement", "median"))
    c = d.groupby(["echantillon", "createur"]).agg(
        videos=("vues", "size"), vues_tot=("vues", "sum"), vues_med=("vues", "median"),
        eng_med=("engagement", "median"), duree_med=("duree_s", "median"),
        theme=("theme", "first"), tranche=("tranche_abonnes", "first"),
        abonnes=("abonnes", "first"), top=("top_performer", "mean")).reset_index()
    c = c.join(ref, on="echantillon")
    c["indice_eng"] = c["eng_med"] / c["eng_ech"]
    c["portee_suffisante"] = c["vues_med"] >= c["vues_ech"]
    c["nommable"] = ~c["createur"].str.contains(NON_NOMMABLES, case=False, regex=True)
    return c


def fr(x, dec=1):
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",")


def pourquoi(row):
    parts = [f"engagement {fr(row['eng_med'] * 100)} % = {fr(row['indice_eng'], 1)}× "
             f"la médiane de son échantillon",
             f"{fr(row['vues_med'] / 1e6)} M de vues médianes"]
    if row["videos"] > 1:
        parts.append(f"{row['videos']} vidéos dans l'échantillon (régularité)")
    if isinstance(row["tranche"], str):
        parts.append(f"{fr(row['abonnes'] / 1e6)} M d'abonnés")
    return " ; ".join(parts)


def leaders_et_treemap(df):
    c = createurs(df)
    # Un créateur présent dans les deux échantillons : on garde sa meilleure ligne.
    c = c.sort_values("indice_eng", ascending=False)
    leaders = []
    for t in THEMES:
        pool = c[(c["theme"] == t) & c["portee_suffisante"] & c["nommable"]]
        pool = pool.drop_duplicates("createur").head(3)
        for rang, (_, r) in enumerate(pool.iterrows(), 1):
            leaders.append({"theme": t, "rang": rang, "createur": r["createur"],
                            "echantillon": r["echantillon"], "videos": int(r["videos"]),
                            "vues_med": r6(r["vues_med"]), "eng_med": r6(r["eng_med"]),
                            "indice_eng": r6(r["indice_eng"]),
                            "duree_med": r6(r["duree_med"]),
                            "abonnes": r6(r["abonnes"]) if pd.notna(r["abonnes"]) else None,
                            "pourquoi": pourquoi(r)})

    # Portée : vues cumulées par créateur, sans compter deux fois une vidéo commune.
    d = (df[df["echantillon"].isin(CREATEURS)]
         .sort_values("echantillon", key=lambda s: s.map({"Tendances": 0, "Humour": 1}))
         .drop_duplicates("video_id"))
    portee = (d.groupby("createur").agg(vues=("vues", "sum"), videos=("vues", "size"),
                                        theme=("theme", "first")).reset_index()
              .sort_values("vues", ascending=False))
    portee["nommable"] = ~portee["createur"].str.contains(NON_NOMMABLES, case=False, regex=True)
    leaders_portee = []
    for t in THEMES:
        r = portee[(portee["theme"] == t) & portee["nommable"]].iloc[0]
        leaders_portee.append({"theme": t, "createur": r["createur"],
                               "vues": r6(r["vues"]), "videos": int(r["videos"])})

    nommes = sorted({x["createur"] for x in leaders} | {x["createur"] for x in leaders_portee})
    tm = []
    for t in THEMES:
        p = portee[portee["theme"] == t]
        for _, r in p[p["createur"].isin(nommes)].iterrows():
            tm.append({"theme": t, "tuile": "@" + r["createur"], "vues": r6(r["vues"])})
        autres = p[~p["createur"].isin(nommes)]
        if len(autres):
            tm.append({"theme": t, "tuile": f"Autres créateurs ({len(autres)})",
                       "vues": r6(autres["vues"].sum())})

    tot = portee["vues"].sum()
    top10 = portee.head(max(1, round(len(portee) * 0.10)))
    themes = []
    for t in THEMES:
        p = portee[portee["theme"] == t]
        v = d[d["createur"].isin(p["createur"])]
        themes.append({"theme": t, "createurs": int(len(p)), "videos": int(len(v)),
                       "vues": r6(p["vues"].sum()), "part_vues": r6(p["vues"].sum() / tot),
                       "eng_med": r6(v["engagement"].median()),
                       "part_top": r6(v["top_performer"].mean())})
    return {
        "leaders": leaders, "leaders_portee": leaders_portee, "nommes": nommes,
        "treemap": tm, "themes": themes,
        "concentration": {"createurs": int(len(portee)), "top10_n": int(len(top10)),
                          "top10_part_vues": r6(top10["vues"].sum() / tot),
                          "vues_total": r6(tot)},
    }


def pseudonymes(df, nommes):
    """Code stable « Créateur 001 » pour tout créateur non nommé (ordre aléatoire fixe)."""
    autres = sorted(set(df["createur"]) - set(nommes))
    random.Random(2021).shuffle(autres)
    return {c: f"Créateur {i:03d}" for i, c in enumerate(autres, 1)}


def main():
    df = pd.read_csv(WORK / "videos_clean.csv")
    for c in ("hashtag_generique", "legende_vide", "format_collab", "question_ou_appel",
              "week_end", "top_performer"):
        df[c] = df[c].astype(bool)

    ech = {}
    for e in ECHANTILLONS:
        d = df[df["echantillon"] == e]
        ech[e] = {"videos": int(len(d)), "createurs": int(d["createur"].nunique()),
                  "vues_med": r6(d["vues"].median()), "eng_med": r6(d["engagement"].median()),
                  "seuil_top": r6(d["engagement"].quantile(0.75)),
                  "duree_med": r6(d["duree_s"].median()),
                  "part_15s_ou_moins": r6((d["duree_s"] <= 15).mean()),
                  "vues_engagement": spearman(d["vues"], d["engagement"]),
                  "age_vues": spearman(d["age_jours"], d["vues"]),
                  "age_engagement": spearman(d["age_jours"], d["engagement"]),
                  "publication_min": d["publication"].min()[:10],
                  "publication_max": d["publication"].max()[:10]}

    cat = leaders_et_treemap(df)
    res = {
        "echantillons": ech,
        "facteurs": analyse_facteurs(df),
        "duree_classes": {e: table_classes(df[df["echantillon"] == e], "classe_duree")
                          for e in ECHANTILLONS},
        "gros_comptes": gros_comptes(df),
        **cat,
        "pseudonymes": pseudonymes(df, cat["nommes"]),
    }
    (WORK / "resultats.json").write_text(json.dumps(res, ensure_ascii=False, indent=1))

    # Résumé lisible.
    for f in res["facteurs"]:
        cells = []
        for e, r in f["resultats"].items():
            if f["type"] == "bin":
                cells.append(f"{e[:4]} {'' if r['ecart'] is None else fr(r['ecart'] * 100)}pt "
                             f"p={r['p']}")
            elif f["type"] == "num":
                cells.append(f"{e[:4]} ρ={fr(r['rho'], 2)} p={fr(r['p'], 3)}")
            else:
                cells.append(f"{e[:4]} écart={fr(r['ecart'] * 100)}pt p={fr(r['p'], 3)}")
        print(f"{f['libelle'][:42]:42s} | {f['verdict']:22s} | " + " | ".join(cells))
    g = res["gros_comptes"]
    print("\nGros comptes :", g["abonnes_vues"], g["abonnes_engagement"],
          "partielle", g["partielle_abonnes_engagement_a_vues_egales"])
    for t in g["tranches"]:
        print("  ", t)
    print("\nThèmes :")
    for t in res["themes"]:
        print("  ", t)
    print("\nLeaders :")
    for x in res["leaders"]:
        print(f"   {x['theme'][:14]:14s} {x['rang']} @{x['createur']:22s} {x['echantillon'][:4]} "
              f"{x['pourquoi']}")
    print("Leaders portée :", [(x["theme"][:10], x["createur"], fr(x["vues"] / 1e6))
                               for x in res["leaders_portee"]])
    print("Nommés :", len(res["nommes"]), "| concentration :", res["concentration"])
    print("Treemap :", len(res["treemap"]), "tuiles")


if __name__ == "__main__":
    main()
