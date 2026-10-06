"""Phase 8 — Audit bloquant du classeur client, sur le XML brut de l'archive.

Vérifie que rien de ce qui a été exclu de la version client n'est récupérable
depuis le fichier, que le classeur est autonome et qu'il s'ouvre sur la Synthèse.
Les chaînes sensibles recherchées sont construites à partir des CSV de data_raw/
(colonnes exclues + pseudos des créateurs non retenus).

Usage : .venv/bin/python scripts/04_audit_client_xlsx.py [chemin.xlsx]
Code retour : 0 si tous les contrôles sont PASS, 1 sinon.
"""
import html
import importlib
import json
import re
import sys
import zipfile
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
RAW = ROOT / "data_raw"
CLIENT = ROOT / "deliverable" / "Dataviz_Studio_TikTok_Client.xlsx"
build = importlib.import_module("03_build_workbook")

AUTORISEES = set(build.COLS)
NOMS_AUTORISES = {"_xlnm.Print_Area", "_xlnm.Print_Titles"}
FEUILLES = ["Synthèse", "Facteurs", "Gros comptes", "Catégories & TreeMap",
            "Méthode & limites", "Données"]
ERREURS_EXCEL = ("#REF!", "#N/A", "#DIV/0!", "#VALUE!", "#NUM!", "#NAME?", "#NULL!")
# Pseudos qui sont aussi des mots du texte du classeur (comptes d'organisations).
LISTE_BLANCHE = {"tiktok"}
# Textes « libres » génériques qui sont aussi du vocabulaire du classeur : le titre
# musical par défaut « son original » est aussi le nom d'une colonne.
TEXTES_GENERIQUES = {"son original"}


class Audit:
    def __init__(self):
        self.lignes = []

    def check(self, nom, ok, detail=""):
        self.lignes.append((nom, bool(ok), detail))

    def rapport(self):
        largeur = max(len(n) for n, _, _ in self.lignes)
        for nom, ok, detail in self.lignes:
            print(f"{'PASS' if ok else 'FAIL'}  {nom:<{largeur}}  {detail}")
        n_ok = sum(ok for _, ok, _ in self.lignes)
        print(f"\n{n_ok}/{len(self.lignes)} contrôles PASS")
        return all(ok for _, ok, _ in self.lignes)


def textes(z):
    """Contenu texte décodé de toutes les parties XML de l'archive."""
    out = {}
    for n in z.namelist():
        if n.endswith((".xml", ".rels", ".vml")):
            out[n] = html.unescape(z.read(n).decode("utf-8", errors="replace"))
    return out


def sensibles(nommes):
    """Chaînes issues des colonnes exclues des CSV bruts, à ne jamais retrouver."""
    tv = pd.read_csv(RAW / "trending_videos.csv")
    lk = pd.read_csv(RAW / "tiktok_collected_liked_videos.csv")
    of = pd.read_csv(RAW / "tiktok_collected_videos.csv")
    fh = pd.read_csv(RAW / "tiktok_funny_hashtag_videos.csv")
    ta = pd.read_csv(RAW / "trending_authors.csv")
    ids, libres, pseudos = set(), set(), set()
    for d in (tv, lk, of):
        ids |= set(d["user_id"].astype(str)) | set(d["video_id"].astype(str))
        libres |= set(d["video_desc"].dropna())
        pseudos |= set(d["user_name"])
    ids |= set(fh["author_id"].astype(str)) | set(fh["video_id"].astype(str)) | set(
        fh["music_id"].astype(str))
    libres |= set(fh["video_desc"].dropna()) | set(fh["author_signature"].dropna())
    libres |= set(fh["author_nickname"].dropna()) | set(fh["music_title"].dropna())
    libres |= set(fh["music_authorName"].dropna())
    libres |= set(ta["Signature"].dropna()) | set(ta["Author Nickname"].dropna())
    pseudos |= set(fh["author_uniqueId"]) | set(ta["Author Unique ID"])
    # Textes libres : on cherche les chaînes assez longues pour être discriminantes.
    libres = {x.strip() for x in libres if len(x.strip()) >= 12
              and x.strip().lower() not in TEXTES_GENERIQUES}
    # Un nom affiché identique au pseudo d'un créateur nommé n'apporte rien de plus.
    libres -= set(nommes)
    pseudos = {p for p in pseudos if p.lower() not in LISTE_BLANCHE} - set(nommes)
    return ids, libres, pseudos


def main(path):
    a = Audit()
    path = Path(path)
    a.check("Extension .xlsx (sans macros)", path.suffix == ".xlsx", path.name)
    z = zipfile.ZipFile(path)
    noms = z.namelist()
    t = textes(z)
    wb_xml = t["xl/workbook.xml"]
    ctypes = t["[Content_Types].xml"]

    # 1. Autonomie : macros, liens externes, connexions, objets incorporés.
    a.check("Aucune macro (vbaProject.bin)", not any("vbaProject" in n for n in noms)
            and "macroEnabled" not in ctypes)
    externes = [n for n in noms if n.startswith("xl/externalLinks/")]
    rel_ext = [n for n, x in t.items() if n.endswith(".rels") and 'TargetMode="External"' in x]
    a.check("Aucun lien vers un autre fichier", not externes and not rel_ext,
            ", ".join(externes + rel_ext))
    a.check("Aucune connexion de données", "xl/connections.xml" not in noms
            and not any("queryTable" in n for n in noms))
    a.check("Aucun objet incorporé", not any(n.startswith(("xl/embeddings/", "xl/activeX/"))
                                             for n in noms))
    a.check("Aucun commentaire ni note", not any(re.search(
        r"(comments|threadedComments|persons)/?.*\.xml$|vmlDrawing", n) for n in noms))

    # 2. Rien de masqué.
    etats = re.findall(r'<sheet [^>]*state="(hidden|veryHidden)"', wb_xml)
    a.check("Aucune feuille masquée ou très masquée", not etats, str(etats))
    caches = []
    for n, x in t.items():
        if n.startswith("xl/worksheets/sheet"):
            if re.search(r'<(row|col) [^>]*hidden="1"', x):
                caches.append(n)
    a.check("Aucune ligne ou colonne masquée", not caches, ", ".join(caches))
    feuilles = re.findall(r'<sheet name="([^"]+)"', wb_xml)
    a.check("Onglets attendus, dans l'ordre", feuilles == FEUILLES, " | ".join(feuilles))

    # 3. Noms définis.
    defs = re.findall(r'<definedName name="([^"]+)"', wb_xml)
    a.check("Aucun nom défini résiduel", set(defs) <= NOMS_AUTORISES,
            f"{len(defs)} nom(s) : {sorted(set(defs))}")

    # 4. Caches de TCD : uniquement des champs autorisés.
    defs_tcd = [n for n in noms if re.match(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", n)]
    champs = set()
    for n in defs_tcd:
        champs |= set(re.findall(r'<cacheField name="([^"]+)"', t[n]))
    a.check("Caches de TCD : champs autorisés seulement", champs <= AUTORISEES,
            f"{len(defs_tcd)} cache(s), {len(champs)} champs"
            + (f", interdits : {sorted(champs - AUTORISEES)}" if champs - AUTORISEES else ""))
    n_tcd = len([n for n in noms if re.match(r"xl/pivotTables/pivotTable\d+\.xml", n)])
    a.check("Tableaux croisés dynamiques présents (≥ 2)", n_tcd >= 2, f"{n_tcd} TCD")
    a.check("TreeMap natif présent (graphique « Compartimentage »)",
            any(n.startswith("xl/charts/chartEx") for n in noms)
            and "layoutId=\"treemap\"" in "".join(v for k, v in t.items() if "chartEx" in k),
            "à insérer dans Excel puis lancer 06_finaliser.py (voir README)")

    # 5. Données personnelles exclues : recherche dans TOUT le contenu de l'archive.
    res = json.loads((ROOT / "work" / "resultats.json").read_text(encoding="utf-8"))
    ids, libres, pseudos = sensibles(res["nommes"])
    corpus = "\n".join(t.values())
    corpus_l = corpus.lower()
    trouves = [i for i in ids if len(i) >= 6 and re.search(rf"(?<!\d){i}(?!\d)", corpus)]
    a.check("Aucun identifiant (utilisateur, vidéo, musique)", not trouves,
            f"{len(ids)} identifiants cherchés" + (f", trouvés : {trouves[:3]}" if trouves else ""))
    trouves = [x for x in libres if x.lower() in corpus_l]
    a.check("Aucune bio, légende, nom affiché ni titre de musique", not trouves,
            f"{len(libres)} textes cherchés" + (f", trouvés : {trouves[:2]}" if trouves else ""))
    trouves = sorted(p for p in pseudos if re.search(
        rf"(?<![\w.]){re.escape(p.lower())}(?![\w])", corpus_l))
    a.check("Aucun pseudo de créateur non retenu", not trouves,
            f"{len(pseudos)} pseudos cherchés" + (f", trouvés : {trouves[:5]}" if trouves else ""))
    urls = re.findall(r"https?://[^\s\"<]+", corpus)
    urls = [u for u in urls if not re.match(r"https?://schemas\.(openxmlformats|microsoft)\.|"
                                            r"https?://purl\.org/|https?://www\.w3\.org/", u)]
    a.check("Aucune URL (vidéo, avatar, lien de profil)", not urls, ", ".join(urls[:3]))
    mails = re.findall(r"[\w.+-]+@[\w-]+\.[a-z]{2,}", corpus)
    a.check("Aucune adresse e-mail", not mails, ", ".join(mails[:3]))

    # 6. Métadonnées.
    core = t.get("docProps/core.xml", "")
    app = t.get("docProps/app.xml", "")
    auteurs = re.findall(r"<(dc:creator|cp:lastModifiedBy)>([^<]+)<", core)
    a.check("Métadonnées : ni auteur ni dernier modificateur", not auteurs, str(auteurs))
    societe = re.findall(r"<(Company|Manager)>([^<]+)<", app)
    a.check("Métadonnées : ni société ni responsable", not societe, str(societe))
    chemins = re.findall(r"(/Users/[^\"<]+|[A-Z]:\\\\[^\"<]+|absPath[^>]*)", corpus)
    a.check("Aucun chemin local", not chemins, str(chemins[:2]))
    a.check("Aucune trace de qui a actualisé les TCD", "refreshedBy" not in corpus)
    a.check("Aucune propriété personnalisée", "docProps/custom.xml" not in noms)

    # 7. Ouverture, erreurs, feuille active.
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ouvert = True
    except Exception as e:  # noqa: BLE001
        ouvert, wb = False, None
        a.check("Ouverture avec openpyxl", False, str(e))
    if ouvert:
        a.check("Ouverture avec openpyxl", True)
        err = [(ws.title, c.coordinate, c.value) for ws in wb.worksheets for row in
               ws.iter_rows() for c in row if isinstance(c.value, str)
               and c.value.strip() in ERREURS_EXCEL]
        a.check("Aucune cellule en erreur (#REF!, #N/A, #DIV/0!…)", not err, str(err[:3]))
        a.check("Feuille active à l'ouverture = Synthèse", wb.active.title == "Synthèse",
                wb.active.title)
        sel = wb["Synthèse"].sheet_view.selection[0].activeCell if wb[
            "Synthèse"].sheet_view.selection else "A1"
        a.check("Cellule A1 sélectionnée sur la Synthèse", sel in (None, "A1"), str(sel))
    autres = [p.name for p in (ROOT / "deliverable").iterdir() if p.name != path.name
              and not p.name.startswith(".")] if path.parent == ROOT / "deliverable" else []
    a.check("deliverable/ ne contient que le classeur client", not autres, ", ".join(autres))
    return a.rapport()


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1] if len(sys.argv) > 1 else CLIENT) else 1)
