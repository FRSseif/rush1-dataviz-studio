"""Étape finale, après l'insertion manuelle du TreeMap dans Excel (voir README).

Enregistrer un classeur dans Excel y ajoute des informations qui n'ont rien à faire
dans la version client : nom de la personne qui a enregistré, chemin local du
fichier, étiquette de confidentialité de l'organisation (propriétés MSIP), nom de
qui a actualisé les TCD, paramètres d'imprimante. Ce script les retire, remet la
Synthèse en feuille active (cellule A1), puis contrôle le classeur.

Ces contrôles n'utilisent que la bibliothèque standard et fonctionnent sans les
données brutes. Si les CSV sont dans data_raw/ et que la chaîne a tourné (work/),
le script relance aussi l'audit complet et la vérification des chiffres.

Usage : python3 scripts/06_finaliser.py [chemin.xlsx]
"""
import importlib
import re
import shutil
import sys
from pathlib import Path

import ooxml_extras as ox

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
CLIENT = ROOT / "deliverable" / "Dataviz_Studio_TikTok_Client.xlsx"
# Entrées de l'audit complet (04) et de la vérification des chiffres (05).
ENTREES = [ROOT / "data_raw" / f for f in (
    "trending_videos.csv", "tiktok_collected_liked_videos.csv", "tiktok_collected_videos.csv",
    "tiktok_funny_hashtag_videos.csv", "trending_authors.csv")] + [ROOT / "work" / "resultats.json"]


def retirer_partie(pkg, part, rels_part):
    """Supprime une partie, sa relation et son type de contenu."""
    if part not in pkg.parts:
        return
    del pkg.parts[part]
    pkg.order.remove(part)
    cible = part.split("/", 1)[1] if rels_part == "_rels/.rels" else part.rsplit("/", 1)[1]
    rels = pkg.text(rels_part)
    rels = re.sub(rf'<Relationship [^>]*Target="[^"]*{re.escape(cible)}"[^>]*/>', "", rels)
    pkg.put(rels_part, rels)
    ct = re.sub(rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "",
                pkg.text("[Content_Types].xml"))
    pkg.put("[Content_Types].xml", ct)


def nettoyer(path):
    pkg = ox.Package(path)
    core = pkg.text("docProps/core.xml")
    core = re.sub(r"<dc:creator>[^<]*</dc:creator>", "<dc:creator></dc:creator>", core)
    core = re.sub(r"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>|<cp:lastModifiedBy/>", "",
                  core)
    core = re.sub(r"<cp:lastPrinted>[^<]*</cp:lastPrinted>", "", core)
    pkg.put("docProps/core.xml", core)
    app = re.sub(r"<(Company|Manager)>[^<]*</\1>", r"<\1></\1>", pkg.text("docProps/app.xml"))
    pkg.put("docProps/app.xml", app)

    retirer_partie(pkg, "docProps/custom.xml", "_rels/.rels")

    wb = pkg.text("xl/workbook.xml")
    wb = re.sub(r"<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<x15ac:absPath[^>]*/>\s*"
                r"</mc:Choice>\s*</mc:AlternateContent>", "", wb)
    wb = re.sub(r'(<workbookView[^>]*?)\s+(activeTab|firstSheet)="\d+"', r"\1", wb)
    wb = re.sub(r'(<workbookView[^>]*?)\s+(activeTab|firstSheet)="\d+"', r"\1", wb)
    pkg.put("xl/workbook.xml", wb)

    synthese = pkg.sheet_part("Synthèse")
    for part in [n for n in pkg.order if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]:
        x = re.sub(r'\s+tabSelected="1"', "", pkg.text(part))
        if part == synthese:
            x = x.replace("<sheetView ", '<sheetView tabSelected="1" ', 1)
            x = re.sub(r"<selection [^>]*/>", '<selection activeCell="A1" sqref="A1"/>', x)
        # Paramètres d'imprimante enregistrés par Excel : retirés avec leur relation.
        x = re.sub(r'(<pageSetup[^>]*?)\s+r:id="rId\d+"', r"\1", x)
        pkg.put(part, x)
        rels = ox.rels_name(part)
        if rels in pkg.parts:
            pkg.put(rels, re.sub(r'<Relationship [^>]*printerSettings[^>]*/>', "",
                                 pkg.text(rels)))
    for part in [n for n in list(pkg.order) if n.startswith("xl/printerSettings/")]:
        del pkg.parts[part]
        pkg.order.remove(part)
        pkg.put("[Content_Types].xml", re.sub(
            rf'<Override PartName="/{re.escape(part)}"[^>]*/>', "",
            pkg.text("[Content_Types].xml")))

    for part in [n for n in pkg.order if n.startswith("xl/pivotCache/pivotCacheDefinition")
                 and n.endswith(".xml")]:
        pkg.put(part, re.sub(r'\s+refreshedBy="[^"]*"', "", pkg.text(part)))

    tmp = path.with_suffix(".tmp.xlsx")
    pkg.save(tmp)
    shutil.move(tmp, path)


def controler(path):
    """Contrôles du fichier lui-même, sans les données brutes."""
    pkg = ox.Package(path)
    texte = {n: pkg.parts[n].decode("utf-8", errors="replace") for n in pkg.order
             if n.endswith((".xml", ".rels"))}
    tout = "\n".join(texte.values())
    wb = texte["xl/workbook.xml"]
    return [
        ("TreeMap natif présent", any(n.startswith("xl/charts/chartEx")
                                      and 'layoutId="treemap"' in x for n, x in texte.items())),
        ("Tableaux croisés dynamiques présents",
         sum(bool(re.match(r"xl/pivotTables/pivotTable\d+\.xml$", n)) for n in pkg.order) >= 2),
        ("Ni auteur ni dernier modificateur",
         not re.search(r"<(dc:creator|cp:lastModifiedBy)>[^<]+<", texte["docProps/core.xml"])),
        ("Aucun chemin local", not re.search(r"/Users/|[A-Z]:\\\\|absPath", tout)),
        ("Aucune propriété personnalisée", "docProps/custom.xml" not in pkg.parts),
        ("Aucune trace de qui a actualisé les TCD", "refreshedBy" not in tout),
        ("Aucune feuille masquée",
         not re.search(r'<sheet [^>]*state="(hidden|veryHidden)"', wb)),
        ("Aucune macro ni lien externe",
         not any("vbaProject" in n or "externalLink" in n for n in pkg.order)),
        ("La Synthèse s'ouvre en premier",
         re.findall(r'<sheet [^>]*name="([^"]+)"', wb)[:1] == ["Synthèse"]
         and not re.search(r'activeTab="[1-9]', wb)),
    ]


def main(path):
    path = Path(path)
    if (path.parent / ("~$" + path.name)).exists():
        print("Le classeur est encore ouvert dans Excel : le fermer, puis relancer.")
        return False
    nettoyer(path)
    print(f"Nettoyé : {path.relative_to(ROOT) if path.is_relative_to(ROOT) else path}\n")
    resultats = controler(path)
    for nom, ok in resultats:
        print(f"{'PASS' if ok else 'FAIL'}  {nom}")
    ok_fichier = all(ok for _, ok in resultats)
    if not all(p.exists() for p in ENTREES):
        print("\nAudit complet et vérification des chiffres non relancés : ils demandent les "
              "CSV dans data_raw/ et les fichiers de work/ (voir README, « Reproduire »).")
        return ok_fichier
    print()
    audit = importlib.import_module("04_audit_client_xlsx")
    verif = importlib.import_module("05_verify_numbers")
    ok_audit = audit.main(path)
    print()
    ok_verif = verif.main(path, avec_excel=False)
    return ok_fichier and ok_audit and ok_verif


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1] if len(sys.argv) > 1 else CLIENT) else 1)
