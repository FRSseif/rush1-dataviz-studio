"""Étape finale, après l'insertion manuelle du TreeMap dans Excel (voir README).

Enregistrer un classeur dans Excel y ajoute des informations qui n'ont rien à faire
dans la version client : nom de la personne qui a enregistré, chemin local du
fichier, étiquette de confidentialité de l'organisation (propriétés MSIP), nom de
qui a actualisé les TCD, paramètres d'imprimante. Ce script les retire, remet la
Synthèse en feuille active (cellule A1), puis relance l'audit et la vérification.

Usage : .venv/bin/python scripts/06_finaliser.py [chemin.xlsx]
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


def main(path):
    path = Path(path)
    nettoyer(path)
    print(f"Nettoyé : {path.relative_to(ROOT) if path.is_relative_to(ROOT) else path}\n")
    audit = importlib.import_module("04_audit_client_xlsx")
    verif = importlib.import_module("05_verify_numbers")
    ok_audit = audit.main(path)
    print()
    ok_verif = verif.main(path, avec_excel=False)
    return ok_audit and ok_verif


if __name__ == "__main__":
    sys.exit(0 if main(sys.argv[1] if len(sys.argv) > 1 else CLIENT) else 1)
