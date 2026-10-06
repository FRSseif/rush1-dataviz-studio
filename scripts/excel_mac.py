"""Validation dans Microsoft Excel pour Mac, piloté par AppleScript.

Excel est « sandboxé » : il ne lit et n'écrit sans demande d'autorisation que dans
son conteneur (~/Library/Containers/com.microsoft.Excel/Data). On y copie le
classeur, Excel l'ouvre, recalcule tout, compte les objets natifs (TCD,
graphiques), puis enregistre une copie « aller-retour » que l'on relit en Python.
Si Excel devait réparer le fichier, les objets réparés disparaîtraient de cette
copie : c'est ce que vérifient les scripts d'audit et de vérification.
"""
import shutil
import subprocess
from pathlib import Path

CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/rush1"


def run(script, timeout=180):
    p = subprocess.run(["osascript", "-e", script], capture_output=True, text=True,
                       timeout=timeout)
    if p.returncode != 0:
        raise RuntimeError(p.stderr.strip())
    return p.stdout.strip()


def roundtrip(src: Path, dest: Path, sheets_with_pivots=()):
    """Ouvre `src` dans Excel, recalcule, enregistre une copie dans `dest`.

    Renvoie un dict : nombre de feuilles, feuille active, TCD par feuille.
    """
    CONTAINER.mkdir(parents=True, exist_ok=True)
    tmp_in, tmp_out = CONTAINER / "controle.xlsx", CONTAINER / "controle_aller_retour.xlsx"
    for f in (tmp_in, tmp_out):
        f.unlink(missing_ok=True)
    shutil.copy(src, tmp_in)
    pivots = "".join(
        f'set end of out to "{s}=" & (count of pivot tables of worksheet "{s}" of wb)\n'
        for s in sheets_with_pivots)
    script = f'''
with timeout of 150 seconds
tell application "Microsoft Excel"
    set display alerts to false
    open POSIX file "{tmp_in}"
    set wb to active workbook
    set out to {{}}
    set end of out to "feuilles=" & (count of worksheets of wb)
    set end of out to "active=" & (name of active sheet of wb)
    {pivots}
    calculate full
    save workbook as wb filename "{tmp_out}" file format Excel XML file format
    close wb saving no
end tell
end timeout
set AppleScript's text item delimiters to "|"
return out as text
'''
    out = run(script)
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(tmp_out, dest)
    for f in (tmp_in, tmp_out):
        f.unlink(missing_ok=True)
    return dict(item.split("=", 1) for item in out.split("|"))
