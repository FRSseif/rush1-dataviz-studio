"""Ajoute à un .xlsx produit par XlsxWriter des tableaux croisés dynamiques (TCD)
natifs, qu'aucune bibliothèque Python ne sait écrire, et règle les couleurs du
thème du classeur (utilisées par les TCD et par le TreeMap natif d'Excel).

Le fichier est post-traité comme une archive OOXML : on ajoute les parties XML,
leurs relations et leurs types de contenu. Les TCD partagent un cache unique
construit sur le tableau structuré source ; `refreshOnLoad` demande à Excel de
les recalculer à l'ouverture. Les valeurs déjà affichées dans les cellules du
TCD sont écrites par le script de construction, pour un rendu correct même avant
actualisation (mode protégé, aperçus).
"""
import re
import zipfile
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_PIVOT_TABLE = NS_R + "/pivotTable"
REL_PIVOT_CACHE = NS_R + "/pivotCacheDefinition"
REL_PIVOT_RECORDS = NS_R + "/pivotCacheRecords"
CT = {
    "pivotTable": "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
    "pivotCacheDefinition":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
    "pivotCacheRecords":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
}


def esc(s):
    return escape(str(s), {'"': "&quot;"})


# --------------------------------------------------------------------------- #
# Spécifications
# --------------------------------------------------------------------------- #
@dataclass
class PivotSpec:
    sheet: str                 # feuille de destination
    name: str                  # nom du TCD
    top_left: tuple            # (ligne, colonne) 0-based de la zone du TCD (hors filtre)
    rows: str                  # champ en lignes
    data: list                 # [(champ, fonction 'sum'|'count'|'average', libellé, numFmtId)]
    cols: str = None           # champ en colonnes (facultatif, un seul champ de données)
    page: tuple = None         # (champ, valeur sélectionnée) : filtre de rapport
    row_caption: str = None
    col_caption: str = None


@dataclass
class SourceTable:
    """Données du tableau structuré source, dans l'ordre exact des colonnes."""
    name: str
    columns: list
    rows: list                 # liste de tuples (str | float | int | None)
    orders: dict = field(default_factory=dict)  # ordre d'affichage des modalités


# --------------------------------------------------------------------------- #
# Utilitaires d'archive
# --------------------------------------------------------------------------- #
def col_letter(c):
    s = ""
    c += 1
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def cell_ref(r, c):
    return f"{col_letter(c)}{r + 1}"


class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.order = z.namelist()
            self.parts = {n: z.read(n) for n in self.order}

    def text(self, name):
        return self.parts[name].decode("utf-8")

    def put(self, name, content):
        if isinstance(content, str):
            content = content.encode("utf-8")
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = content

    def save(self, path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
            for n in self.order:
                z.writestr(n, self.parts[n])

    def add_override(self, part, ctype):
        ct = self.text("[Content_Types].xml")
        if f'PartName="/{part}"' not in ct:
            ct = ct.replace("</Types>",
                            f'<Override PartName="/{part}" ContentType="{ctype}"/></Types>')
            self.put("[Content_Types].xml", ct)

    def add_rel(self, rels_part, rtype, target):
        if rels_part in self.parts:
            xml = self.text(rels_part)
        else:
            xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                   'relationships"></Relationships>')
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', xml)]
        rid = f"rId{max(ids, default=0) + 1}"
        xml = xml.replace("</Relationships>",
                          f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
                          "</Relationships>")
        self.put(rels_part, xml)
        return rid

    def sheet_part(self, sheet_name):
        wb = self.text("xl/workbook.xml")
        m = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="(rId\d+)"' % re.escape(
            esc(sheet_name)), wb)
        if not m:
            raise KeyError(sheet_name)
        rels = self.text("xl/_rels/workbook.xml.rels")
        target = re.search(r'Id="%s" Type="[^"]+" Target="([^"]+)"' % m.group(1), rels)
        if not target:
            target = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % m.group(1), rels)
        return "xl/" + target.group(1).lstrip("/").removeprefix("xl/")


def rels_name(part):
    d, f = part.rsplit("/", 1)
    return f"{d}/_rels/{f}.rels"


# --------------------------------------------------------------------------- #
# Cache de TCD
# --------------------------------------------------------------------------- #
def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _fmt_num(v):
    if float(v).is_integer():
        return str(int(v))
    return repr(float(v))


class Cache:
    def __init__(self, src: SourceTable):
        self.src = src
        self.fields = []
        for j, name in enumerate(src.columns):
            vals = [r[j] for r in src.rows]
            present = [v for v in vals if v is not None and v != ""]
            blank = len(present) < len(vals)
            if present and all(_num(v) for v in present):
                self.fields.append({"name": name, "kind": "num", "blank": blank,
                                    "min": min(present), "max": max(present),
                                    "int": all(float(v).is_integer() for v in present)})
            else:
                uniq = list(dict.fromkeys(str(v) for v in present))
                order = src.orders.get(name)
                if order:
                    uniq = [u for u in order if u in uniq] + [u for u in uniq if u not in order]
                self.fields.append({"name": name, "kind": "str", "blank": blank,
                                    "items": uniq})

    def index(self, name):
        return self.src.columns.index(name)

    def items(self, name):
        """Modalités (y compris la modalité vide en dernier) d'un champ texte."""
        f = self.fields[self.index(name)]
        return f["items"] + ([None] if f["blank"] else [])

    def definition_xml(self):
        out = []
        for f in self.fields:
            if f["kind"] == "num":
                attrs = ['containsString="0"']
                if f["blank"]:
                    attrs.append('containsBlank="1"')
                else:
                    attrs.insert(0, 'containsSemiMixedTypes="0"')
                attrs.append('containsNumber="1"')
                if f["int"]:
                    attrs.append('containsInteger="1"')
                attrs.append(f'minValue="{_fmt_num(f["min"])}" maxValue="{_fmt_num(f["max"])}"')
                out.append(f'<cacheField name="{esc(f["name"])}" numFmtId="0">'
                           f'<sharedItems {" ".join(attrs)}/></cacheField>')
            else:
                items = "".join(f'<s v="{esc(v)}"/>' for v in f["items"])
                n = len(f["items"]) + (1 if f["blank"] else 0)
                blank = ' containsBlank="1"' if f["blank"] else ""
                if f["blank"]:
                    items += "<m/>"
                out.append(f'<cacheField name="{esc(f["name"])}" numFmtId="0">'
                           f'<sharedItems{blank} count="{n}">{items}</sharedItems>'
                           f'</cacheField>')
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotCacheDefinition xmlns="{NS_MAIN}" xmlns:r="{NS_R}" r:id="rId1" '
            f'refreshOnLoad="1" createdVersion="6" refreshedVersion="6" '
            f'minRefreshableVersion="3" recordCount="{len(self.src.rows)}">'
            f'<cacheSource type="worksheet"><worksheetSource name="{esc(self.src.name)}"/>'
            f'</cacheSource><cacheFields count="{len(self.fields)}">{"".join(out)}'
            '</cacheFields></pivotCacheDefinition>')

    def records_xml(self):
        lookup = [({v: i for i, v in enumerate(f["items"])} if f["kind"] == "str" else None)
                  for f in self.fields]
        recs = []
        for row in self.src.rows:
            cells = []
            for f, lk, v in zip(self.fields, lookup, row):
                if v is None or v == "":
                    cells.append(f'<x v="{len(f["items"])}"/>' if f["kind"] == "str"
                                 else "<m/>")
                elif f["kind"] == "num":
                    cells.append(f'<n v="{_fmt_num(v)}"/>')
                else:
                    cells.append(f'<x v="{lk[str(v)]}"/>')
            recs.append("<r>" + "".join(cells) + "</r>")
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<pivotCacheRecords xmlns="{NS_MAIN}" xmlns:r="{NS_R}" '
                f'count="{len(recs)}">{"".join(recs)}</pivotCacheRecords>')


def pivot_shape(cache: Cache, spec: PivotSpec):
    """Lignes/colonnes visibles du TCD (modalités ayant des données après filtre)."""
    rows = cache.src.rows
    if spec.page:
        j = cache.index(spec.page[0])
        rows = [r for r in rows if r[j] == spec.page[1]]
    jr = cache.index(spec.rows)
    present_r = {r[jr] if r[jr] not in ("", None) else None for r in rows}
    row_items = [i for i, v in enumerate(cache.items(spec.rows)) if v in present_r]
    col_items = []
    if spec.cols:
        jc = cache.index(spec.cols)
        present_c = {r[jc] if r[jc] not in ("", None) else None for r in rows}
        col_items = [i for i, v in enumerate(cache.items(spec.cols)) if v in present_c]
    return rows, row_items, col_items


def pivot_values(cache: Cache, spec: PivotSpec):
    """Calcule le rendu du TCD (pour l'écrire dans les cellules) : liste de lignes."""
    rows, row_items, col_items = pivot_shape(cache, spec)
    jr = cache.index(spec.rows)
    r_vals = cache.items(spec.rows)

    def agg(subset, field_name, func):
        j = cache.index(field_name)
        vals = [r[j] for r in subset if r[j] not in (None, "")]
        if func == "count":
            return len(vals)
        if not vals:
            return None
        return sum(vals) if func == "sum" else sum(vals) / len(vals)

    def key(r, j):
        return r[j] if r[j] not in ("", None) else None

    table = []
    if spec.cols:
        (fname, func, label, _), = spec.data
        jc = cache.index(spec.cols)
        c_vals = cache.items(spec.cols)
        table.append([label, spec.col_caption or "Étiquettes de colonnes"]
                     + [None] * len(col_items))
        table.append([spec.row_caption or "Étiquettes de lignes"]
                     + [c_vals[i] if c_vals[i] is not None else "(vide)" for i in col_items]
                     + ["Total"])
        for i in row_items:
            sub = [r for r in rows if key(r, jr) == r_vals[i]]
            table.append([r_vals[i] if r_vals[i] is not None else "(vide)"]
                         + [agg([r for r in sub if key(r, jc) == c_vals[k]], fname, func)
                            for k in col_items] + [agg(sub, fname, func)])
        table.append(["Total"] + [agg([r for r in rows if key(r, jc) == c_vals[k]], fname, func)
                                  for k in col_items] + [agg(rows, fname, func)])
    else:
        table.append([spec.row_caption or "Étiquettes de lignes"] + [d[2] for d in spec.data])
        for i in row_items:
            sub = [r for r in rows if key(r, jr) == r_vals[i]]
            table.append([r_vals[i] if r_vals[i] is not None else "(vide)"]
                         + [agg(sub, d[0], d[1]) for d in spec.data])
        table.append(["Total"] + [agg(rows, d[0], d[1]) for d in spec.data])
    return table


def pivot_table_xml(cache: Cache, spec: PivotSpec, cache_id):
    rows, row_items, col_items = pivot_shape(cache, spec)
    table = pivot_values(cache, spec)
    n_rows, n_cols = len(table), len(table[0])
    r0, c0 = spec.top_left
    ref = f"{cell_ref(r0, c0)}:{cell_ref(r0 + n_rows - 1, c0 + n_cols - 1)}"
    if spec.cols:
        loc = f'<location ref="{ref}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"'
    else:
        loc = f'<location ref="{ref}" firstHeaderRow="0" firstDataRow="1" firstDataCol="1"'
    loc += ' rowPageCount="1" colPageCount="1"/>' if spec.page else "/>"

    data_idx = {cache.index(d[0]) for d in spec.data}
    pf = []
    for j, f in enumerate(cache.fields):
        name = f["name"]
        attrs = []
        items = ""
        if name in (spec.rows, spec.cols) or (spec.page and name == spec.page[0]):
            axis = ("axisRow" if name == spec.rows else
                    "axisCol" if name == spec.cols else "axisPage")
            attrs.append(f'axis="{axis}"')
            n = len(cache.items(name))
            items = (f'<items count="{n + 1}">'
                     + "".join(f'<item x="{i}"/>' for i in range(n))
                     + '<item t="default"/></items>')
        if j in data_idx:
            attrs.append('dataField="1"')
        attrs.append('showAll="0"')
        pf.append(f'<pivotField {" ".join(attrs)}>{items}</pivotField>' if items
                  else f'<pivotField {" ".join(attrs)}/>')

    def items_xml(tag, idxs):
        out = []
        for i in idxs:
            out.append("<i><x/></i>" if i == 0 else f'<i><x v="{i}"/></i>')
        out.append('<i t="grand"><x/></i>')
        return f'<{tag} count="{len(out)}">{"".join(out)}</{tag}>'

    xml = [loc, f'<pivotFields count="{len(pf)}">{"".join(pf)}</pivotFields>',
           f'<rowFields count="1"><field x="{cache.index(spec.rows)}"/></rowFields>',
           items_xml("rowItems", row_items)]
    if spec.cols:
        xml.append(f'<colFields count="1"><field x="{cache.index(spec.cols)}"/></colFields>')
        xml.append(items_xml("colItems", col_items))
    elif len(spec.data) > 1:
        xml.append('<colFields count="1"><field x="-2"/></colFields>')
        xml.append(f'<colItems count="{len(spec.data)}">' + "".join(
            "<i><x/></i>" if k == 0 else f'<i i="{k}"><x v="{k}"/></i>'
            for k in range(len(spec.data))) + "</colItems>")
    else:
        xml.append('<colItems count="1"><i/></colItems>')
    if spec.page:
        sel = cache.items(spec.page[0]).index(spec.page[1])
        xml.append(f'<pageFields count="1"><pageField fld="{cache.index(spec.page[0])}" '
                   f'item="{sel}" hier="-1"/></pageFields>')
    dfs = []
    for fname, func, label, numfmt in spec.data:
        sub = "" if func == "sum" else f' subtotal="{func}"'
        dfs.append(f'<dataField name="{esc(label)}" fld="{cache.index(fname)}"{sub} '
                   f'baseField="0" baseItem="0" numFmtId="{numfmt}"/>')
    xml.append(f'<dataFields count="{len(dfs)}">{"".join(dfs)}</dataFields>')
    xml.append('<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" '
               'showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>')
    caps = ""
    if spec.row_caption:
        caps += f' rowHeaderCaption="{esc(spec.row_caption)}"'
    if spec.col_caption:
        caps += f' colHeaderCaption="{esc(spec.col_caption)}"'
    head = (f'<pivotTableDefinition xmlns="{NS_MAIN}" name="{esc(spec.name)}" '
            f'cacheId="{cache_id}" applyNumberFormats="0" applyBorderFormats="0" '
            'applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" '
            'applyWidthHeightFormats="1" dataCaption="Valeurs" grandTotalCaption="Total" '
            f'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" '
            f'itemPrintTitles="1" createdVersion="6" indent="0" outline="1" '
            f'outlineData="1" multipleFieldFilters="0"{caps}>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + head
            + "".join(xml) + "</pivotTableDefinition>")


def inject_pivots(pkg: Package, src: SourceTable, specs):
    cache = Cache(src)
    cache_id = 1
    pkg.put("xl/pivotCache/pivotCacheDefinition1.xml", cache.definition_xml())
    pkg.put("xl/pivotCache/pivotCacheRecords1.xml", cache.records_xml())
    pkg.add_rel("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", REL_PIVOT_RECORDS,
                "pivotCacheRecords1.xml")
    pkg.add_override("xl/pivotCache/pivotCacheDefinition1.xml", CT["pivotCacheDefinition"])
    pkg.add_override("xl/pivotCache/pivotCacheRecords1.xml", CT["pivotCacheRecords"])
    rid = pkg.add_rel("xl/_rels/workbook.xml.rels", REL_PIVOT_CACHE,
                      "pivotCache/pivotCacheDefinition1.xml")
    wb = pkg.text("xl/workbook.xml")
    block = f'<pivotCaches><pivotCache cacheId="{cache_id}" r:id="{rid}"/></pivotCaches>'
    wb, n = re.subn(r"(<calcPr[^>]*/>)", r"\1" + block, wb, count=1)
    if n != 1:
        raise ValueError("calcPr introuvable dans workbook.xml")
    pkg.put("xl/workbook.xml", wb)

    for k, spec in enumerate(specs, 1):
        part = f"xl/pivotTables/pivotTable{k}.xml"
        pkg.put(part, pivot_table_xml(cache, spec, cache_id))
        pkg.add_rel(rels_name(part), REL_PIVOT_CACHE,
                    "../pivotCache/pivotCacheDefinition1.xml")
        pkg.add_override(part, CT["pivotTable"])
        sheet = pkg.sheet_part(spec.sheet)
        pkg.add_rel(rels_name(sheet), REL_PIVOT_TABLE, f"../pivotTables/pivotTable{k}.xml")
    return cache


# --------------------------------------------------------------------------- #
# Styles et thème
# --------------------------------------------------------------------------- #
def numfmt_id(pkg: Package, code):
    """Identifiant du format de nombre `code` déclaré dans styles.xml."""
    styles = pkg.text("xl/styles.xml")
    for fid, fcode in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/>', styles):
        if fcode == esc(code):
            return int(fid)
    raise KeyError(code)


def set_theme_colors(pkg: Package, accents):
    """Remplace les couleurs accent1..accent6 du thème par `accents` (hexadécimal).

    Le TreeMap natif colore chaque catégorie avec accent1, accent2… dans l'ordre
    d'apparition : on aligne ainsi ses couleurs sur la charte du classeur.
    """
    theme = pkg.text("xl/theme/theme1.xml")
    for k, color in enumerate(accents, 1):
        theme, n = re.subn(rf'(<a:accent{k}><a:srgbClr val=")[0-9A-Fa-f]{{6}}(")',
                           rf"\g<1>{color}\2", theme)
        if n != 1:
            raise ValueError(f"accent{k} introuvable dans le thème")
    pkg.put("xl/theme/theme1.xml", theme)
