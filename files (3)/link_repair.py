"""
link_repair.repair_external_links(path)

Registers external-workbook references that openpyxl wrote as raw-string
formulas (e.g.  ='[C:/data/File.xlsx]Sheet'!A1) into a proper OOXML
external-link table and converts the in-cell formulas to Excel's native
indexed form (='[1]Sheet'!A1).

Result: Excel opens the file WITHOUT the "recover unreadable content /
repair links" dialog, and the links resolve correctly.

Guarantees
----------
* Existing (already-registered) external links are never modified or dropped.
* Idempotent: running twice is a no-op the second time.
* Defensive: on ANY error the original file is restored byte-for-byte, so it
  can never leave the output worse than before.
"""
import os, re, shutil, zipfile, tempfile
from urllib.parse import quote
from xml.sax.saxutils import unescape as _xml_unescape

NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG  = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_EXT     = NS_R + "/externalLink"
REL_EXTPATH = NS_R + "/externalLinkPath"
CT_EXT = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"

_BRACKET   = re.compile(r"\[([^\[\]]+)\]")
_TOKEN     = re.compile(r"'?\[([^\[\]]+)\]([^'!\[\]]*)'?!")
_XL_EXT    = (".xlsx", ".xlsm", ".xls", ".xlsb")


def _uri(path):
    """Build a file:// URI from a real (already XML-unescaped) filesystem path."""
    p = path.replace("\\", "/")
    if p.lower().startswith("file:"):
        return p
    q = quote(p, safe="/:")
    # POSIX absolute path (/home/..)  -> file:///home/..   (avoid file:////)
    if q.startswith("/"):
        return "file://" + q
    # Windows / relative -> file:///C:/..
    return "file:///" + q


def _esc(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def repair_external_links(xlsx_path):
    backup = xlsx_path + ".linkbak"
    shutil.copy2(xlsx_path, backup)
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            parts = {n: z.read(n) for n in z.namelist()}

        sheets = [n for n in parts if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]

        # ---- existing external links (index -> normalised target) --------------
        existing_max = 0
        existing_norm = {}                       # normalised uri -> idx
        for n in list(parts):
            m = re.match(r"xl/externalLinks/externalLink(\d+)\.xml$", n)
            if not m:
                continue
            idx = int(m.group(1))
            existing_max = max(existing_max, idx)
            relp = f"xl/externalLinks/_rels/externalLink{idx}.xml.rels"
            tm = re.search(r'Target="([^"]+)"', parts.get(relp, b"").decode("utf-8"))
            if tm:
                existing_norm[tm.group(1).replace("\\", "/").lower()] = idx

        # ---- scan formulas for raw-string external refs ------------------------
        # Keys are the RAW (xml-escaped) path exactly as it appears in the sheet
        # XML – needed so the in-body rewrite matches.  We keep the real
        # (unescaped) path alongside for building link targets / dedup.
        path_sheets = {}     # raw_path -> {raw_sheet names}
        path_real   = {}     # raw_path -> real (unescaped) filesystem path
        for n in sheets:
            body = parts[n].decode("utf-8")
            for f in re.findall(r"<f[^>]*>(.*?)</f>", body, flags=re.S):
                for raw_path, raw_sheet in _TOKEN.findall(f):
                    real_path = _xml_unescape(raw_path)
                    if os.path.splitext(real_path)[1].lower() in _XL_EXT:
                        path_sheets.setdefault(raw_path, set()).add(raw_sheet)
                        path_real[raw_path] = real_path

        # keep only paths that are NOT already registered
        new_paths = {}
        for path in path_sheets:
            if _uri(path_real[path]).lower() not in existing_norm:
                new_paths[path] = None
        if not new_paths:
            os.remove(backup)
            return False                         # nothing to do

        # assign indices after the existing ones
        idx = existing_max
        for path in new_paths:
            idx += 1
            new_paths[path] = idx

        # map EVERY referenced (raw) path -> its final index (existing or new)
        path_index = {}
        for path in path_sheets:
            u = _uri(path_real[path]).lower()
            path_index[path] = existing_norm.get(u, new_paths.get(path))

        # ---- create parts for the new links ------------------------------------
        for path, i in new_paths.items():
            # raw_sheet came from XML (already escaped) -> use verbatim
            names = "".join(f'<sheetName val="{s}"/>' for s in sorted(path_sheets[path]))
            parts[f"xl/externalLinks/externalLink{i}.xml"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<externalLink xmlns="{NS_MAIN}">'
                f'<externalBook xmlns:r="{NS_R}" r:id="rId1">'
                f'<sheetNames>{names}</sheetNames></externalBook></externalLink>'
            ).encode("utf-8")
            parts[f"xl/externalLinks/_rels/externalLink{i}.xml.rels"] = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{NS_PKG}">'
                f'<Relationship Id="rId1" Type="{REL_EXTPATH}" '
                f'Target="{_esc(_uri(path_real[path]))}" TargetMode="External"/></Relationships>'
            ).encode("utf-8")

        # ---- content types -----------------------------------------------------
        ct = parts["[Content_Types].xml"].decode("utf-8")
        add = "".join(
            f'<Override PartName="/xl/externalLinks/externalLink{i}.xml" ContentType="{CT_EXT}"/>'
            for i in new_paths.values()
            if f"/xl/externalLinks/externalLink{i}.xml" not in ct
        )
        if add:
            ct = ct.replace("</Types>", add + "</Types>")
        parts["[Content_Types].xml"] = ct.encode("utf-8")

        # ---- workbook rels : one relationship per new link ---------------------
        wr = parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8")
        if not wr.strip():
            wr = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  f'<Relationships xmlns="{NS_PKG}"></Relationships>')
        next_rid = max([int(x) for x in re.findall(r'Id="rId(\d+)"', wr)] or [0])
        new_refs = []                            # (idx, rId) appended in index order
        add_rels = ""
        for i in sorted(new_paths.values()):
            next_rid += 1
            rid = f"rId{next_rid}"
            new_refs.append((i, rid))
            add_rels += (f'<Relationship Id="{rid}" Type="{REL_EXT}" '
                         f'Target="externalLinks/externalLink{i}.xml"/>')
        wr = wr.replace("</Relationships>", add_rels + "</Relationships>")
        parts["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")

        # ---- workbook.xml : APPEND new externalReference elements --------------
        wb = parts["xl/workbook.xml"].decode("utf-8")
        if "xmlns:r=" not in wb.split(">", 1)[0]:
            wb = re.sub(r"<workbook\b", f'<workbook xmlns:r="{NS_R}"', wb, count=1)
        refs = "".join(f'<externalReference r:id="{rid}"/>' for _, rid in new_refs)
        if "<externalReferences>" in wb:
            wb = wb.replace("</externalReferences>", refs + "</externalReferences>", 1)
        else:
            block = f"<externalReferences>{refs}</externalReferences>"
            if "<definedNames" in wb:
                wb = wb.replace("<definedNames", block + "<definedNames", 1)
            elif "<calcPr" in wb:
                wb = wb.replace("<calcPr", block + "<calcPr", 1)
            elif "</sheets>" in wb:
                wb = wb.replace("</sheets>", "</sheets>" + block, 1)
        # force clean recalc on open
        if "<calcPr" in wb:
            def _fc(mo):
                a = mo.group(1)
                if "fullCalcOnLoad" in a:
                    return "<calcPr" + re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', a) + "/>"
                return "<calcPr" + a + ' fullCalcOnLoad="1"/>'
            wb = re.sub(r"<calcPr([^>]*?)/>", _fc, wb, count=1)
        parts["xl/workbook.xml"] = wb.encode("utf-8")

        # ---- rewrite raw-string formulas : [path] -> [idx] ---------------------
        for n in sheets:
            body = parts[n].decode("utf-8")
            body = _BRACKET.sub(
                lambda mo: f"[{path_index[mo.group(1)]}]" if mo.group(1) in path_index else mo.group(0),
                body,
            )
            parts[n] = body.encode("utf-8")

        # ---- drop stale calcChain (Excel rebuilds it, avoids repair prompt) -----
        if "xl/calcChain.xml" in parts:
            del parts["xl/calcChain.xml"]
            ct = parts["[Content_Types].xml"].decode("utf-8")
            ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
            parts["[Content_Types].xml"] = ct.encode("utf-8")
            wr = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
            wr = re.sub(r'<Relationship[^>]*calcChain\.xml"[^>]*/>', "", wr)
            parts["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")

        # ---- repackage ---------------------------------------------------------
        fd, tmp = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n, data in parts.items():
                z.writestr(n, data)
        shutil.move(tmp, xlsx_path)
        os.remove(backup)
        return True
    except Exception:
        shutil.move(backup, xlsx_path)           # restore original untouched
        raise
