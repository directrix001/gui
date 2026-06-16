"""
variance_engine.py  —  v6.3
==========================================
Builds on v6.0 (no broken links / no workbook repair dialog).

NEW in v6.3
  • INTERNAL read-only / protection fix.  v6.2 only cleared the filesystem
    read-only attribute; this version also strips the read-only settings
    stored INSIDE the workbook, which were copied from the master and made
    the output open read-only / un-editable:
        - <fileSharing>        (Always-Open-Read-Only / Read-only-recommended
                                / write-reservation "password to modify")
        - <workbookProtection> (structure lock that blocks add/hide sheets)
        - _MarkAsFinal         (Mark as Final, in docProps/custom.xml)
    Done on the copied output before either tier runs (see
    _strip_workbook_locks).  The xlwings tier also force-opens read-write
    and unprotects in-session as a safety net (_open_read_write).

NEW in v6.2
  • Read-only master fix: the output copy is forced writable immediately
    after copying (shutil.copy2 otherwise preserves the master's read-only
    permission bits, leaving the output un-editable).  See _make_writable().

NEW in v6.1
  • Existing MTD/YTD sheets in the MASTER file are HIDDEN (not deleted)
    in the generated output, before the comparison-scenario tabs are added.
    A sheet is treated as periodic if its name contains "mtd" or "ytd"
    (case-insensitive).  Implemented in BOTH engine tiers:
        - xlwings tier  -> sheet.api.Visible = 0  (xlSheetHidden)
        - zip tier      -> state="hidden" on the <sheet> element
  • Safety: after hiding, the workbook's activeTab/firstSheet are pointed
    at the first VISIBLE sheet so Excel never opens on a hidden tab
    (which would otherwise trigger a repair prompt).

v6.0 root-cause fixes (unchanged) are documented inline below.

BUG-1  Absolute target paths with leading "/"
        FIX: strip leading "/" with lstrip("/") before joining.
BUG-2  Injected <sheet> elements missing xmlns:r declaration
        FIX: ensure workbook.xml root carries xmlns:r.
BUG-3  openpyxl.save() after zip surgery destroys patches
        FIX: audit log injected as raw XML; openpyxl.save() never called.
BUG-4  sharedStrings.xml absent in output but referenced by copied sheets
        FIX: copy sharedStrings.xml from template + register it.
BUG-5  styles.xml index mismatch
        FIX: replace output styles.xml with template's when template is larger.
"""

import os
import re
import stat
import shutil
import zipfile
import tempfile
from datetime import datetime

# openpyxl used ONLY for reading sheet names from the template (read-only, no save)
from openpyxl import load_workbook as _openpyxl_load

# ── Relationship / content-type URI constants ────────────────────────────────
_REL_WS      = ("http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships/worksheet")
_REL_SS      = ("http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships/sharedStrings")
_REL_STYLES  = ("http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships/styles")
_NS_R        = ("http://schemas.openxmlformats.org/officeDocument/"
                "2006/relationships")
_CT_WS       = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.worksheet+xml")
_CT_SS       = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sharedStrings+xml")
_CT_STYLES   = ("application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.styles+xml")

# ═══════════════════════════════════════════════════════════════
#  CONFIGURATION  — override if auto-detect picks wrong sheet
# ═══════════════════════════════════════════════════════════════
MTD_SHEET_NAME: str | None = None
YTD_SHEET_NAME: str | None = None

# Keywords that mark a sheet as "periodic" and therefore hidden in the master.
PERIODIC_KEYWORDS = ("mtd", "ytd")


def _is_periodic(name: str) -> bool:
    low = (name or "").lower()
    return any(k in low for k in PERIODIC_KEYWORDS)


def _make_writable(path: str):
    """Clear the read-only attribute so the file can be modified / replaced."""
    try:
        os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    except Exception as e:
        print(f"[engine] Could not clear read-only on {os.path.basename(path)}: {e}")


def _strip_workbook_locks(output_path: str):
    """
    Remove every INTERNAL read-only / protection mechanism from the output
    .xlsx so it opens fully editable and so the Excel-COM tier can add and
    hide sheets.  Operates directly on the zip, so it benefits BOTH tiers.

    Removes:
      • <fileSharing .../>        -> "Always Open Read-Only" /
                                     "Read-only recommended" /
                                     write-reservation (modify) password
      • <workbookProtection .../> -> structure / window lock that blocks
                                     adding, deleting, hiding sheets
      • _MarkAsFinal property     -> "Mark as Final" (docProps/custom.xml)
    """
    removed: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        ex = os.path.join(tmp, "x")
        with zipfile.ZipFile(output_path, "r") as z:
            z.extractall(ex)

        # ── xl/workbook.xml : fileSharing + workbookProtection ──────────────
        wb = os.path.join(ex, "xl", "workbook.xml")
        if os.path.isfile(wb):
            text = orig = _rt(wb)

            if re.search(r'<fileSharing\b', text, re.I):
                removed.append("fileSharing (read-only/modify-password)")
            text = re.sub(r'<fileSharing\b[^>]*/\s*>', '', text, flags=re.I)
            text = re.sub(r'<fileSharing\b[^>]*>.*?</fileSharing\s*>',
                          '', text, flags=re.I | re.DOTALL)

            if re.search(r'<workbookProtection\b', text, re.I):
                removed.append("workbookProtection (structure lock)")
            text = re.sub(r'<workbookProtection\b[^>]*/\s*>', '', text, flags=re.I)
            text = re.sub(r'<workbookProtection\b[^>]*>.*?</workbookProtection\s*>',
                          '', text, flags=re.I | re.DOTALL)

            if text != orig:
                _wt(wb, text)

        # ── docProps/custom.xml : Mark as Final ─────────────────────────────
        cust = os.path.join(ex, "docProps", "custom.xml")
        if os.path.isfile(cust):
            text = orig = _rt(cust)
            if re.search(r'name=["\']_MarkAsFinal["\']', text, re.I):
                removed.append("Mark as Final")
                text = re.sub(
                    r'<property\b[^>]*name=["\']_MarkAsFinal["\'][^>]*>'
                    r'.*?</property\s*>',
                    '', text, flags=re.I | re.DOTALL)
                # also handle a self-closing <property .../> form
                text = re.sub(
                    r'<property\b[^>]*name=["\']_MarkAsFinal["\'][^>]*/\s*>',
                    '', text, flags=re.I)
            if text != orig:
                _wt(cust, text)

        if removed:
            _repack(ex, output_path)
            _make_writable(output_path)
            print(f"[engine] Stripped internal locks: {', '.join(removed)}")
        else:
            print("[engine] No internal read-only/protection flags found")


# ═══════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════

def run_variance(args: dict) -> str:
    """
    Build one output .xlsx:
      • Base     = master file copied to output folder
      • Hidden   = every existing MTD/YTD sheet in the master
      • Added    = MTD + YTD tabs from template for every scenario pair
      • Last tab = Input Log audit sheet
    Returns full path of the saved output file.
    """
    pairs         = args["scenario_pairs"]          # [(sc1,sc2), ...]
    quarter       = args["quarter"]
    month         = args["month"]
    master_file   = args["master_file"]
    template_file = args["template_file"]
    out_folder    = args["output_folder"]
    ts            = args["timestamp"]

    os.makedirs(out_folder, exist_ok=True)

    if not os.path.isfile(template_file):
        raise FileNotFoundError(f"Template file not found:\n{template_file}")
    if not os.path.isfile(master_file):
        raise FileNotFoundError(f"Master file not found:\n{master_file}")

    mtd_name, ytd_name = _resolve_template_sheets(template_file)

    q_label     = quarter.split()[0]
    output_path = os.path.join(
        out_folder, f"Variance__{q_label}_{month}__v{ts}.xlsx")
    shutil.copy2(master_file, output_path)

    # The master may be read-only (network share / locked attribute).
    # copy2() preserves permission bits, so the fresh copy can inherit the
    # read-only flag and then nothing can write to it.  Force it writable.
    _make_writable(output_path)

    # The master may also carry an INTERNAL read-only / protection setting
    # (Always-Open-Read-Only, Read-only-recommended, write-reservation
    # password, Mark-as-Final, or structure protection).  These live inside
    # the workbook XML and are copied along with it, so the output opens
    # read-only and/or blocks sheet add/hide.  Strip them now, before either
    # engine tier touches the file.
    _strip_workbook_locks(output_path)

    print(f"[engine] Output base: {os.path.basename(output_path)}")

    # ── TIER 1: xlwings COM (Excel-native copy, best quality) ───────────────
    if _try_xlwings(output_path, template_file, pairs,
                    mtd_name, ytd_name, args):
        print(f"[engine] Done (xlwings): {os.path.basename(output_path)}")
        return output_path

    # ── TIER 2: zip / XML surgery (no Excel required, stdlib only) ──────────
    print("[engine] Using zip-level copy.")
    _zip_copy(output_path, template_file, pairs, mtd_name, ytd_name, args)
    print(f"[engine] Done (zip): {os.path.basename(output_path)}")
    return output_path


# ═══════════════════════════════════════════════════════════════
#  TIER 1 — xlwings / Excel COM
# ═══════════════════════════════════════════════════════════════

def _open_read_write(app, abspath):
    """
    Open a workbook through xlwings/Excel forcing read-write, ignoring any
    read-only recommendation or write-reservation prompt.  Falls back
    gracefully across xlwings versions.
    """
    # 1) Newer xlwings kwargs
    try:
        return app.books.open(abspath, read_only=False,
                              ignore_read_only_recommended=True)
    except TypeError:
        pass
    except Exception:
        pass
    # 2) Raw COM Workbooks.Open with explicit flags
    try:
        com = app.api.Workbooks.Open(
            abspath,
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True)
        # wrap the COM book back into an xlwings Book
        import xlwings as xw
        for b in app.books:
            if b.fullname.lower() == abspath.lower():
                return b
        return xw.Book(impl=app.books(com.Name).impl)
    except Exception:
        pass
    # 3) Last resort: plain open (file locks were already stripped)
    return app.books.open(abspath)


def _try_xlwings(output_path, template_file, pairs,
                 mtd_name, ytd_name, args) -> bool:
    try:
        import xlwings as xw
    except ImportError:
        print("[engine] xlwings not installed.")
        return False
    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts  = False
        app.screen_updating = False

        tpl_wb = app.books.open(os.path.abspath(template_file))
        out_wb = _open_read_write(app, os.path.abspath(output_path))

        # Belt-and-suspenders: clear any surviving in-session read-only state.
        try:
            out_wb.api.ReadOnlyRecommended = False
        except Exception:
            pass
        try:
            out_wb.api.Final = False     # clear Mark-as-Final if present
        except Exception:
            pass
        try:
            # Unprotect structure so sheets can be copied / hidden / deleted.
            out_wb.api.Unprotect()
        except Exception:
            pass

        # Snapshot the master's ORIGINAL sheet names before we add anything,
        # so the new comparison tabs (which also contain "MTD"/"YTD") are
        # never mistaken for existing master sheets.
        original_names = [sh.name for sh in out_wb.sheets]

        for sc1, sc2 in pairs:
            lbl = f"{sc1} vs {sc2}"
            for src, dst in [(mtd_name, f"MTD {lbl}"[:31]),
                              (ytd_name, f"YTD {lbl}"[:31])]:
                for sh in out_wb.sheets:
                    if sh.name == dst:
                        sh.delete(); break
                tpl_wb.sheets[src].api.Copy(After=out_wb.sheets[-1].api)
                out_wb.sheets[-1].name = dst
                print(f"[engine][xlwings] '{src}' → '{dst}'")

        # ── Hide existing MTD/YTD sheets from the master ────────────────────
        existing_now = {sh.name for sh in out_wb.sheets}
        for name in original_names:
            if _is_periodic(name) and name in existing_now:
                try:
                    out_wb.sheets[name].api.Visible = 0   # xlSheetHidden
                    print(f"[engine][xlwings] Hid master sheet: '{name}'")
                except Exception as e:
                    print(f"[engine][xlwings] Could not hide '{name}': {e}")

        # Break external links back to template
        try:
            links = out_wb.api.LinkSources(1)
            if links:
                tl = os.path.basename(template_file).lower()
                for lnk in links:
                    if tl in str(lnk).lower():
                        out_wb.api.BreakLink(Name=lnk, Type=1)
                        print(f"[engine][xlwings] Broke: {lnk}")
        except Exception as e:
            print(f"[engine][xlwings] BreakLink: {e}")

        app.api.CalculateFull()

        # Write audit log via xlwings (plain values, no formulas)
        try:
            ws = out_wb.sheets.add("Input Log", after=out_wb.sheets[-1])
            _xw_write_audit(ws, args, pairs)
        except Exception as e:
            print(f"[engine][xlwings] Audit: {e}")

        # Make sure Excel opens on a visible sheet, not a hidden one.
        try:
            for sh in out_wb.sheets:
                if sh.api.Visible == -1:   # xlSheetVisible
                    sh.activate()
                    break
        except Exception:
            pass

        out_wb.save(); out_wb.close()
        tpl_wb.close(); app.quit()
        return True
    except Exception as e:
        print(f"[engine] xlwings error ({e})")
        try:
            if app: app.quit()
        except Exception:
            pass
        return False


def _xw_write_audit(ws, args, pairs):
    r = 1
    ws.cells(r,1).value = "INPUT FILE LOG"; r+=1
    for lbl, val in [("Generated", datetime.now().strftime("%d %b %Y  %H:%M")),
                     ("Quarter",   args.get("quarter","")),
                     ("Month",     args.get("month","")),
                     ("Master",    args.get("master_file","")),
                     ("Template",  args.get("template_file",""))]:
        ws.cells(r,1).value=lbl; ws.cells(r,2).value=val; r+=1
    r+=1; ws.cells(r,1).value="SCENARIO PAIRS"; r+=1
    for c,h in enumerate(["#","Sc1","Sc2","MTD Tab","YTD Tab"],1):
        ws.cells(r,c).value=h
    r+=1
    for i,(s1,s2) in enumerate(pairs,1):
        lb=f"{s1} vs {s2}"
        for c,v in enumerate([i,s1,s2,f"MTD {lb}"[:31],f"YTD {lb}"[:31]],1):
            ws.cells(r,c).value=v
        r+=1


# ═══════════════════════════════════════════════════════════════
#  TIER 2 — zip / XML surgery
# ═══════════════════════════════════════════════════════════════

def _zip_copy(output_path: str, template_file: str,
              pairs: list, mtd_name: str, ytd_name: str, args: dict):

    tpl_base    = os.path.basename(template_file)       # "Template.xlsx"
    tpl_base_nx = os.path.splitext(tpl_base)[0]         # "Template"

    with tempfile.TemporaryDirectory() as tmp:
        out_dir = os.path.join(tmp, "out")
        tpl_dir = os.path.join(tmp, "tpl")

        with zipfile.ZipFile(output_path,   "r") as z: z.extractall(out_dir)
        with zipfile.ZipFile(template_file, "r") as z: z.extractall(tpl_dir)

        # ── Step 0: ensure workbook.xml root declares xmlns:r (BUG-2 fix) ──
        _ensure_r_namespace_on_root(out_dir)

        # ── Step 0b: hide existing MTD/YTD sheets in the master ─────────────
        _hide_master_periodic_sheets(out_dir)

        # ── Step 1: sync styles + sharedStrings (BUG-4, BUG-5) ─────────────
        _sync_styles(tpl_dir, out_dir)
        _sync_shared_strings(tpl_dir, out_dir)

        # ── Step 2: build sheet map from template ───────────────────────────
        tpl_map  = _sheet_map(tpl_dir)   # {name: "/xl/worksheets/sheetN.xml"}
        next_sh  = _next_sheet_idx(out_dir)
        next_rid = _next_rid(out_dir)

        to_copy = []
        for sc1, sc2 in pairs:
            lbl = f"{sc1} vs {sc2}"
            to_copy.append((mtd_name, f"MTD {lbl}"[:31]))
            to_copy.append((ytd_name, f"YTD {lbl}"[:31]))

        # ── Step 3: copy each sheet ─────────────────────────────────────────
        for tpl_name, new_tab in to_copy:
            if tpl_name not in tpl_map:
                raise ValueError(
                    f"Sheet '{tpl_name}' not in template.\n"
                    f"Available: {list(tpl_map.keys())}")

            tpl_target = tpl_map[tpl_name]          # "/xl/worksheets/sheet2.xml"
            tpl_rel    = tpl_target.lstrip("/")      # BUG-1 fix: strip leading /
            new_rel    = f"xl/worksheets/sheet{next_sh}.xml"
            new_abs    = f"/xl/worksheets/sheet{next_sh}.xml"
            rid        = f"rId{next_rid}"

            # 3a. Copy worksheet XML, strip external-link prefixes
            src = os.path.join(tpl_dir, tpl_rel)
            dst = os.path.join(out_dir, new_rel)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            _strip_ext_links(dst, tpl_base, tpl_base_nx)

            # 3b. Copy worksheet .rels + all dependencies
            tpl_ws_rels = os.path.join(
                tpl_dir, "xl", "worksheets", "_rels",
                os.path.basename(tpl_rel) + ".rels")
            if os.path.isfile(tpl_ws_rels):
                out_rels_dir = os.path.join(out_dir,"xl","worksheets","_rels")
                os.makedirs(out_rels_dir, exist_ok=True)
                out_ws_rels = os.path.join(
                    out_rels_dir, f"sheet{next_sh}.xml.rels")
                shutil.copy2(tpl_ws_rels, out_ws_rels)
                _copy_deps(tpl_ws_rels, tpl_dir, out_dir)

            # 3c. Register in workbook.xml, rels, content-types
            _wb_add_sheet(out_dir, new_tab, rid)
            _rels_add(out_dir, rid, new_abs, _REL_WS)
            _ct_add(out_dir, f"/xl/worksheets/sheet{next_sh}.xml", _CT_WS)

            next_sh  += 1
            next_rid += 1
            print(f"[engine][zip] '{tpl_name}' → '{new_tab}'")

        # ── Step 4: purge all external-link records ──────────────────────────
        _purge_ext_refs(out_dir, tpl_base, tpl_base_nx)

        # ── Step 5: inject audit log as raw XML (BUG-3 fix — no openpyxl save)
        _inject_audit(out_dir, args, pairs, next_sh, next_rid)

        # ── Step 6: point activeTab/firstSheet at first VISIBLE sheet ────────
        _fix_active_tab(out_dir)

        # ── Step 7: repack ───────────────────────────────────────────────────
        _repack(out_dir, output_path)


# ── xmlns:r on workbook root (BUG-2) ─────────────────────────────────────────

def _ensure_r_namespace_on_root(out_dir: str):
    """
    Add xmlns:r declaration to <workbook> root element if absent.
    """
    path = os.path.join(out_dir, "xl", "workbook.xml")
    text = _rt(path)
    if 'xmlns:r=' in text:
        return   # already declared
    text = re.sub(
        r'(<workbook\b)',
        r'\1 xmlns:r="' + _NS_R + '"',
        text, count=1)
    _wt(path, text)
    print("[engine][zip] Added xmlns:r to workbook root")


# ── hide existing MTD/YTD master sheets (NEW in v6.1) ────────────────────────

def _hide_master_periodic_sheets(out_dir: str):
    """
    Add state="hidden" to every <sheet> in workbook.xml whose name contains
    a periodic keyword (mtd / ytd).  Runs BEFORE any new sheet is added, so
    only the master's own sheets are affected.
    """
    path = os.path.join(out_dir, "xl", "workbook.xml")
    text = _rt(path)
    hidden: list[str] = []

    def _repl(m):
        tag  = m.group(0)                       # full <sheet .../> element
        name = _attr(tag, "name")
        rid  = _attr(tag, "r:id")
        if not (name and rid):                  # not a sheet-listing entry
            return tag
        if not _is_periodic(name):
            return tag
        hidden.append(name)
        # strip any existing state attribute, then add state="hidden"
        inner = tag[len("<sheet"):]
        inner = inner.rstrip(">").rstrip("/").rstrip()
        inner = re.sub(r'\s+state\s*=\s*["\'][^"\']*["\']', '', inner)
        return f'<sheet{inner} state="hidden"/>'

    text = re.sub(r'<sheet\b[^>]*>', _repl, text)
    _wt(path, text)
    if hidden:
        print(f"[engine][zip] Hid master MTD/YTD sheets: {hidden}")
    else:
        print("[engine][zip] No master MTD/YTD sheets to hide")


# ── active-tab safety (NEW in v6.1) ──────────────────────────────────────────

def _fix_active_tab(out_dir: str):
    """
    Ensure activeTab / firstSheet in <workbookView> point at the first VISIBLE
    sheet.  Excel refuses to open a workbook whose active tab is hidden and
    will offer to "repair" it — exactly what this engine avoids.
    """
    path = os.path.join(out_dir, "xl", "workbook.xml")
    text = _rt(path)

    visibility: list[bool] = []   # True = hidden
    for m in re.finditer(r'<sheet\b[^>]*>', text):
        tag = m.group(0)
        if not _attr(tag, "r:id"):
            continue
        is_hidden = bool(
            re.search(r'state\s*=\s*["\'](hidden|veryHidden)["\']', tag, re.I))
        visibility.append(is_hidden)

    first_visible = next((i for i, h in enumerate(visibility) if not h), 0)

    if not re.search(r'<workbookView\b', text):
        _wt(path, text)
        return

    def _wv(m):
        tag = m.group(0)
        self_close = tag.rstrip().endswith("/>")
        body = tag[:-2] if self_close else tag[:-1]
        body = re.sub(r'\s+activeTab\s*=\s*["\'][^"\']*["\']', '', body)
        body = re.sub(r'\s+firstSheet\s*=\s*["\'][^"\']*["\']', '', body)
        body = body.rstrip()
        body += f' firstSheet="{first_visible}" activeTab="{first_visible}"'
        return body + ("/>" if self_close else ">")

    text = re.sub(r'<workbookView\b[^>]*?>', _wv, text, count=1)
    _wt(path, text)
    print(f"[engine][zip] activeTab -> first visible sheet (index {first_visible})")


# ── styles sync (BUG-5) ───────────────────────────────────────────────────────

def _sync_styles(tpl_dir: str, out_dir: str):
    src = os.path.join(tpl_dir, "xl", "styles.xml")
    dst = os.path.join(out_dir, "xl", "styles.xml")
    if not os.path.isfile(src):
        return
    if (not os.path.isfile(dst) or
            os.path.getsize(src) > os.path.getsize(dst)):
        shutil.copy2(src, dst)
        print("[engine][zip] Synced styles.xml from template")


# ── sharedStrings sync (BUG-4) ───────────────────────────────────────────────

def _sync_shared_strings(tpl_dir: str, out_dir: str):
    src = os.path.join(tpl_dir, "xl", "sharedStrings.xml")
    dst = os.path.join(out_dir, "xl", "sharedStrings.xml")
    if not os.path.isfile(src):
        return
    shutil.copy2(src, dst)
    _ct_add(out_dir, "/xl/sharedStrings.xml", _CT_SS)
    rels_path = os.path.join(out_dir, "xl", "_rels", "workbook.xml.rels")
    if "sharedStrings" not in _rt(rels_path):
        _rels_add(out_dir, f"rId{_next_rid(out_dir)}",
                  "/xl/sharedStrings.xml", _REL_SS)
    print("[engine][zip] Synced sharedStrings.xml from template")


# ── sheet map ─────────────────────────────────────────────────────────────────

def _sheet_map(base_dir: str) -> dict:
    """
    Returns {sheet_name: target_path_as_in_rels}
    """
    wb_text   = _rt(os.path.join(base_dir, "xl", "workbook.xml"))
    rels_text = _rt(os.path.join(base_dir, "xl", "_rels", "workbook.xml.rels"))

    rid2tgt: dict[str, str] = {}
    for m in re.finditer(r'<Relationship\b([^>]+)>', rels_text):
        a   = m.group(1)
        rid = _attr(a, "Id")
        tgt = _attr(a, "Target")
        if rid and tgt:
            rid2tgt[rid] = tgt

    result: dict[str, str] = {}
    for m in re.finditer(r'<sheet\b([^>]+)/?\s*>', wb_text):
        a    = m.group(1)
        name = _attr(a, "name")
        rid  = _attr(a, "r:id")
        if name and rid and rid in rid2tgt:
            result[name] = rid2tgt[rid]
    return result


def _attr(attrs: str, key: str) -> str:
    """Extract attribute value from an attribute string."""
    m = re.search(r'\b' + re.escape(key) + r'\s*=\s*["\']([^"\']*)["\']',
                  attrs)
    return m.group(1) if m else ""


# ── index helpers ─────────────────────────────────────────────────────────────

def _next_sheet_idx(out_dir: str) -> int:
    d = os.path.join(out_dir, "xl", "worksheets")
    if not os.path.isdir(d):
        return 1
    nums = [int(m.group(1))
            for f in os.listdir(d)
            if (m := re.match(r"^sheet(\d+)\.xml$", f))]
    return max(nums, default=0) + 1


def _next_rid(out_dir: str) -> int:
    text = _rt(os.path.join(out_dir, "xl", "_rels", "workbook.xml.rels"))
    nums = [int(m.group(1)) for m in re.finditer(r'\brId(\d+)\b', text)]
    return max(nums, default=0) + 1


# ── raw XML patchers ──────────────────────────────────────────────────────────

def _wb_add_sheet(out_dir: str, name: str, rid: str):
    """Insert <sheet> into workbook.xml before </sheets>."""
    path = os.path.join(out_dir, "xl", "workbook.xml")
    text = _rt(path)
    ids  = [int(m.group(1))
            for m in re.finditer(r'\bsheetId=["\'](\d+)["\']', text)]
    sid  = max(ids, default=0) + 1
    safe = (name.replace("&","&amp;").replace('"',"&quot;")
                .replace("<","&lt;").replace(">","&gt;"))
    el   = (f'<sheet xmlns:r="{_NS_R}" name="{safe}" '
            f'sheetId="{sid}" r:id="{rid}"/>')
    if "</sheets>" in text:
        text = text.replace("</sheets>", el + "</sheets>", 1)
    else:
        text = text.replace("</workbook>",
                            f"<sheets>{el}</sheets></workbook>", 1)
    _wt(path, text)


def _rels_add(out_dir: str, rid: str, target: str, rel_type: str):
    path = os.path.join(out_dir, "xl", "_rels", "workbook.xml.rels")
    text = _rt(path)
    el   = (f'<Relationship Id="{rid}" Type="{rel_type}" '
            f'Target="{target}"/>')
    text = text.replace("</Relationships>", el + "</Relationships>", 1)
    _wt(path, text)


def _ct_add(out_dir: str, part: str, ct: str):
    path = os.path.join(out_dir, "[Content_Types].xml")
    text = _rt(path)
    if part in text:
        return
    el   = f'<Override PartName="{part}" ContentType="{ct}"/>'
    text = text.replace("</Types>", el + "</Types>", 1)
    _wt(path, text)


# ── strip external-link prefixes from formula cells ──────────────────────────

def _strip_ext_links(ws_path: str, tpl_base: str, tpl_base_nx: str):
    """
    Remove '[Template.xlsx]' or '[Template]' from every <f>…</f> element.
    """
    pats = [
        re.compile(r'\[' + re.escape(tpl_base)           + r'\]', re.I),
        re.compile(r'\[' + re.escape(tpl_base_nx) + r'\.xlsx\]', re.I),
        re.compile(r'\[' + re.escape(tpl_base_nx) + r'\.xlsm\]', re.I),
        re.compile(r'\[' + re.escape(tpl_base_nx)         + r'\]', re.I),
    ]
    def _clean(s: str) -> str:
        for p in pats: s = p.sub("", s)
        return s
    text = _rt(ws_path)
    text = re.sub(
        r'(<f(?:\s[^>]*)?>)(.*?)(</f>)',
        lambda m: m.group(1) + _clean(m.group(2)) + m.group(3),
        text, flags=re.DOTALL)
    _wt(ws_path, text)


# ── purge external references ─────────────────────────────────────────────────

def _purge_ext_refs(out_dir: str, tpl_base: str, tpl_base_nx: str):
    """
    Remove all external-link artefacts.
    """
    wb_path = os.path.join(out_dir, "xl", "workbook.xml")
    text    = _rt(wb_path)

    text = re.sub(
        r'<externalReferences\b[^>]*>.*?</externalReferences\s*>',
        '', text, flags=re.DOTALL | re.I)
    text = re.sub(
        r'<externalReferences\b[^/]*/\s*>',
        '', text, flags=re.I)

    tlo = re.escape(tpl_base.lower())
    nlo = re.escape(tpl_base_nx.lower())
    text = re.sub(
        r'<definedName\b[^>]*>.*?</definedName\s*>',
        lambda m: ''
            if (tlo in m.group(0).lower() or nlo in m.group(0).lower())
            else m.group(0),
        text, flags=re.DOTALL | re.I)
    _wt(wb_path, text)

    rels_path = os.path.join(out_dir, "xl", "_rels", "workbook.xml.rels")
    rt = _rt(rels_path)
    rt = re.sub(
        r'<Relationship\b[^>]*externalLink[^>]*/\s*>',
        '', rt, flags=re.I)
    _wt(rels_path, rt)

    ext_dir = os.path.join(out_dir, "xl", "externalLinks")
    if os.path.isdir(ext_dir):
        shutil.rmtree(ext_dir)
        print("[engine][zip] Removed externalLinks folder")


# ── copy chart / drawing / image dependencies ────────────────────────────────

def _copy_deps(rels_path: str, tpl_dir: str, out_dir: str):
    text = _rt(rels_path)
    for m in re.finditer(r'\bTarget=["\']([^"\']+)["\']', text):
        tgt = m.group(1)
        if tgt.startswith("#") or tgt.startswith("http"):
            continue
        src_rel = os.path.normpath(
            os.path.join("xl", "worksheets", tgt)).replace("\\", "/")
        src = os.path.join(tpl_dir, src_rel)
        dst = os.path.join(out_dir, src_rel)
        if os.path.isfile(src) and not os.path.isfile(dst):
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(src, dst)
            dep_rels = os.path.join(
                os.path.dirname(src), "_rels",
                os.path.basename(src) + ".rels")
            if os.path.isfile(dep_rels):
                dr_dst = os.path.join(
                    os.path.dirname(dst), "_rels",
                    os.path.basename(dst) + ".rels")
                os.makedirs(os.path.dirname(dr_dst), exist_ok=True)
                shutil.copy2(dep_rels, dr_dst)


# ── audit log injected as raw XML (BUG-3 fix) ────────────────────────────────

def _inject_audit(out_dir: str, args: dict, pairs: list,
                  sh_idx: int, rid_idx: int):
    """
    Build 'Input Log' sheet as minimal worksheet XML and inject it into the zip.
    """
    rows: list[list[str]] = []
    rows.append(["INPUT FILE LOG"])
    rows.append(["Generated",   datetime.now().strftime("%d %b %Y  %H:%M")])
    rows.append(["Quarter",     args.get("quarter", "")])
    rows.append(["Month",       args.get("month", "")])
    rows.append(["Master File", args.get("master_file", "")])
    rows.append(["Template",    args.get("template_file", "")])
    rows.append([])
    rows.append(["SCENARIO PAIRS"])
    rows.append(["#", "Scenario 1", "Scenario 2", "MTD Tab", "YTD Tab"])
    for i, (s1, s2) in enumerate(pairs, 1):
        lb = f"{s1} vs {s2}"
        rows.append([str(i), s1, s2, f"MTD {lb}"[:31], f"YTD {lb}"[:31]])
    rows.append([])
    rows.append(["INPUT FILES"])
    rows.append(["#", "Folder", "File Name", "Full Path"])
    for i, fp in enumerate(args.get("input_files", []), 1):
        rows.append([str(i), os.path.dirname(fp),
                     os.path.basename(fp), fp])

    def _xe(s: str) -> str:          # XML-escape a string
        return (s.replace("&","&amp;").replace("<","&lt;")
                 .replace(">","&gt;").replace('"',"&quot;"))

    sd = ""
    for ri, row in enumerate(rows, 1):
        if not row:
            continue
        row_xml = ""
        for ci, val in enumerate(row, 1):
            ref = f"{_col_ltr(ci)}{ri}"
            row_xml += (f'<c r="{ref}" t="inlineStr">'
                        f'<is><t>{_xe(str(val))}</t></is></c>')
        sd += f'<row r="{ri}">{row_xml}</row>'

    ws_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/'
        'spreadsheetml/2006/main">'
        '<sheetData>' + sd + '</sheetData>'
        '</worksheet>'
    )

    ws_rel  = f"xl/worksheets/sheet{sh_idx}.xml"
    ws_abs  = f"/xl/worksheets/sheet{sh_idx}.xml"
    ws_path = os.path.join(out_dir, ws_rel)
    os.makedirs(os.path.dirname(ws_path), exist_ok=True)
    _wt(ws_path, ws_xml)

    _wb_add_sheet(out_dir, "Input Log", f"rId{rid_idx}")
    _rels_add(out_dir, f"rId{rid_idx}", ws_abs, _REL_WS)
    _ct_add(out_dir, f"/xl/worksheets/sheet{sh_idx}.xml", _CT_WS)
    print("[engine][zip] Injected 'Input Log'")


def _col_ltr(n: int) -> str:
    r = ""
    while n:
        n, rem = divmod(n - 1, 26)
        r = chr(65 + rem) + r
    return r


# ── repack ────────────────────────────────────────────────────────────────────

def _repack(src_dir: str, out_path: str):
    tmp = out_path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(src_dir):
            if root == src_dir:
                files = list(files)
                if "[Content_Types].xml" in files:
                    files.remove("[Content_Types].xml")
                    files.insert(0, "[Content_Types].xml")
            for fname in files:
                abs_p = os.path.join(root, fname)
                arc   = os.path.relpath(abs_p, src_dir).replace("\\", "/")
                zf.write(abs_p, arc)
    if os.path.exists(out_path):
        _make_writable(out_path)
    os.replace(tmp, out_path)


# ── raw text read / write ─────────────────────────────────────────────────────

def _rt(path: str) -> str:
    with open(path, "rb") as f:
        raw = f.read()
    m = re.match(rb'<\?xml[^>]+encoding=["\']([^"\']+)["\']', raw[:80])
    enc = m.group(1).decode("ascii") if m else "utf-8"
    return raw.decode(enc, errors="replace")


def _wt(path: str, text: str):
    with open(path, "wb") as f:
        f.write(text.encode("utf-8"))


# ═══════════════════════════════════════════════════════════════
#  TEMPLATE SHEET AUTO-DETECT
# ═══════════════════════════════════════════════════════════════

def _resolve_template_sheets(tpl_path: str):
    wb    = _openpyxl_load(tpl_path, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    mtd = MTD_SHEET_NAME or next(
        (n for n in names if "mtd" in n.lower()), None)
    ytd = YTD_SHEET_NAME or next(
        (n for n in names if "ytd" in n.lower()), None)
    if not mtd:
        raise ValueError(
            f"No MTD sheet found in template.\nSheets: {names}\n"
            "Set MTD_SHEET_NAME at top of variance_engine.py.")
    if not ytd:
        raise ValueError(
            f"No YTD sheet found in template.\nSheets: {names}\n"
            "Set YTD_SHEET_NAME at top of variance_engine.py.")
    return mtd, ytd


# ═══════════════════════════════════════════════════════════════
#  STANDALONE TEST
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    test_args = {
        "scenario_pairs":  [("FC 2+10", "Actuals"), ("Budget", "LY Actuals")],
        "quarter":         "Q1 (Apr-Jun)",
        "month":           "April",
        "input_folders":   [],
        "input_files":     [],
        "master_file":     r"C:\Test\Variance_Master.xlsx",
        "template_file":   r"C:\Test\Variance_Template.xlsx",
        "output_folder":   r"C:\Test\Output",
        "timestamp":       datetime.now().strftime("%Y%m%d_%H%M%S"),
    }
    out = run_variance(test_args)
    print(f"Saved to: {out}")
