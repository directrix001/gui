# Fix: broken external links / "recover links" dialog, duplicate outputs, slow runs

Same output as before (identical formulas, cells, sheet names, and output file
names) — only the *mechanics* of reading/writing the workbooks changed.

## Root causes found

1. **Cross-workbook links weren't registered (the "recover links" dialog).**
   The engine writes external references as raw formula strings, e.g.
   `='[C:/…/HFM Managerial.xlsx]Managerial'!V8`. openpyxl stores that text
   verbatim but creates **no** `externalReferences` entry and **no**
   `xl/externalLinks/` parts. Excel then has to *repair* the file on open to
   rebuild those links — that's the recovery prompt — and the links often end
   up broken.

2. **NaN / ±inf written into the injected "Master" sheet.**
   `copy_master_sheet` wrote DataFrame values straight into cells. Any NaN/inf
   becomes an invalid `<v></v>` numeric node, which *also* triggers Excel's
   "we found a problem / recover" dialog. This runs on the file that becomes
   the output.

3. **Duplicate/overwriting outputs.** Every output was saved to 2–4 places
   (`P_and_L Actuals Files/`, `New KPI Files/`, `Mandatory_folder/`,
   `Goodwill_files/`) and `copy_excel_file` appended `_1`, `_2`, `_3` … to the
   Temp copy on every run.

4. **Slow.** `extract_formula` re-opened and fully re-parsed the same template
   ~24× per run, and each workbook was serialized to disk 2× (duplicate saves).

## What changed

### New file: `link_repair.py`
`repair_external_links(path)` registers every raw-string external reference
into a proper OOXML external-link table and rewrites the in-cell formulas to
Excel's native indexed form (`='[1]Sheet'!A1`). Excel then opens the file with
**no repair dialog** and **working links**.
It is: append-only (never touches links already present in the template),
idempotent, and defensive — on any error it restores the original file
untouched, so it can never make output worse than before.

### `Submit_2_Helper_function.py`
- `paste_values_KPI_PL`: saves **once** to the output folder (dropped the
  `P_and_L Actuals Files/` duplicate) and calls `repair_external_links`.
- `extract_formula`: now uses a cached workbook load → the template is parsed
  once instead of ~18×.

### `Nissan_Helper_act_function.py`
- `paste_values_KPI_PL`: added `keep_links=True`, saves **once** to the output
  folder (dropped the dead `New KPI Files/` + `Mandatory_folder/` copies),
  repairs links. (Also uses `os.path.join` instead of a hard-coded `\`.)
- `paste_values_goodwill`: added `keep_links=True`, saves once + repairs links,
  then keeps the single `Goodwill_files/` copy (that one **is** required — it's
  read back as next month's goodwill template) by byte-copying the repaired
  file instead of re-serializing.
- `extract_formula`: cached workbook load (parsed once instead of ~6×).

### `ui_main.py`
- `copy_master_sheet`: sanitizes NaN/±inf → blank and coerces numpy scalars
  before writing the Master sheet (fixes cause #2). Also faster (`append`).
- `copy_excel_file`: overwrites the working copy in place instead of spawning
  `…_1`, `…_2` duplicates each run (falls back to a timestamped name only if a
  stale copy is locked/open).

## How to verify
Open a generated output in Excel — there should be no "recover"/repair prompt,
and **Data ▸ Edit Links** should list the input workbooks with status **OK**
(pointing at the real input file paths). Re-running should no longer pile up
`_1/_2` copies or extra output folders, and the run should be noticeably faster.

## Deployment note
Keep `link_repair.py` in the same folder as the other modules (it's imported as
`from link_repair import repair_external_links`). No new dependencies — it uses
only the Python standard library.
