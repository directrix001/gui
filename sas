Just replace every `_fhmf(key)` call and the `_sr("sga_total")` lookup with direct row numbers. Here's the pattern:

**Before:**
```python
formula_act_2  = _fhmf("wholesale_volume")
formula_act_3  = _fhmf("gross_sales_wo_rd")
...
sga_row = _sr("sga_total")
formula_act_12 = (
    _safe_gen_formula(sga_row + 1, col_hmf, sheet_man, file_path_HMF)
    if (HAS_HMF and HAS_FILE_HMF and col_hmf and sga_row) else Z
)
```

**After:**
```python
formula_act_1  = _safe_gen_formula_1(34,  col_new_kpi, sheet_new, result) if (HAS_FILE_RESULT and col_new_kpi) else Z
formula_act_2  = _safe_gen_formula(ROW_WHOLESALE_VOL,       col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_3  = _safe_gen_formula(ROW_GROSS_SALES_WO_RD,   col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_4  = _safe_gen_formula(ROW_RD_REVENUES,         col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_5  = _safe_gen_formula(ROW_CI,                  col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_6  = _safe_gen_formula(ROW_MFG_FIXED_PLANTS,    col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_7  = _safe_gen_formula(ROW_VENDOR_TOOLING,      col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_8  = _safe_gen_formula(ROW_RD_COSTS,            col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_9  = _safe_gen_formula(ROW_MFG_FIXED_OOS,       col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_10 = _safe_gen_formula(ROW_FMI,                 col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_11 = _safe_gen_formula(ROW_GA_AND_OTHER,        col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_12 = _safe_gen_formula(ROW_SGA_TOTAL + 1,       col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
formula_act_13 = _safe_gen_formula(ROW_WARRANTY_RS,         col_hmf, sheet_man, file_path_HMF) if (HAS_HMF and HAS_FILE_HMF and col_hmf) else Z
```

Then define the constants at the top of the function (or as module-level constants), filling in your actual row numbers:

```python
# ── Hardcoded HMF row indices (no searching) ──────────────────────────────────
ROW_WHOLESALE_VOL     = 5   # replace with real row
ROW_GROSS_SALES_WO_RD = 8   # replace with real row
ROW_RD_REVENUES       = 12  # replace with real row
ROW_CI                = 15  # replace with real row
ROW_MFG_FIXED_PLANTS  = 20  # replace with real row
ROW_VENDOR_TOOLING    = 24  # replace with real row
ROW_RD_COSTS          = 28  # replace with real row
ROW_MFG_FIXED_OOS     = 31  # replace with real row
ROW_FMI               = 35  # replace with real row
ROW_GA_AND_OTHER      = 40  # replace with real row
ROW_SGA_TOTAL         = 44  # replace with real row (+1 applied inline above)
ROW_WARRANTY_RS       = 50  # replace with real row
```

You also no longer need the `sga_row` variable or the `sga_row` guard in the condition — since the row is a non-zero constant, the `if (HAS_HMF and HAS_FILE_HMF and col_hmf)` guard is sufficient.
