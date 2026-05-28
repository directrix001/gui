    sheet_new  = "New KPI Actual"
    sheet_pl   = "P&L Actual"
    sheet_act  = "Act"
 
    formula_act_1  = _safe_gen_formula_1(34, col_new_kpi, sheet_new, result) if (HAS_FILE_RESULT and col_new_kpi) else Z
    formula_act_2  = _fhmf("wholesale_volume")
    formula_act_3  = _fhmf("gross_sales_wo_rd")
    formula_act_4  = _fhmf("rd_revenues")
    formula_act_5  = _fhmf("ci")
    formula_act_6  = _fhmf("mfg_fixed_plants")
    formula_act_7  = _fhmf("vendor_tooling")
    formula_act_8  = _fhmf("rd_costs")
    formula_act_9  = _fhmf("mfg_fixed_out_of_scope")
    formula_act_10 = _fhmf("fmi")
    formula_act_11 = _fhmf("ga_and_other_items")
 
    sga_row = _sr("sga_total")
    formula_act_12 = (
        _safe_gen_formula(sga_row + 1, col_hmf, sheet_man, file_path_HMF)
        if (HAS_HMF and HAS_FILE_HMF and col_hmf and sga_row) else Z
    )
    formula_act_13 = _fhmf("warranty_rs")
