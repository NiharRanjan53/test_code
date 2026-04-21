def apply_workscope_inclusion_exclusion(
    result: dict,
    main_df: pd.DataFrame,
    workscope: str,
    pipeline_type: str, #material and repair and Vendor
    remarks: str
    ) -> pd.DataFrame:
    
    main_df.to_excel("main_df_inwks_inc_exc.xlsx", index=False)
    print(f"1 - {main_df.shape}")
    df = main_df.copy()
    df.columns = [c.strip().lower() for c in df.columns]
    print("df cols",pipeline_type,df.columns)
   
    # =======================================================
    # Load WSPG Inclusion List
    # =======================================================
    wspg_df = get_dataframe_by_file_and_sheet(
            all_dataframes=result["dataframes"],
            file_key="contract",
            sheet_name_substring="wspg inclusion list"
        ).copy()
    wspg_df.columns = [c.strip().lower() for c in wspg_df.columns]
    part_no_col = next((c for c in wspg_df.columns if "part no" in c), None)
    wspg_inclusion_col = next((c for c in wspg_df.columns if "wspg inclusion" in c), None)

    wspg_df[wspg_inclusion_col] = wspg_df[wspg_inclusion_col].str.lower().str.split().str.join(' ')
    
    if not all([part_no_col, wspg_inclusion_col]):
        raise ValueError("Required column not found in WSPG Inclusion List DataFrame.")
    
    def fill_suffixes_from_right(text):
        parts = [p.strip() for p in text.split(',')]
        for i in range(len(parts) - 2, -1, -1):
            if not parts[i].endswith(('Min', 'Perf')):
                right_neighbor = parts[i+1]
                if "Perf" in right_neighbor:
                    parts[i] = f"{parts[i]} Perf"
                else:
                    parts[i] = f"{parts[i]} Min"
        return ", ".join(parts)

    wspg_df[wspg_inclusion_col] = wspg_df[wspg_inclusion_col].apply(fill_suffixes_from_right)

    # =======================================================
    # Load WORKSCOPE Level
    # =======================================================
    wks_level_df = get_dataframe_by_file_and_sheet(
            all_dataframes=result["dataframes"],
            file_key="contract",
            sheet_name_substring="workscope level"
    ).copy()
    
    wks_level_df.columns = [c.strip().lower() for c in wks_level_df.columns]
    
    ws_idefnition_col = next((c for c in wks_level_df.columns if "wokscope defnitions" in c), None)
    workscope_col = next((c for c in wks_level_df.columns if workscope.lower() in c), None)
    workscope = workscope.lower()
    
    if not all([ws_idefnition_col, workscope_col]):
        raise ValueError("Required column not found in Workscope Level DataFrame.")
    
    wks_level_df = wks_level_df[["wokscope defnitions", workscope]]

    wks_level_df = wks_level_df.replace('', np.nan).dropna(subset=[workscope])
    wks_level_df['cs_module'] = wks_level_df[ws_idefnition_col].str.extract(r'\d{2}-(.*?)\s')
    wks_level_df[workscope] = pd.to_numeric(wks_level_df[workscope], errors='coerce')
    wks_level_df = wks_level_df.dropna(subset=[workscope, 'cs_module'])
    wks_level_df[workscope] = wks_level_df[workscope].astype(int)
    wks_level_df = wks_level_df[['cs_module', workscope]]
   
    # =======================================================
    # Load SVR
    # =======================================================
    svr_df = get_dataframe_by_file_and_sheet(
            all_dataframes=result["dataframes"],
            file_key="svr",
            sheet_name_substring="Sheet1"
    ).copy()

    svr_df = svr_df.drop_duplicates()
    svr_df.columns = svr_df.columns.astype(str).str.strip().str.lower()

    to_keep = ["module", "initial workscope", "final workscope"]
    svr_df = svr_df[to_keep].dropna(how='all')

    for col in ["initial workscope", "final workscope"]:
        svr_df[col] = svr_df[col].astype(str).str.extract(r'(\d+)')

    # =======================================================
    # Merge
    # =======================================================
    main_df = main_df.merge(
        wks_level_df[['cs_module', workscope]].rename(columns={workscope: 'initial workscope'}), 
        left_on='module',
        right_on='cs_module',
        how='left'
    )

    main_df = main_df.merge(
        svr_df[["module", "final workscope"]].rename(columns={"module": "svr_module"}), 
        left_on='module', 
        right_on='svr_module', 
        how='left'
    )

    main_df = main_df.merge(
        wspg_df[[part_no_col, wspg_inclusion_col]],
        left_on='part_nbr', 
        right_on=part_no_col, 
        how='left'
    )

    # =======================================================
    # ATA
    # =======================================================
    main_df["ata4"] = main_df["ata_long"].astype(str).str.extract(r"(\d{4})", expand=False)
    main_df["ata4"] = main_df["ata4"].astype("Int64")

    main_df["covered_status"] = "No WSPG match"

    # =======================================================
    # MASK
    # =======================================================
    if pipeline_type == "material":
        target_mask = main_df["remarks"].astype(str).str.upper().eq("MATL UNKNOWN")
    elif pipeline_type == "vendor":
        target_mask = main_df["remarks"].astype(str).str.upper().eq("VEND UNKNOWN")
    else:
        target_mask = main_df["remarks"].astype(str).str.upper().eq("REPAIR UNKNOWN")

    # =======================================================
    # RULE: ATA 72
    # =======================================================
    mask72 = target_mask & main_df["ata4"].between(7200, 7209, inclusive="both")
    main_df.loc[mask72, ["remarks", "covered_status"]] = ["Inclusion as 72", "Covered"]

    # =======================================================
    # RULE: WKS = 3
    # =======================================================
    if pipeline_type == "material":
        mask_full = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["initial workscope"] == 3)
    elif pipeline_type == "repair":
        mask_full = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["wks_num"] == "3")
    else:
        mask_full = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & ((main_df["wks_num"] == 3) | (main_df["final ata ws"] == 3))

    main_df.loc[mask_full, ["remarks", "covered_status"]] = ["Inclusion as WKS is full", "Covered"]

    # =======================================================
    # PREP
    # =======================================================
    main_df["module_clean"] = main_df["module"].fillna("UNKNOWN_MOD")
    main_df["ata4_clean"] = main_df["ata4"].fillna(0).astype(int)

    # =======================================================
    # ✅ FIX FUNCTION
    # =======================================================
    def check_inclusion_logic(module, wks, inclusion_text):
        module = str(module).strip()

        if pd.isna(module) or module == "nan":
            return None

        if pd.isna(inclusion_text):
            return True

        parts = [p.strip() for p in str(inclusion_text).split(",")]

        parsed = []
        for p in parts:
            tokens = p.split()
            if len(tokens) == 1:
                parsed.append([tokens[0], None])
            else:
                parsed.append([tokens[0], tokens[1]])

        for i in range(len(parsed) - 2, -1, -1):
            if parsed[i][1] is None:
                right = parsed[i + 1][1]
                parsed[i][1] = right if right else "Min"

        parsed_dict = {m: lvl for m, lvl in parsed}

        if module not in parsed_dict:
            return False

        lvl = parsed_dict[module]

        if wks == 1:
            return lvl == "Min"
        elif wks == 2:
            return lvl == "Perf"
        elif wks == 3:
            return False

        return False

    # =======================================================
    # WKS=1
    # =======================================================
    if pipeline_type == "material":
        mask1 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["initial workscope"] == 1)
    elif pipeline_type == "repair":
        mask1 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["wks_num"] == 1)
    else:
        mask1 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["wks_num"] == 1)

    is_included_1 = main_df.apply(
        lambda row: check_inclusion_logic(
            row["module_clean"],
            row["initial workscope"],
            row[wspg_inclusion_col]
        ),
        axis=1
    )

    module_not_nan = main_df["module"].notna()

    main_df.loc[mask1 & (is_included_1 == True) & module_not_nan, ["remarks", "covered_status"]] = ["Exclusion as WKS = 1", "Not Covered"]
    main_df.loc[mask1 & (is_included_1 == False) & module_not_nan, ["remarks", "covered_status"]] = ["Inclusion as WKS = 1", "Covered"]

    # =======================================================
    # WKS=2
    # =======================================================
    if pipeline_type == "material":
        mask2 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["initial workscope"] == 2)

        is_included_2 = main_df.apply(
            lambda row: check_inclusion_logic(
                row["module_clean"],
                row["initial workscope"],
                row[wspg_inclusion_col]
            ),
            axis=1
        )

        main_df.loc[mask2 & (is_included_2 == True) & module_not_nan, ["remarks", "covered_status"]] = ["Exclusion as WKS = 2", "Not Covered"]
        main_df.loc[mask2 & (is_included_2 == False) & module_not_nan, ["remarks", "covered_status"]] = ["Inclusion as WKS = 2", "Covered"]

    elif pipeline_type == "repair":
        mask2 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["wks_num"] == 2)
    else:
        mask2 = target_mask & main_df["ata4"].between(7221, 7263, inclusive="both") & (main_df["wks_num"] == 2)

    # =======================================================
    # Identifying the Adder Creep (ONLY for MATERIAL)
    # =======================================================
    # This logic is applied AFTER all Inclusion/Exclusion rules
    #
    # Purpose:
    # - Identify creep cases where workscope increases
    #
    # Conditions:
    # - ATA range = 7261 to 7263
    # - Initial Workscope = 1
    # - Final Workscope = 2 → "Creap 1 to 2"
    # - Final Workscope = 3 → "Creap 1 to 3"
    #
    # IMPORTANT:
    # - This WILL overwrite previous remarks (Inclusion/Exclusion)
    # - This is intentional as per current business logic

    if pipeline_type == "material":

        # Ensure final workscope is numeric (safety fix)
        main_df["final workscope"] = pd.to_numeric(main_df["final workscope"], errors="coerce")

        # -------------------------------
        # Creep 1 → 2
        # -------------------------------
        adder_mask_2 = (
            target_mask
            & main_df["ata4"].between(7261, 7263, inclusive="both")
            & (main_df['initial workscope'] == 1)
            & (main_df["final workscope"] == 2)
        )

        main_df.loc[adder_mask_2, ['remarks']] = ['Creap 1 to 2']

        # -------------------------------
        # Creep 1 → 3
        # -------------------------------
        adder_mask_3 = (
            target_mask
            & main_df["ata4"].between(7261, 7263, inclusive="both")
            & (main_df['initial workscope'] == 1)
            & (main_df["final workscope"] == 3)
        )

        main_df.loc[adder_mask_3, ['remarks']] = ['Creap 1 to 3']

    # =======================================================
    # FINAL
    # =======================================================
    main_df.drop(columns=["module_clean", "ata4_clean"], inplace=True)

    main_df.to_excel("apply_workscope_inclusion_exclusion_v4_mat.xlsx", index=False)

    return main_df