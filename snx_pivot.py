def build_r2_sheet2_pivot(df_sheet1: pd.DataFrame) -> pd.DataFrame:
    df = df_sheet1.copy()

    df[R2_COL_DRUG] = df[R2_COL_DRUG].astype(str).str.strip()
    df[R2_COL_STATUS] = df[R2_COL_STATUS].apply(r2_clean_text)
    df[R2_COL_REASON] = df[R2_COL_REASON].apply(r2_clean_text)

    # keep case_id as text (don’t title-case it)
    df[R2_COL_CASE_ID] = df[R2_COL_CASE_ID].astype(str).str.strip()

    # numeric versions of buckets for sums
    for name, _, _ in R2_BUCKETS:
        df[name] = df[name].apply(lambda v: 1 if v == 1 else 0)

    pivot = pd.pivot_table(
        df,
        index=[R2_COL_DRUG, R2_COL_STATUS, R2_COL_REASON, R2_COL_CASE_ID],  # ✅ added
        values=[b[0] for b in R2_BUCKETS],
        aggfunc="sum",
        fill_value=0
    ).sort_index(level=[0, 1, 2, 3], ascending=[True, True, True, True]).reset_index()

    return pivot
