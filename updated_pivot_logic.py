def clean_reason(text):
    if pd.isna(text):
        return "Unknown"
    text = str(text).strip()
    text = " ".join(text.split())  # remove extra spaces
    text = text.lower()            # normalize case
    return text.title()            # convert to Title Case


def build_excel_like_pivot(df: pd.DataFrame) -> pd.DataFrame:

    # Normalize reason text
    df[COL_REASON] = df[COL_REASON].apply(clean_reason)

    # Detail rows (sorted alphabetically by reason)
    detail = (
        df.groupby([COL_DRUG, COL_REASON], as_index=False)[COL_COUNT]
        .sum()
        .sort_values([COL_DRUG, COL_REASON], ascending=[True, True])
    )

    # Drug totals
    totals = (
        df.groupby(COL_DRUG, as_index=False)[COL_COUNT]
        .sum()
        .sort_values(COL_DRUG, ascending=True)
    )

    rows = []
    grand_total = df[COL_COUNT].sum()

    for _, t in totals.iterrows():
        drug = t[COL_DRUG]
        drug_total = int(t[COL_COUNT])

        # Drug header row
        rows.append([drug, "", drug_total])

        # Sorted reasons under drug
        sub = detail[detail[COL_DRUG] == drug]
        for _, r in sub.iterrows():
            rows.append(["", "  " + str(r[COL_REASON]), int(r[COL_COUNT])])

        rows.append(["", "", ""])

    rows.append(["Grand Total", "", int(grand_total)])

    return pd.DataFrame(rows, columns=[COL_DRUG, COL_REASON, COL_COUNT])
