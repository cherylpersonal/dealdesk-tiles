import pandas as pd
from typing import Tuple, List


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace and lowercase all column names."""
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    return df


def parse_daf(file) -> Tuple[pd.DataFrame, List[str]]:
    """
    Read DAF Excel file.
    Returns (dataframe, warnings_list).
    Strips whitespace from sku_code and normalises it to uppercase.
    Aggregates duplicate SKUs by summing boxes and averaging nef (weighted).
    """
    warnings = []
    df = pd.read_excel(file, dtype=str)
    df = _normalise_columns(df)

    numeric_cols = ["boxes", "nef", "list_value"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "sku_code" in df.columns:
        df["sku_code"] = df["sku_code"].str.strip().str.upper()

    # Handle duplicate SKUs: sum boxes, weighted average nef
    if "sku_code" in df.columns and df.duplicated("sku_code").any():
        warnings.append(
            "Duplicate SKU codes detected — aggregating: boxes summed, "
            "NEF weighted-averaged by boxes."
        )
        df["_weighted_nef"] = df["nef"] * df["boxes"]
        agg_dict = {"boxes": "sum", "_weighted_nef": "sum"}

        # carry forward first value for all meta/string columns
        meta_cols = [
            c for c in df.columns
            if c not in ["sku_code", "boxes", "nef", "_weighted_nef"]
        ]
        for c in meta_cols:
            agg_dict[c] = "first"

        df = df.groupby("sku_code", as_index=False).agg(agg_dict)
        df["nef"] = df["_weighted_nef"] / df["boxes"]
        df.drop(columns=["_weighted_nef"], inplace=True)

    return df, warnings


def parse_cost_sheet(file) -> Tuple[pd.DataFrame, List[str]]:
    """
    Read cost sheet Excel file.
    Returns (dataframe, warnings_list).
    """
    warnings = []
    df = pd.read_excel(file, dtype=str)
    df = _normalise_columns(df)

    for col in ["area_per_box", "buying_price"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "sku_code" in df.columns:
        df["sku_code"] = df["sku_code"].str.strip().str.upper()

    if "sku_code" in df.columns and df.duplicated("sku_code").any():
        warnings.append(
            "Duplicate SKU codes in cost sheet — keeping first occurrence per SKU."
        )
        df = df.drop_duplicates("sku_code", keep="first")

    return df, warnings
