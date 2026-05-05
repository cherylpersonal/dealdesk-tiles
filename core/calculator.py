import pandas as pd
from typing import Tuple, List


def merge_data(
    daf_df: pd.DataFrame, cost_df: pd.DataFrame
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Left-join DAF with cost sheet on sku_code.
    Returns merged dataframe and list of SKUs missing from cost sheet.
    """
    warnings = []
    cost_cols = cost_df[["sku_code", "area_per_box", "buying_price"]]
    merged = daf_df.merge(cost_cols, on="sku_code", how="left")

    missing = merged[merged["area_per_box"].isna()]["sku_code"].tolist()
    if missing:
        warnings.append(
            f"SKUs not found in cost sheet (excluded from calculations): {missing}"
        )
        merged = merged[merged["area_per_box"].notna()].copy()

    return merged, warnings


def calculate_margins(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """
    Add calculated columns to merged dataframe.
    Returns (df_with_calculations, totals_dict).
    """
    df = df.copy()

    df["total_area"] = df["boxes"] * df["area_per_box"]
    df["revenue"] = df["nef"] * df["total_area"]
    df["cost"] = df["buying_price"] * df["total_area"]
    df["margin_value"] = df["revenue"] - df["cost"]
    df["margin_percent"] = df["margin_value"] / df["revenue"]

    # Replace inf/-inf from zero-revenue rows
    df["margin_percent"] = df["margin_percent"].replace(
        [float("inf"), float("-inf")], None
    )

    totals = {
        "total_revenue": df["revenue"].sum(),
        "total_cost": df["cost"].sum(),
        "total_margin_value": df["margin_value"].sum(),
        "overall_margin_percent": (
            df["margin_value"].sum() / df["revenue"].sum()
            if df["revenue"].sum() > 0
            else 0
        ),
    }

    return df, totals
