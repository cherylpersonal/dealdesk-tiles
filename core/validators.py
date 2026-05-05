import pandas as pd
from typing import Tuple, List

DAF_REQUIRED_COLS = ["sku_code", "boxes", "nef"]
DAF_META_COLS = [
    "daf_ref_no", "lob", "channel", "state", "project_name",
    "developer_name", "dealer_name", "zonal_coordinator", "submitted_date"
]
COST_REQUIRED_COLS = ["sku_code", "area_per_box", "buying_price"]


def validate_daf(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    errors = []
    missing = [c for c in DAF_REQUIRED_COLS if c not in df.columns]
    if missing:
        errors.append(f"Missing required columns: {missing}")
        return False, errors

    for col in ["boxes", "nef"]:
        if col in df.columns:
            bad = df[df[col] <= 0]
            if not bad.empty:
                errors.append(
                    f"Column '{col}' has zero or negative values in rows: "
                    f"{bad.index.tolist()}"
                )

    dups = df[df.duplicated("sku_code", keep=False)]["sku_code"].unique()
    if len(dups) > 0:
        errors.append(
            f"Duplicate SKU codes in DAF: {list(dups)}. "
            "Rows will be aggregated (summed) per SKU."
        )

    return len(errors) == 0 or all("Duplicate" in e for e in errors), errors


def validate_cost_sheet(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    errors = []
    missing = [c for c in COST_REQUIRED_COLS if c not in df.columns]
    if missing:
        errors.append(f"Missing required columns: {missing}")
        return False, errors

    for col in ["area_per_box", "buying_price"]:
        if col in df.columns:
            bad = df[df[col] <= 0]
            if not bad.empty:
                errors.append(
                    f"Column '{col}' has zero or negative values for SKUs: "
                    f"{df.loc[bad.index, 'sku_code'].tolist()}"
                )

    return len(errors) == 0, errors


def check_missing_skus(daf_df: pd.DataFrame, cost_df: pd.DataFrame) -> List[str]:
    daf_skus = set(daf_df["sku_code"].str.strip().str.upper())
    cost_skus = set(cost_df["sku_code"].str.strip().str.upper())
    return list(daf_skus - cost_skus)
