import pandas as pd
from dataclasses import dataclass
from typing import List


@dataclass
class ApprovalResult:
    level: str          # "Deal Desk" | "PM Head" | "CEO"
    reasons: List[str]  # human-readable trigger explanations
    sku_flags: pd.DataFrame  # per-SKU flag detail


def evaluate_approval(df: pd.DataFrame, totals: dict) -> ApprovalResult:
    """
    Apply approval matrix to calculated data.

    Matrix:
      Deal Desk : ALL SKU margins >= 10% AND overall >= 20%
      PM Head   : ALL SKU margins >= 5%  AND overall >= 10%
      CEO       : ANY SKU margin  <  5%  OR  overall <  10%
    """
    reasons = []

    # Per-SKU flags
    df = df.copy()
    df["_flag"] = "ok"
    df.loc[df["margin_percent"] < 0.10, "_flag"] = "below_10pct"
    df.loc[df["margin_percent"] < 0.05, "_flag"] = "below_5pct"

    skus_below_10 = df[df["margin_percent"] < 0.10]["sku_code"].tolist()
    skus_below_5 = df[df["margin_percent"] < 0.05]["sku_code"].tolist()
    overall = totals["overall_margin_percent"]

    sku_flags = df[["sku_code", "margin_percent", "_flag"]].copy()
    sku_flags["margin_percent_display"] = (
        sku_flags["margin_percent"] * 100
    ).round(2).astype(str) + "%"

    # Determine level
    if skus_below_5 or overall < 0.10:
        level = "CEO"
        if skus_below_5:
            reasons.append(
                f"SKU(s) with margin below 5%: {skus_below_5}"
            )
        if overall < 0.10:
            reasons.append(
                f"Overall deal margin {overall*100:.1f}% is below 10%"
            )

    elif skus_below_10 or overall < 0.20:
        level = "PM Head"
        if skus_below_10:
            reasons.append(
                f"SKU(s) with margin below 10%: {skus_below_10}"
            )
        if overall < 0.20:
            reasons.append(
                f"Overall deal margin {overall*100:.1f}% is below 20%"
            )

    else:
        level = "Deal Desk"
        reasons.append(
            f"All SKU margins >= 10% and overall margin "
            f"{overall*100:.1f}% >= 20%"
        )

    return ApprovalResult(level=level, reasons=reasons, sku_flags=sku_flags)
