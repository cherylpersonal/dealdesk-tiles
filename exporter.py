import io
import pandas as pd
from datetime import datetime, timedelta
from .approval import ApprovalResult


CALC_COLS = ["total_area", "revenue", "cost", "margin_value", "margin_percent"]

DISPLAY_COL_ORDER = [
    "sku_code", "boxes", "nef", "area_per_box", "buying_price",
    "total_area", "revenue", "cost", "margin_value", "margin_percent",
]

META_COLS = [
    "daf_ref_no", "lob", "channel", "state", "project_name",
    "developer_name", "dealer_name", "zonal_coordinator",
    "submitted_date", "list_value",
]


def generate_output_file(
    df: pd.DataFrame,
    totals: dict,
    approval: ApprovalResult,
) -> bytes:
    """
    Build processed DAF as Excel bytes.
    Sheet 1: Calculations with total row + approval block.
    Sheet 2: Tracker row.
    """
    output = io.BytesIO()

    # Build display dataframe — show available columns in order
    display_cols = [c for c in DISPLAY_COL_ORDER if c in df.columns]
    calc_df = df[display_cols].copy()

    # Format margin_percent as fraction (Excel will format as %)
    pct_col = "margin_percent"

    # Total row
    total_row = {c: "" for c in display_cols}
    total_row["sku_code"] = "TOTAL"
    if "boxes" in total_row:
        total_row["boxes"] = df["boxes"].sum()
    if "total_area" in total_row:
        total_row["total_area"] = df["total_area"].sum()
    if "revenue" in total_row:
        total_row["revenue"] = totals["total_revenue"]
    if "cost" in total_row:
        total_row["cost"] = totals["total_cost"]
    if "margin_value" in total_row:
        total_row["margin_value"] = totals["total_margin_value"]
    if pct_col in total_row:
        total_row[pct_col] = totals["overall_margin_percent"]

    total_df = pd.DataFrame([total_row])
    final_df = pd.concat([calc_df, total_df], ignore_index=True)

    tracker_row = build_tracker_row(df, totals, approval)
    tracker_df = pd.DataFrame([tracker_row])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, sheet_name="Calculations", index=False)
        tracker_df.to_excel(writer, sheet_name="Tracker", index=False)

        wb = writer.book
        calc_ws = writer.sheets["Calculations"]
        tracker_ws = writer.sheets["Tracker"]

        # Formats
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        currency_fmt = wb.add_format({"num_format": "#,##0.00"})
        header_fmt = wb.add_format({"bold": True, "bg_color": "#E8F0FE", "border": 1})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#FFF3CD", "border": 1})
        pct_total_fmt = wb.add_format(
            {"bold": True, "bg_color": "#FFF3CD", "border": 1, "num_format": "0.00%"}
        )

        # Approval level colour
        level_colors = {
            "Deal Desk": "#D4EDDA",
            "PM Head": "#FFF3CD",
            "CEO": "#F8D7DA",
        }
        approval_fmt = wb.add_format(
            {
                "bold": True,
                "bg_color": level_colors.get(approval.level, "#FFFFFF"),
                "border": 1,
                "font_size": 12,
            }
        )

        # Write approval block below data
        data_rows = len(final_df) + 1  # +1 for header
        approval_start = data_rows + 2

        calc_ws.write(approval_start, 0, "Recommended Approval Level", header_fmt)
        calc_ws.write(approval_start, 1, approval.level, approval_fmt)
        calc_ws.write(approval_start + 1, 0, "Reason(s)", header_fmt)
        calc_ws.write(approval_start + 1, 1, " | ".join(approval.reasons))

        # Column widths + number formats for calc sheet
        col_map = {c: i for i, c in enumerate(display_cols)}
        for col_name, width, fmt in [
            ("sku_code", 18, None),
            ("boxes", 10, None),
            ("nef", 12, currency_fmt),
            ("area_per_box", 14, currency_fmt),
            ("buying_price", 14, currency_fmt),
            ("total_area", 14, currency_fmt),
            ("revenue", 16, currency_fmt),
            ("cost", 16, currency_fmt),
            ("margin_value", 16, currency_fmt),
            ("margin_percent", 16, pct_fmt),
        ]:
            if col_name in col_map:
                idx = col_map[col_name]
                calc_ws.set_column(idx, idx, width, fmt)

        # Format total row
        total_row_idx = len(calc_df) + 1  # +1 for header row (0-indexed)
        for col_name, col_idx in col_map.items():
            fmt = pct_total_fmt if col_name == "margin_percent" else total_fmt
            calc_ws.write(total_row_idx, col_idx, total_row.get(col_name, ""), fmt)

        # Auto-width tracker sheet columns
        for i, col in enumerate(tracker_df.columns):
            calc_ws  # already handled
            tracker_ws.set_column(i, i, max(len(str(col)) + 4, 16))

    output.seek(0)
    return output.read()


def build_tracker_row(
    df: pd.DataFrame,
    totals: dict,
    approval: ApprovalResult,
) -> dict:
    """
    Build one tracker row dict from available data.
    Fields that need manual entry are left empty string.
    """
    # Pull meta from first row of df (all rows share the same DAF meta)
    meta = df.iloc[0] if not df.empty else {}

    def get(col):
        return meta.get(col, "") if hasattr(meta, "get") else ""

    submitted = get("submitted_date")
    try:
        submitted_dt = pd.to_datetime(submitted)
        six_hr_due = submitted_dt + timedelta(hours=6)
        six_hr_str = six_hr_due.strftime("%Y-%m-%d %H:%M")
        submitted_str = submitted_dt.strftime("%Y-%m-%d")
    except Exception:
        six_hr_str = ""
        submitted_str = str(submitted)

    list_value = df["list_value"].sum() if "list_value" in df.columns else ""
    deal_value = totals["total_revenue"]
    avg_discount = (
        (float(list_value) - deal_value) / float(list_value)
        if list_value and float(list_value) > 0
        else ""
    )

    return {
        "DAF Reference No.": get("daf_ref_no"),
        "LOB": get("lob"),
        "Channel": get("channel"),
        "State": get("state"),
        "Project Name": get("project_name"),
        "Developer Name": get("developer_name"),
        "Dealer Name": get("dealer_name"),
        "Zonal Coordinator": get("zonal_coordinator"),
        "Submitted Date": submitted_str,
        "DD Received Date": "",          # manual
        "6-Hr Due Date": six_hr_str,
        "Response Date": "",             # manual
        "TAT": "",                       # manual (filled after response)
        "SLA Met": "",                   # manual / derived from TAT
        "Quote Version": "",             # manual
        "Approval Level": approval.level,
        "Approver Name": "",             # manual
        "Approval Status": "",           # manual
        "Approval Date": "",             # manual
        "List Value": list_value,
        "Deal Value": round(deal_value, 2),
        "Avg Discount %": (
            f"{avg_discount*100:.2f}%" if isinstance(avg_discount, float) else ""
        ),
        "Deal Margin %": f"{totals['overall_margin_percent']*100:.2f}%",
    }
