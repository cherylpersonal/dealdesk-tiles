import streamlit as st
import pandas as pd
from core.parser import parse_daf, parse_cost_sheet
from core.validators import validate_daf, validate_cost_sheet, check_missing_skus
from core.calculator import merge_data, calculate_margins
from core.approval import evaluate_approval
from core.exporter import generate_output_file

st.set_page_config(
    page_title="DAF Processor",
    page_icon="📋",
    layout="wide",
)

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.approval-badge {
    display: inline-block;
    padding: 8px 20px;
    border-radius: 6px;
    font-size: 20px;
    font-weight: 600;
    margin-bottom: 8px;
}
.level-dd  { background: #D4EDDA; color: #155724; }
.level-pm  { background: #FFF3CD; color: #856404; }
.level-ceo { background: #F8D7DA; color: #721C24; }
.reason-item { font-size: 14px; color: #555; margin: 4px 0; }
</style>
""", unsafe_allow_html=True)


# ── Header ────────────────────────────────────────────────────────────────────
st.title("📋 DAF Processor")
st.caption("Upload a Discount Approval Form and Cost Sheet to calculate margins and get an approval recommendation.")

st.divider()

# ── File Uploads ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.subheader("Upload 1 — DAF")
    daf_file = st.file_uploader(
        "Drag or browse your DAF Excel file",
        type=["xlsx"],
        key="daf",
    )
    if daf_file is None:
        st.caption("Expected columns: `daf_ref_no`, `lob`, `channel`, `state`, `project_name`, `developer_name`, `dealer_name`, `zonal_coordinator`, `submitted_date`, `sku_code`, `boxes`, `nef`, `list_value`")

with col2:
    st.subheader("Upload 2 — Cost Sheet")
    cost_file = st.file_uploader(
        "Drag or browse your Cost Sheet Excel file",
        type=["xlsx"],
        key="cost",
    )
    if cost_file is None:
        st.caption("Expected columns: `sku_code`, `area_per_box`, `buying_price`")

st.divider()


# ── Processing ────────────────────────────────────────────────────────────────
if not (daf_file and cost_file):
    st.info("Upload both files above to begin processing.")
    st.stop()

with st.spinner("Reading files…"):
    daf_df, daf_warnings = parse_daf(daf_file)
    cost_df, cost_warnings = parse_cost_sheet(cost_file)

# Show parse warnings
all_warnings = daf_warnings + cost_warnings
if all_warnings:
    for w in all_warnings:
        st.warning(f"⚠️ {w}")

# Validate
daf_ok, daf_errors = validate_daf(daf_df)
cost_ok, cost_errors = validate_cost_sheet(cost_df)

if not daf_ok:
    for e in daf_errors:
        st.error(f"DAF error: {e}")
    st.stop()

if not cost_ok:
    for e in cost_errors:
        st.error(f"Cost Sheet error: {e}")
    st.stop()

# Check missing SKUs
missing_skus = check_missing_skus(daf_df, cost_df)
if missing_skus:
    st.warning(
        f"⚠️ These SKUs from the DAF are not in the cost sheet and will be excluded: "
        f"{missing_skus}"
    )

# Merge + Calculate
with st.spinner("Calculating margins…"):
    merged_df, merge_warnings = merge_data(daf_df, cost_df)
    for w in merge_warnings:
        st.warning(f"⚠️ {w}")

    if merged_df.empty:
        st.error("No matching SKUs found. Check that sku_code values match between both files.")
        st.stop()

    calc_df, totals = calculate_margins(merged_df)
    approval = evaluate_approval(calc_df, totals)


# ── Approval Result ───────────────────────────────────────────────────────────
st.subheader("Approval Recommendation")

level_class = {
    "Deal Desk": "level-dd",
    "PM Head": "level-pm",
    "CEO": "level-ceo",
}[approval.level]

st.markdown(
    f'<div class="approval-badge {level_class}">{approval.level}</div>',
    unsafe_allow_html=True,
)
for reason in approval.reasons:
    st.markdown(f'<p class="reason-item">→ {reason}</p>', unsafe_allow_html=True)

st.divider()

# ── Summary Metrics ───────────────────────────────────────────────────────────
st.subheader("Deal Summary")
m1, m2, m3, m4 = st.columns(4)
m1.metric("Total Revenue", f"₹{totals['total_revenue']:,.0f}")
m2.metric("Total Cost", f"₹{totals['total_cost']:,.0f}")
m3.metric("Total Margin", f"₹{totals['total_margin_value']:,.0f}")
m4.metric(
    "Overall Margin %",
    f"{totals['overall_margin_percent']*100:.1f}%",
    delta=None,
)

st.divider()

# ── Per-SKU Table ─────────────────────────────────────────────────────────────
st.subheader("Per-SKU Calculations")

display_df = calc_df[[
    "sku_code", "boxes", "nef", "area_per_box", "buying_price",
    "total_area", "revenue", "cost", "margin_value", "margin_percent",
]].copy()

# Formatting for display
display_df["margin_percent_display"] = (
    display_df["margin_percent"] * 100
).round(2).astype(str) + "%"

# Colour-code margin column
def highlight_margin(row):
    mp = row.get("margin_percent", 1)
    if mp is None:
        return [""] * len(row)
    if mp < 0.05:
        color = "background-color: #F8D7DA"
    elif mp < 0.10:
        color = "background-color: #FFF3CD"
    else:
        color = "background-color: #D4EDDA"
    styles = [""] * len(row)
    if "margin_percent" in row.index:
        styles[row.index.get_loc("margin_percent")] = color
    return styles

styled = display_df.style.apply(highlight_margin, axis=1).format({
    "nef": "₹{:.2f}",
    "area_per_box": "{:.2f} sqft",
    "buying_price": "₹{:.2f}",
    "total_area": "{:,.2f}",
    "revenue": "₹{:,.2f}",
    "cost": "₹{:,.2f}",
    "margin_value": "₹{:,.2f}",
    "margin_percent": "{:.2%}",
})
st.dataframe(styled, use_container_width=True, hide_index=True)

# Total row
total_row_display = pd.DataFrame([{
    "sku_code": "TOTAL",
    "boxes": calc_df["boxes"].sum(),
    "nef": "—",
    "area_per_box": "—",
    "buying_price": "—",
    "total_area": f"{calc_df['total_area'].sum():,.2f}",
    "revenue": f"₹{totals['total_revenue']:,.2f}",
    "cost": f"₹{totals['total_cost']:,.2f}",
    "margin_value": f"₹{totals['total_margin_value']:,.2f}",
    "margin_percent": f"{totals['overall_margin_percent']*100:.2f}%",
}])
st.dataframe(total_row_display, use_container_width=True, hide_index=True)

st.divider()

# ── Tracker Preview ───────────────────────────────────────────────────────────
with st.expander("Tracker Row Preview"):
    from core.exporter import build_tracker_row
    tracker = build_tracker_row(calc_df, totals, approval)
    tracker_df = pd.DataFrame([tracker])
    st.dataframe(tracker_df.T.rename(columns={0: "Value"}), use_container_width=True)
    st.caption("Fields left blank require manual entry (Response Date, TAT, Approver, etc.)")

st.divider()

# ── Download ──────────────────────────────────────────────────────────────────
st.subheader("Download Output")

with st.spinner("Generating Excel output…"):
    output_bytes = generate_output_file(calc_df, totals, approval)

daf_ref = ""
if "daf_ref_no" in calc_df.columns:
    daf_ref = str(calc_df["daf_ref_no"].iloc[0]).replace("/", "-")
filename = f"DAF_Processed_{daf_ref or 'output'}.xlsx"

st.download_button(
    label="⬇️ Download Processed DAF (.xlsx)",
    data=output_bytes,
    file_name=filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
st.caption("Output contains two sheets: **Calculations** (with total row + approval block) and **Tracker** (structured row for your tracking sheet).")
