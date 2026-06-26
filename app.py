"""
app.py
Streamlit UI for the ArcGIS Special Event Staff Converter.
Supports multiple file uploads — each file processed individually,
all rows combined into one output workbook.
"""

import os
import io
import traceback

import pandas as pd
import streamlit as st

from staff_transformer import (
    ARCGIS_COLUMNS,
    load_source_file,
    detect_source_format,
    find_current_staffing_header_row,
    reparse_staffing_from_raw,
    transform_special_event_workup,
    transform_current_staffing_report,
    validate_output,
    write_to_template,
    STAFFING_HEADER_IDENTIFIERS,
)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), "templates", "Sample_Batch_Load_Event_Staff_Template.xlsx"
)
OUTPUT_FILENAME = "ArcGIS_Special_Event_Staff_Upload.xlsx"

# ---------------------------------------------------------------------------
# Page setup
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="ArcGIS Special Event Staff Converter",
    page_icon="🚔",
    layout="wide",
)

st.title("ArcGIS Special Event Staff Converter")
st.markdown(
    """
    This tool converts HPD staffing export files into the format required by the
    **ArcGIS Special Event Solution** batch upload workbook.

    **Supported input files:**
    - `CurrentStaffingReport.csv` — exported from the HPD staffing system
    - `SpecialEventWorkupforGIS.csv` — exported from the Special Event Workup tool

    **How to use:**
    1. Upload one or more staffing files using the file uploader below.
    2. Review the input preview and validation warnings.
    3. Click **Download ArcGIS Upload Workbook** to get the completed Excel file.
    4. Import the downloaded file into ArcGIS using the batch upload process.
    """
)

# ---------------------------------------------------------------------------
# Sidebar settings
# ---------------------------------------------------------------------------

st.sidebar.header("⚙️ Settings")
st.sidebar.markdown(
    "These values are used as defaults when a field is blank or missing in the source file. "
    "Change them here before uploading if your event requires different values."
)

st.sidebar.markdown("**Default Field Values**")
default_unit_type_workup = st.sidebar.text_input(
    "Default Unit Type (SpecialEventWorkup)",
    value="Vehicle",
    help="Applied to unittype when the source UnitType column is blank.",
)
default_unit_type_staffing = st.sidebar.text_input(
    "Default Unit Type (CurrentStaffingReport)",
    value="Traffic Control",
    help="Applied to unittype for all CurrentStaffingReport rows.",
)
default_staff_status = st.sidebar.text_input(
    "Default Staff Status",
    value="On Duty",
    help="Applied to staffstatus when the source StaffStatus column is blank.",
)
default_staff_agency = st.sidebar.text_input(
    "Default Staff Agency",
    value="HPD",
    help="Applied to staffagency for all rows.",
)
default_event_status = st.sidebar.text_input(
    "Default Event Status",
    value="Event Active",
    help="Applied to eventstatus for all rows.",
)

st.sidebar.markdown("**Time Settings**")
offset_hours = st.sidebar.number_input(
    "Time Offset Hours",
    min_value=-24,
    max_value=24,
    value=5,
    step=1,
    help=(
        "Hours added to all ShiftStart and ShiftEnd times. "
        "Default is +5 to convert from UTC to Central time (CST/CDT). "
        "Change to 0 if your source times are already in local time."
    ),
)

st.sidebar.markdown("---")
st.sidebar.markdown("**Template Status**")
template_exists = os.path.isfile(TEMPLATE_PATH)
if template_exists:
    st.sidebar.success("✅ ArcGIS template found")
else:
    st.sidebar.error(
        "❌ Template NOT found\n\n"
        "Place `Sample_Batch_Load_Event_Staff_Template.xlsx` "
        "in the `templates/` folder and restart the app."
    )

# ---------------------------------------------------------------------------
# Step 1 — File upload
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 1 — Upload Staffing Files")
st.markdown(
    "Drag and drop one or more staffing CSV files. You can upload multiple files at once "
    "(for example, Days, Evenings, and Nights workups) and they will all be combined into "
    "a single output workbook."
)

uploaded_files = st.file_uploader(
    "Upload staffing files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Accepts CurrentStaffingReport.csv or SpecialEventWorkupforGIS.csv (any filename).",
)

if not uploaded_files:
    st.info("👆 Upload one or more files above to begin.")
    st.stop()

# ---------------------------------------------------------------------------
# Helper: process one uploaded file
# ---------------------------------------------------------------------------

def process_file(uploaded_file):
    file_bytes = uploaded_file.read()
    file_name  = uploaded_file.name

    def _make_buf():
        buf = io.BytesIO(file_bytes)
        buf.name = file_name
        return buf

    # Load
    try:
        raw_df, ext = load_source_file(_make_buf())
    except Exception as e:
        return None, None, [f"**{file_name}**: Could not read file — {e}"], "error"

    # Detect format
    fmt = detect_source_format(raw_df)
    source_df = raw_df.copy()
    _header_row_used = None

    if fmt == "unknown":
        header_row = find_current_staffing_header_row(raw_df)
        if header_row is not None:
            source_df = reparse_staffing_from_raw(raw_df, header_row)
            fmt = detect_source_format(source_df)
            _header_row_used = header_row

    if fmt == "unknown":
        return None, None, [
            f"**{file_name}**: Could not identify source format. "
            f"Expected columns for SpecialEventWorkupforGIS or CurrentStaffingReport were not found."
        ], "unknown"

    fmt_labels = {
        "workup":   "SpecialEventWorkupforGIS.csv",
        "staffing": "CurrentStaffingReport.csv",
    }
    fmt_label = fmt_labels.get(fmt, fmt)

    # Strip title rows for staffing format if not already done
    if fmt == "staffing" and _header_row_used is None:
        header_row = find_current_staffing_header_row(raw_df)
        if header_row is not None and header_row > 0:
            source_df = reparse_staffing_from_raw(raw_df, header_row)

    # Column warnings
    col_warnings = []
    if fmt == "workup":
        expected = {"UnitId", "StaffRank", "StaffName", "ShiftStart", "ShiftEnd"}
        actual_lower = {str(c).strip().lower() for c in source_df.columns}
        missing = [c for c in expected if c.lower() not in actual_lower]
        if missing:
            col_warnings.append(f"**{file_name}**: Missing source columns: {missing}")
    elif fmt == "staffing":
        expected = {"EmpID", "RankDescription", "LastName", "FirstName",
                    "RadioCallNumber", "UnitNo", "shiftStart", "ShiftEnd"}
        actual_lower = {str(c).strip().lower() for c in source_df.columns}
        missing = [c for c in expected if c.lower() not in actual_lower]
        if missing:
            col_warnings.append(f"**{file_name}**: Missing source columns: {missing}")

    # Transform
    try:
        if fmt == "workup":
            output_df, transform_warnings = transform_special_event_workup(
                source_df,
                offset_hours=float(offset_hours),
                default_unit_type=default_unit_type_workup,
                default_staff_status=default_staff_status,
                default_staff_agency=default_staff_agency,
                default_event_status=default_event_status,
            )
        else:
            output_df, transform_warnings = transform_current_staffing_report(
                source_df,
                offset_hours=float(offset_hours),
                default_unit_type=default_unit_type_staffing,
                default_staff_status=default_staff_status,
                default_staff_agency=default_staff_agency,
                default_event_status=default_event_status,
            )
    except Exception as e:
        return None, source_df, [
            f"**{file_name}**: Transformation failed — {e}\n{traceback.format_exc()}"
        ], fmt_label

    prefixed_warnings = [f"**{file_name}**: {w}" for w in transform_warnings]
    all_warnings = col_warnings + prefixed_warnings

    return output_df, source_df, all_warnings, fmt_label


# ---------------------------------------------------------------------------
# Step 2 — Process files and show input previews
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 2 — Input Preview")
st.markdown(
    "Each uploaded file is shown below. Expand a file to verify the source data "
    "was read correctly before reviewing the converted output."
)

all_output_dfs = []
all_warnings   = []
file_summaries = []

for uf in uploaded_files:
    output_df, source_df, warnings, fmt_label = process_file(uf)
    all_warnings.extend(warnings)

    row_count = len(output_df) if output_df is not None and not output_df.empty else 0
    file_summaries.append((uf.name, fmt_label, row_count))

    if source_df is not None:
        with st.expander(
            f"📄 {uf.name}  —  {fmt_label}  —  {row_count} rows converted",
            expanded=False,
        ):
            st.markdown(
                "This is the raw source data read from the file. "
                "The app automatically strips report title rows and finds the real column header."
            )
            st.dataframe(source_df.head(20), use_container_width=True)
            if len(source_df) > 20:
                st.caption(f"Showing first 20 of {len(source_df)} rows.")

    if output_df is not None and not output_df.empty:
        all_output_dfs.append(output_df)

# ---------------------------------------------------------------------------
# File summary table
# ---------------------------------------------------------------------------

st.subheader("Files Processed")
st.markdown(
    "Summary of all uploaded files. "
    "If a file shows **error** or **0 rows**, check the Warnings section below."
)
summary_data = {
    "File":            [s[0] for s in file_summaries],
    "Detected Format": [s[1] for s in file_summaries],
    "Rows Converted":  [s[2] for s in file_summaries],
}
st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Stop early with warnings if no output produced
# ---------------------------------------------------------------------------

if not all_output_dfs:
    st.divider()
    st.subheader("⚠️ Warnings")
    for w in all_warnings:
        st.warning(w)
    st.error(
        "No rows were produced from any uploaded file. "
        "Review the warnings above and check that you uploaded a supported file format."
    )
    st.stop()

combined_df = pd.concat(all_output_dfs, ignore_index=True)

# ---------------------------------------------------------------------------
# Step 3 — Validation summary
# ---------------------------------------------------------------------------

val_warnings = validate_output(combined_df)
all_warnings += val_warnings

st.divider()
st.subheader("Step 3 — Validation Summary")
st.markdown(
    "The app checks the converted data for common issues before writing the output file. "
    "Warnings do **not** stop the download — review them and correct the source data if needed."
)

if not all_warnings:
    st.success("✅ No validation issues found. Data looks clean.")
else:
    st.markdown(f"**{len(all_warnings)} issue(s) found:**")
    for w in all_warnings:
        st.warning(w)

# ---------------------------------------------------------------------------
# Step 4 — Transformed output preview
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 4 — Converted Output Preview")
st.markdown(
    "This is the data that will be written into the ArcGIS batch upload workbook. "
    "Column names match the **Staff List** sheet in the ArcGIS template exactly. "
    "Dates have been offset by the configured time offset hours."
)

preview_df = combined_df.copy()
for col in ("unitshiftstart", "unitshiftend"):
    if col in preview_df.columns:
        preview_df[col] = preview_df[col].apply(
            lambda v: v.strftime("%Y/%m/%d %I:%M:%S %p") if hasattr(v, "strftime") else str(v)
        )

st.dataframe(preview_df, use_container_width=True)
st.caption(
    f"**{len(combined_df)} total rows** from {len(all_output_dfs)} file(s) ready for upload."
)

# ---------------------------------------------------------------------------
# Step 5 — Download
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Step 5 — Download ArcGIS Upload Workbook")
st.markdown(
    "Click the button below to download the completed Excel workbook. "
    "Open the downloaded file and import it into ArcGIS using the "
    "Special Event Solution batch upload process."
)

if not template_exists:
    st.error(
        "Cannot generate download — ArcGIS template not found at "
        f"`{TEMPLATE_PATH}`. Place the file in the `templates/` folder and restart the app."
    )
    st.stop()

try:
    workbook_bytes = write_to_template(combined_df, TEMPLATE_PATH)
except Exception as e:
    st.error(f"Failed to write template: {e}")
    with st.expander("Traceback"):
        st.code(traceback.format_exc())
    st.stop()

st.download_button(
    label="⬇️ Download ArcGIS Upload Workbook",
    data=workbook_bytes,
    file_name=OUTPUT_FILENAME,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption(
    f"Downloaded file: `{OUTPUT_FILENAME}` — "
    "contains all converted rows in the ArcGIS Staff List format."
)
