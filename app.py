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
st.caption(
    "Upload one or more staffing CSVs, preview the conversion, and download "
    "the completed ArcGIS batch upload workbook."
)

# ---------------------------------------------------------------------------
# Sidebar settings
# ---------------------------------------------------------------------------

st.sidebar.header("Settings")

default_unit_type_workup = st.sidebar.text_input(
    "Default Unit Type (SpecialEventWorkup)", value="Vehicle"
)
default_unit_type_staffing = st.sidebar.text_input(
    "Default Unit Type (CurrentStaffingReport)", value="Traffic Control"
)
default_staff_status = st.sidebar.text_input(
    "Default Staff Status", value="On Duty"
)
default_staff_agency = st.sidebar.text_input(
    "Default Staff Agency", value="HPD"
)
default_event_status = st.sidebar.text_input(
    "Default Event Status", value="Event Active"
)
offset_hours = st.sidebar.number_input(
    "Time Offset Hours", min_value=-24, max_value=24, value=5, step=1,
    help="Hours added to all shift start/end times. Default is +5 (UTC to CST/CDT).",
)

st.sidebar.markdown("---")
st.sidebar.markdown(
    "**Template path:**\n\n`templates/Sample_Batch_Load_Event_Staff_Template.xlsx`"
)
template_exists = os.path.isfile(TEMPLATE_PATH)
if template_exists:
    st.sidebar.success("Template found")
else:
    st.sidebar.error("Template NOT found — place file in `templates/` folder")

# ---------------------------------------------------------------------------
# File upload — multiple files allowed
# ---------------------------------------------------------------------------

uploaded_files = st.file_uploader(
    "Upload staffing files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Upload one or more CurrentStaffingReport or SpecialEventWorkupforGIS files.",
)

if not uploaded_files:
    st.info("Drag and drop one or more CSV or XLSX files above to begin.")
    st.stop()

# ---------------------------------------------------------------------------
# Helper: process one uploaded file -> (output_df, all_warnings, fmt_label)
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
        return None, [f"**{file_name}**: Could not read file — {e}"], "error"

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
        return None, [
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
        actual_lower = {c.strip().lower() for c in source_df.columns}
        missing = [c for c in expected if c.lower() not in actual_lower]
        if missing:
            col_warnings.append(f"**{file_name}**: Missing source columns: {missing}")
    elif fmt == "staffing":
        expected = {"EmpID", "RankDescription", "LastName", "FirstName",
                    "RadioCallNumber", "UnitNo", "shiftStart", "ShiftEnd"}
        actual_lower = {c.strip().lower() for c in source_df.columns}
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
        return None, [f"**{file_name}**: Transformation failed — {e}\n{traceback.format_exc()}"], fmt_label

    # Prefix file warnings with filename
    prefixed_warnings = [f"**{file_name}**: {w}" for w in transform_warnings]
    all_warnings = col_warnings + prefixed_warnings

    return output_df, all_warnings, fmt_label


# ---------------------------------------------------------------------------
# Process all uploaded files
# ---------------------------------------------------------------------------

st.divider()

all_output_dfs = []
all_warnings   = []
file_summaries = []

for uf in uploaded_files:
    output_df, warnings, fmt_label = process_file(uf)
    all_warnings.extend(warnings)

    if output_df is not None and not output_df.empty:
        all_output_dfs.append(output_df)
        file_summaries.append((uf.name, fmt_label, len(output_df)))
    else:
        file_summaries.append((uf.name, fmt_label, 0))

# ---------------------------------------------------------------------------
# Per-file summary table
# ---------------------------------------------------------------------------

st.subheader(f"Files Processed: {len(uploaded_files)}")

summary_data = {
    "File": [s[0] for s in file_summaries],
    "Detected Format": [s[1] for s in file_summaries],
    "Rows": [s[2] for s in file_summaries],
}
st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# Combine all output rows
# ---------------------------------------------------------------------------

if not all_output_dfs:
    st.error("No rows produced from any uploaded file.")
    st.stop()

combined_df = pd.concat(all_output_dfs, ignore_index=True)

# ---------------------------------------------------------------------------
# Validation summary
# ---------------------------------------------------------------------------

val_warnings = validate_output(combined_df)
all_warnings += val_warnings

st.divider()
st.subheader("Validation Summary")

if not all_warnings:
    st.success("No validation issues found.")
else:
    for w in all_warnings:
        st.warning(w)

# ---------------------------------------------------------------------------
# Combined output preview
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Transformed Output Preview (All Files Combined)")

preview_df = combined_df.copy()
for col in ("unitshiftstart", "unitshiftend"):
    if col in preview_df.columns:
        preview_df[col] = preview_df[col].apply(
            lambda v: v.strftime("%Y/%m/%d %I:%M:%S %p") if hasattr(v, "strftime") else str(v)
        )

st.dataframe(preview_df, use_container_width=True)
st.caption(f"{len(combined_df)} total rows from {len(all_output_dfs)} file(s) ready for upload.")

# ---------------------------------------------------------------------------
# Write to template and provide download
# ---------------------------------------------------------------------------

st.divider()

if not template_exists:
    st.error(
        "Cannot generate download — template not found at "
        f"`{TEMPLATE_PATH}`. Place the file there and reload."
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
    label="Download ArcGIS Upload Workbook",
    data=workbook_bytes,
    file_name=OUTPUT_FILENAME,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption(f"Output file: `{OUTPUT_FILENAME}`")
