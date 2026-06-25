"""
app.py
Streamlit UI for the ArcGIS Special Event Staff Converter.
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
    reload_staffing_with_real_header,
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
    "Upload a staffing CSV, preview the conversion, and download the "
    "completed ArcGIS batch upload workbook."
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
# File upload
# ---------------------------------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload staffing file",
    type=["csv", "xlsx", "xls"],
    help="Accepts CurrentStaffingReport.csv or SpecialEventWorkupforGIS.csv",
)

if not uploaded_file:
    st.info("Drag and drop a CSV or XLSX file above to begin.")
    st.stop()

# ---------------------------------------------------------------------------
# Load file — buffer bytes so we can re-read without seek() issues
# ---------------------------------------------------------------------------

st.divider()

file_bytes = uploaded_file.read()
file_name  = uploaded_file.name

def _make_buf():
    buf = io.BytesIO(file_bytes)
    buf.name = file_name  # load_source_file reads .name for extension detection
    return buf

try:
    raw_df, ext = load_source_file(_make_buf())
except ValueError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

st.subheader("Raw Upload Preview")
st.dataframe(raw_df.head(20), use_container_width=True)
st.caption(f"Loaded {len(raw_df)} rows, {len(raw_df.columns)} columns.")

# ---------------------------------------------------------------------------
# Detect format — scan column names first, then row values if needed
# ---------------------------------------------------------------------------

fmt = detect_source_format(raw_df)
source_df = raw_df.copy()
_header_row_used = None

if fmt == "unknown":
    # Columns are title-row garbage (Textbox23, Unnamed:N, etc.).
    # Scan row VALUES for the real staffing header row.
    header_row = find_current_staffing_header_row(raw_df)
    if header_row is not None:
        buf = _make_buf()
        buf.name = file_name
        source_df = reload_staffing_with_real_header(buf, header_row)
        fmt = detect_source_format(source_df)
        _header_row_used = header_row
        if fmt != "unknown":
            st.warning(
                f"Report title rows stripped. Real header found at row {header_row + 1}."
            )
            st.subheader("Re-parsed Data Preview")
            st.dataframe(source_df.head(20), use_container_width=True)
            st.caption(f"Re-parsed: {len(source_df)} rows, {len(source_df.columns)} columns.")

fmt_labels = {
    "workup":   "SpecialEventWorkupforGIS.csv",
    "staffing": "CurrentStaffingReport.csv",
    "unknown":  "Unknown format",
}
fmt_label = fmt_labels.get(fmt, "Unknown format")

if fmt == "unknown":
    st.error(
        "Could not identify the source format. "
        "Expected columns for SpecialEventWorkupforGIS.csv or "
        "CurrentStaffingReport.csv were not found."
    )
    st.stop()

st.info(f"Detected source format: **{fmt_label}**")

# ---------------------------------------------------------------------------
# CurrentStaffingReport: strip title rows if not already done above
# ---------------------------------------------------------------------------

if fmt == "staffing" and _header_row_used is None:
    header_row = find_current_staffing_header_row(raw_df)
    if header_row is not None and header_row > 0:
        st.warning(
            f"Report title rows detected. Real header at row {header_row + 1}."
        )
        buf = _make_buf()
        buf.name = file_name
        source_df = reload_staffing_with_real_header(buf, header_row)
        st.subheader("Re-parsed Data Preview")
        st.dataframe(source_df.head(20), use_container_width=True)
        st.caption(f"Re-parsed: {len(source_df)} rows, {len(source_df.columns)} columns.")

# ---------------------------------------------------------------------------
# Check for missing expected columns
# ---------------------------------------------------------------------------

col_warnings = []

if fmt == "workup":
    expected = {
        "UnitId", "StaffRank", "StaffName", "ShiftStart", "ShiftEnd"
    }
    actual_lower = {c.strip().lower() for c in source_df.columns}
    missing = [c for c in expected if c.lower() not in actual_lower]
    if missing:
        col_warnings.append(f"Missing expected source columns: {missing}")

elif fmt == "staffing":
    expected = {
        "EmpID", "RankDescription", "LastName", "FirstName",
        "RadioCallNumber", "UnitNo", "shiftStart", "ShiftEnd",
    }
    actual_lower = {c.strip().lower() for c in source_df.columns}
    missing = [c for c in expected if c.lower() not in actual_lower]
    if missing:
        col_warnings.append(f"Missing expected source columns: {missing}")

# ---------------------------------------------------------------------------
# Transform
# ---------------------------------------------------------------------------

transform_warnings = []
output_df = pd.DataFrame()

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
    elif fmt == "staffing":
        output_df, transform_warnings = transform_current_staffing_report(
            source_df,
            offset_hours=float(offset_hours),
            default_unit_type=default_unit_type_staffing,
            default_staff_status=default_staff_status,
            default_staff_agency=default_staff_agency,
            default_event_status=default_event_status,
        )
except Exception as e:
    st.error(f"Transformation failed: {e}")
    with st.expander("Traceback"):
        st.code(traceback.format_exc())
    st.stop()

# ---------------------------------------------------------------------------
# Validation summary
# ---------------------------------------------------------------------------

val_warnings = validate_output(output_df)
all_warnings = col_warnings + transform_warnings + val_warnings

st.divider()
st.subheader("Validation Summary")

if not all_warnings:
    st.success("No validation issues found.")
else:
    for w in all_warnings:
        st.warning(w)

# ---------------------------------------------------------------------------
# Output preview
# ---------------------------------------------------------------------------

st.divider()
st.subheader("Transformed Output Preview")

if output_df.empty:
    st.error("Transformed output is empty — nothing to write.")
    st.stop()

# Display datetime columns as strings for preview readability
preview_df = output_df.copy()
for col in ("unitshiftstart", "unitshiftend"):
    if col in preview_df.columns:
        preview_df[col] = preview_df[col].apply(
            lambda v: v.strftime("%Y/%m/%d %I:%M:%S %p") if hasattr(v, "strftime") else str(v)
        )

st.dataframe(preview_df, use_container_width=True)
st.caption(f"{len(output_df)} rows ready for upload.")

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
    workbook_bytes = write_to_template(output_df, TEMPLATE_PATH)
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
