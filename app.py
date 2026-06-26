"""
app.py
------
Streamlit web UI for the ArcGIS Special Event Staff Converter.

This is the application entry point. It handles:
  - File upload (multiple files, drag-and-drop)
  - Per-file detection, transformation, and input preview
  - Validation summary across all uploaded files
  - Combined output preview
  - Download as an ArcGIS-ready Excel workbook

Run locally with:
    streamlit run app.py

All business logic (column mapping, rank normalization, datetime parsing,
template writing) lives in staff_transformer.py -- this file is UI only.
Keeping them separate makes the transformer testable without Streamlit.

Requires Python 3.9+ with packages from requirements.txt, or the ArcGIS
Pro conda environment with those packages installed via pip.
"""

import os           # path operations for locating the template file at startup
import io           # BytesIO for buffering uploaded file bytes for reliable multi-read
import traceback    # format full Python stack traces for display in error expanders
from datetime import datetime, timedelta, time as dt_time  # timestamp, shift picker math

import pandas as pd      # DataFrame for the summary table and multi-file concat
import streamlit as st   # entire UI framework: widgets, layout, state, download

# Import all transformation functions from the business logic module.
# Every function here is tested and reusable without Streamlit.
from staff_transformer import (
    ARCGIS_COLUMNS,
    UNIT_SHIFT_OPTIONS,
    UNIT_TYPE_OPTIONS,
    STAFF_STATUS_OPTIONS,
    STAFF_AGENCY_OPTIONS,
    EVENT_STATUS_OPTIONS,
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


# ===========================================================================
# CONFIGURATION CONSTANTS
# ===========================================================================

# Absolute path to the ArcGIS template workbook.
# __file__ is this script's path; os.path.dirname gives its containing directory.
# Joining with "templates/" ensures the path is correct regardless of which
# directory the user runs "streamlit run app.py" from.
TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), "templates", "Sample_Batch_Load_Event_Staff_Template.xlsx"
)

# Output filename base -- timestamp appended at download time so each download
# is uniquely named and the analyst can sort/identify files by when they were generated.
OUTPUT_FILENAME_BASE = "Batch_Load_Event_Staff_Template"


# ===========================================================================
# PAGE SETUP
# ===========================================================================

# set_page_config() MUST be the first Streamlit call in the script.
# Placing any other st.* call before this raises a StreamlitAPIException.
# layout="wide" uses the full browser window width for the data preview tables.
st.set_page_config(
    page_title="ArcGIS Special Event Staff Converter",
    page_icon="🚔",
    layout="wide",
)

st.title("ArcGIS Special Event Staff Converter")

# Top-of-page overview so new users understand the tool before uploading files.
# Written as a numbered list to match the numbered steps shown on the page.
st.markdown(
    """
    This tool converts HPD staffing export files into the format required by the
    **ArcGIS Special Event Solution** batch upload workbook.

    **Supported input files:**
    - `CurrentStaffingReport.csv` — exported from the HPD staffing system
    - `SpecialEventWorkupforGIS.csv` — exported from the Special Event Workup tool

    **How to use:**
    1. Upload one or more staffing files using the file uploader below.
    2. Review the input preview and compare it to the converted output in Step 3.
    3. Check the validation summary in Step 4 for any warnings.
    4. Click **Download ArcGIS Upload Workbook** in Step 5 to get the completed Excel file.
    5. Import the downloaded file into ArcGIS using the batch upload process.
    """
)


# ===========================================================================
# SIDEBAR -- USER-CONFIGURABLE SETTINGS
# ===========================================================================

st.sidebar.header("⚙️ Settings")
st.sidebar.caption("Click any field to expand and edit. Filled values override all output rows.")

# ---------------------------------------------------------------------------
# Default Field Values -- each in its own collapsible expander
# ---------------------------------------------------------------------------

with st.sidebar.expander("Default Unit Type (SpecialEventWorkup)"):
    default_unit_type_workup = st.selectbox(
        "unit_type_workup", options=UNIT_TYPE_OPTIONS,
        index=UNIT_TYPE_OPTIONS.index("Vehicle"), label_visibility="collapsed",
        help="Overrides unittype for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Unit Type (CurrentStaffingReport)"):
    default_unit_type_staffing = st.selectbox(
        "unit_type_staffing", options=UNIT_TYPE_OPTIONS,
        index=UNIT_TYPE_OPTIONS.index("Vehicle"), label_visibility="collapsed",
        help="Overrides unittype for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Unit Radio"):
    default_unit_radio = st.text_input(
        "unit_radio", value="", label_visibility="collapsed",
        help="Overrides unitradio for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Unit Duties"):
    default_unit_duties = st.text_input(
        "unit_duties", value="", label_visibility="collapsed",
        help="Overrides unitduties for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Unit Shift"):
    default_unitshift = st.selectbox(
        "unit_shift", options=UNIT_SHIFT_OPTIONS, index=0, label_visibility="collapsed",
        help="Overrides unitshift for all rows. Leave blank to use derived/source values.",
    )

with st.sidebar.expander("Default Staff Status"):
    default_staff_status = st.selectbox(
        "staff_status", options=STAFF_STATUS_OPTIONS,
        index=STAFF_STATUS_OPTIONS.index("On Duty"), label_visibility="collapsed",
        help="Overrides staffstatus for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Staff Agency"):
    default_staff_agency = st.selectbox(
        "staff_agency", options=STAFF_AGENCY_OPTIONS,
        index=STAFF_AGENCY_OPTIONS.index("HPD"), label_visibility="collapsed",
        help="Overrides staffagency for all rows. Leave blank to use source file values.",
    )

with st.sidebar.expander("Default Event Status"):
    default_event_status = st.selectbox(
        "event_status", options=EVENT_STATUS_OPTIONS,
        index=EVENT_STATUS_OPTIONS.index("Event Active"), label_visibility="collapsed",
        help="Overrides eventstatus for all rows. Leave blank to use source file values.",
    )

# ---------------------------------------------------------------------------
# Time offset -- collapsible
# ---------------------------------------------------------------------------

with st.sidebar.expander("Time Offset Hours"):
    offset_hours = st.number_input(
        "offset_hours", min_value=-24, max_value=24, value=5, step=1,
        label_visibility="collapsed",
        help=(
            "Hours added to ShiftStart and ShiftEnd times. "
            "+5 converts UTC to CDT. Set to 0 if source times are already local."
        ),
    )

# ---------------------------------------------------------------------------
# Default Shift Times -- date + time pickers, collapsible per field
# ---------------------------------------------------------------------------

default_unitshiftstart = None
with st.sidebar.expander("Default Shift Start"):
    st.caption(f"Selected time + {int(offset_hours)}h offset will be stored.")
    ss_enable = st.checkbox("Enable shift start override", key="ss_enable")
    if ss_enable:
        ss_date = st.date_input("Date", value=datetime.today().date(), key="ss_date")
        ss_time = st.time_input("Time", value=dt_time(6, 0), key="ss_time")
        raw_ss = datetime.combine(ss_date, ss_time)
        default_unitshiftstart = raw_ss + timedelta(hours=float(offset_hours))
        st.caption(f"Stored as: {default_unitshiftstart.strftime('%Y/%m/%d %I:%M %p')}")

default_unitshiftend = None
with st.sidebar.expander("Default Shift End"):
    st.caption(f"Selected time + {int(offset_hours)}h offset will be stored.")
    se_enable = st.checkbox("Enable shift end override", key="se_enable")
    if se_enable:
        se_date = st.date_input("Date", value=datetime.today().date(), key="se_date")
        se_time = st.time_input("Time", value=dt_time(14, 0), key="se_time")
        raw_se = datetime.combine(se_date, se_time)
        default_unitshiftend = raw_se + timedelta(hours=float(offset_hours))
        st.caption(f"Stored as: {default_unitshiftend.strftime('%Y/%m/%d %I:%M %p')}")

# ---------------------------------------------------------------------------
# Template status indicator
# ---------------------------------------------------------------------------
# The ArcGIS template workbook MUST be present before the app can generate
# output. We check at startup and display the status prominently so the
# analyst sees the issue immediately -- not after uploading and processing files.
# ---------------------------------------------------------------------------
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


# ===========================================================================
# STEP 1 -- FILE UPLOAD
# ===========================================================================

st.divider()
st.subheader("Step 1 — Upload Staffing Files")
st.markdown(
    "Drag and drop one or more staffing CSV files. You can upload multiple files at once "
    "(for example, Days, Evenings, and Nights workups) and they will all be combined into "
    "a single output workbook."
)

# accept_multiple_files=True allows the analyst to combine multiple shifts or
# multiple event days into a single upload batch. All uploaded files are processed
# individually and their rows are concatenated before writing to the template.
uploaded_files = st.file_uploader(
    "Upload staffing files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Accepts CurrentStaffingReport.csv or SpecialEventWorkupforGIS.csv (any filename).",
)

# Stop rendering the rest of the page until at least one file is uploaded.
# st.stop() exits the Streamlit script run immediately -- nothing below executes
# until the user uploads a file and Streamlit reruns the script.
if not uploaded_files:
    st.info("👆 Upload one or more files above to begin.")
    st.stop()


# ===========================================================================
# HELPER: process_file()
# ===========================================================================

def process_file(uploaded_file):
    """
    Process a single uploaded file through the full pipeline: load -> detect -> transform.

    Returns a 4-tuple: (output_df, source_df, warnings, fmt_label)
        output_df:  20-column ArcGIS DataFrame ready for the template; None on error
        source_df:  cleaned source DataFrame shown in the input preview; None on load error
        warnings:   list of warning strings to display in the Step 3 validation section
        fmt_label:  human-readable format string (e.g., "CurrentStaffingReport.csv")
                    or "error" / "unknown" if the file could not be processed

    -------------------------------------------------------------------------
    WHY WE BUFFER FILE BYTES BEFORE ANY READS:
    -------------------------------------------------------------------------
    Streamlit's UploadedFile object is a file-like wrapper. After any code
    calls file_obj.read() once, the internal read position is at EOF. Calling
    file_obj.seek(0) is NOT reliably supported across all platforms and
    Streamlit versions -- position may not reset, causing the second read to
    return empty bytes and crashing the transform silently.

    Solution: read ALL bytes once into the `file_bytes` variable, then create
    a fresh io.BytesIO object (with .name set) for EVERY function that needs
    to read the file. Each BytesIO starts at position 0 automatically.

    The _make_buf() inner function encapsulates this pattern cleanly.
    -------------------------------------------------------------------------
    """
    # Read all file bytes into memory exactly once
    file_bytes = uploaded_file.read()
    file_name  = uploaded_file.name

    def _make_buf():
        """Return a fresh BytesIO buffer at position 0, with .name set for extension detection."""
        buf = io.BytesIO(file_bytes)
        buf.name = file_name  # load_source_file() reads .name to determine file type
        return buf

    # ------------------------------------------------------------------
    # Phase 1: Load raw file into DataFrame
    # ------------------------------------------------------------------
    # load_source_file() uses Python's csv module (not pandas) for CSV files
    # to handle variable-width rows (title rows + data rows). See staff_transformer.py
    # for the detailed explanation of why pandas cannot handle this file structure.
    try:
        raw_df, ext = load_source_file(_make_buf())
    except Exception as e:
        # Return gracefully with error label -- a bad file must not crash the whole app
        return None, None, [f"**{file_name}**: Could not read file — {e}"], "error"

    # ------------------------------------------------------------------
    # Phase 2: Detect source format from column names
    # ------------------------------------------------------------------
    # For SpecialEventWorkupforGIS.csv this succeeds immediately because the
    # first row IS the header. For CurrentStaffingReport.csv the first rows
    # are title rows so column names are integers (0, 1, 2, ...) and format
    # detection returns "unknown".
    fmt       = detect_source_format(raw_df)
    source_df = raw_df.copy()  # shown in input preview; may be updated below
    _header_row_used = None     # track whether header detection already ran

    # ------------------------------------------------------------------
    # Phase 3: Fallback header detection for title-row CSVs
    # ------------------------------------------------------------------
    # When format is "unknown", scan row VALUES for known staffing column
    # headers (e.g., "Division", "RankDescription", "EmpID"). Once found,
    # reparse_staffing_from_raw() slices the raw DataFrame at that row to
    # produce a properly-headered DataFrame without re-reading the file.
    if fmt == "unknown":
        header_row = find_current_staffing_header_row(raw_df)
        if header_row is not None:
            source_df = reparse_staffing_from_raw(raw_df, header_row)
            fmt = detect_source_format(source_df)  # re-detect on clean DataFrame
            _header_row_used = header_row

    # If format is still unknown after fallback, the file structure is not recognized
    if fmt == "unknown":
        return None, None, [
            f"**{file_name}**: Could not identify source format. "
            f"Expected columns for SpecialEventWorkupforGIS or CurrentStaffingReport were not found."
        ], "unknown"

    # Map internal format keys to display names shown in the UI summary table
    fmt_labels = {
        "workup":   "SpecialEventWorkupforGIS.csv",
        "staffing": "CurrentStaffingReport.csv",
    }
    fmt_label = fmt_labels.get(fmt, fmt)

    # ------------------------------------------------------------------
    # Phase 4: Strip title rows for staffing format (if not done yet)
    # ------------------------------------------------------------------
    # If format was detected directly on the first pass (without the fallback
    # above), raw_df may still include title rows at the top of the DataFrame.
    # Reparse now so transform_current_staffing_report() sees only data rows.
    if fmt == "staffing" and _header_row_used is None:
        header_row = find_current_staffing_header_row(raw_df)
        if header_row is not None and header_row > 0:
            source_df = reparse_staffing_from_raw(raw_df, header_row)

    # ------------------------------------------------------------------
    # Phase 5: Check for missing source columns
    # ------------------------------------------------------------------
    # Warn if expected columns are absent -- the transform will produce blank
    # values for those fields, which may cause ArcGIS validation errors downstream.
    col_warnings = []
    if fmt == "workup":
        expected     = {"UnitId", "StaffRank", "StaffName", "ShiftStart", "ShiftEnd"}
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

    # ------------------------------------------------------------------
    # Phase 6: Transform source DataFrame to ArcGIS output schema
    # ------------------------------------------------------------------
    # Each transform function returns (output_df, warnings_list).
    # We pass all sidebar settings so the transform uses the analyst's
    # configured defaults rather than hardcoded values.
    try:
        if fmt == "workup":
            output_df, transform_warnings = transform_special_event_workup(
                source_df,
                offset_hours=float(offset_hours),
                default_unit_type=default_unit_type_workup,
                default_unit_radio=default_unit_radio,
                default_unit_duties=default_unit_duties,
                default_unitshift=default_unitshift,
                default_staff_status=default_staff_status,
                default_staff_agency=default_staff_agency,
                default_event_status=default_event_status,
                default_unitshiftstart=default_unitshiftstart,
                default_unitshiftend=default_unitshiftend,
            )
        else:  # "staffing"
            output_df, transform_warnings = transform_current_staffing_report(
                source_df,
                offset_hours=float(offset_hours),
                default_unit_type=default_unit_type_staffing,
                default_unit_radio=default_unit_radio,
                default_unit_duties=default_unit_duties,
                default_unitshift=default_unitshift,
                default_staff_status=default_staff_status,
                default_staff_agency=default_staff_agency,
                default_event_status=default_event_status,
                default_unitshiftstart=default_unitshiftstart,
                default_unitshiftend=default_unitshiftend,
            )
    except Exception as e:
        # Return source_df so the input preview is still shown even if transform fails.
        # Full traceback is included in the warning for debugging.
        return None, source_df, [
            f"**{file_name}**: Transformation failed — {e}\n{traceback.format_exc()}"
        ], fmt_label

    # Prefix all transform warnings with the filename so the analyst can tell
    # which file each warning came from when multiple files are uploaded at once.
    prefixed_warnings = [f"**{file_name}**: {w}" for w in transform_warnings]
    all_warnings = col_warnings + prefixed_warnings

    return output_df, source_df, all_warnings, fmt_label


# ===========================================================================
# STEP 2 -- PROCESS ALL UPLOADED FILES AND SHOW INPUT PREVIEWS
# ===========================================================================

st.divider()
st.subheader("Step 2 — Input Preview")
st.markdown(
    "Each uploaded file is shown below. Expand a file to verify the source data "
    "was read correctly before reviewing the converted output."
)

# These lists accumulate results across all uploaded files for later steps
all_output_dfs = []   # list of output DataFrames -- pd.concat'd after the loop
all_warnings   = []   # all warning strings from all files -- shown in Step 3
file_summaries = []   # (filename, format_label, row_count) -- for the summary table

for uf in uploaded_files:
    output_df, source_df, warnings, fmt_label = process_file(uf)
    all_warnings.extend(warnings)

    # Count converted rows; 0 when output_df is None (error) or empty
    row_count = len(output_df) if output_df is not None and not output_df.empty else 0
    file_summaries.append((uf.name, fmt_label, row_count))

    # Show per-file input preview in a collapsible expander.
    # expanded=False keeps the UI compact when many files are uploaded.
    if source_df is not None:
        with st.expander(
            f"📄 {uf.name}  —  {fmt_label}  —  {row_count} rows converted",
            expanded=False,
        ):
            st.markdown(
                "This is the raw source data read from the file. "
                "The app automatically strips report title rows and finds the real column header."
            )
            # Show first 20 rows to keep the UI responsive for large staffing files
            st.dataframe(source_df.head(20), use_container_width=True)
            if len(source_df) > 20:
                st.caption(f"Showing first 20 of {len(source_df)} rows.")

    # Collect non-empty output DataFrames to combine into the final output
    if output_df is not None and not output_df.empty:
        all_output_dfs.append(output_df)


# ===========================================================================
# FILES PROCESSED -- SUMMARY TABLE
# ===========================================================================

st.subheader("Files Processed")
st.markdown(
    "Summary of all uploaded files. "
    "If a file shows **error** or **0 rows**, check the Warnings section below."
)

# Build a simple summary DataFrame displayed as a clean table with no index column
summary_data = {
    "File":            [s[0] for s in file_summaries],
    "Detected Format": [s[1] for s in file_summaries],
    "Rows Converted":  [s[2] for s in file_summaries],
}
st.dataframe(pd.DataFrame(summary_data), use_container_width=True, hide_index=True)


# ===========================================================================
# EARLY STOP -- NO OUTPUT PRODUCED
# ===========================================================================

# If every file failed to produce output rows, show warnings and stop.
# There is nothing to validate, preview, or download without output rows.
if not all_output_dfs:
    st.divider()
    st.subheader("⚠️ Warnings")
    for w in all_warnings:
        st.warning(w)
    st.error(
        "No rows were produced from any uploaded file. "
        "Review the warnings above and check that you uploaded a supported file format."
    )
    st.stop()  # exits the script; nothing below executes

# Combine all per-file output DataFrames into a single DataFrame.
# ignore_index=True resets the row index so there are no duplicate index values
# when multiple files each produced rows starting at index 0.
combined_df = pd.concat(all_output_dfs, ignore_index=True)


# ===========================================================================
# STEP 3 -- CONVERTED OUTPUT PREVIEW
# ===========================================================================
# Placed immediately after Step 2 (input preview) so the analyst can scroll
# between the source table and the converted output table to compare them.

st.divider()
st.subheader("Step 3 — Converted Output Preview")
st.markdown(
    "This is the data that will be written into the ArcGIS batch upload workbook. "
    "Column names match the **Staff List** sheet in the ArcGIS template exactly. "
    "Dates have been offset by the configured time offset hours."
)

# Format datetime columns as readable strings for display ONLY.
# combined_df holds Python datetime objects in unitshiftstart and unitshiftend --
# Streamlit's st.dataframe() renders those as raw epoch integers without formatting.
# We create a display copy (preview_df) so the actual combined_df retains datetime
# objects for correct Excel serial number writing in write_to_template().
preview_df = combined_df.copy()
for col in ("unitshiftstart", "unitshiftend"):
    if col in preview_df.columns:
        preview_df[col] = preview_df[col].apply(
            # strftime() for datetime objects; str() fallback for raw strings (parse failures)
            lambda v: v.strftime("%Y/%m/%d %I:%M:%S %p") if hasattr(v, "strftime") else str(v)
        )

# data_editor: dropdown columns are editable; all others are read-only.
_editable_cols = {"unitshift", "unittype", "staffstatus", "staffagency", "eventstatus"}
_readonly_cols  = [c for c in preview_df.columns if c not in _editable_cols]
edited_df = st.data_editor(
    preview_df,
    column_config={
        "unitshift": st.column_config.SelectboxColumn(
            "unitshift", options=UNIT_SHIFT_OPTIONS, required=False,
        ),
        "unittype": st.column_config.SelectboxColumn(
            "unittype", options=UNIT_TYPE_OPTIONS, required=False,
        ),
        "staffstatus": st.column_config.SelectboxColumn(
            "staffstatus", options=STAFF_STATUS_OPTIONS, required=False,
        ),
        "staffagency": st.column_config.SelectboxColumn(
            "staffagency", options=STAFF_AGENCY_OPTIONS, required=False,
        ),
        "eventstatus": st.column_config.SelectboxColumn(
            "eventstatus", options=EVENT_STATUS_OPTIONS, required=False,
        ),
    },
    disabled=_readonly_cols,
    use_container_width=True,
    hide_index=True,
)
# Apply per-row edits back into combined_df before template write.
for _col in _editable_cols:
    combined_df[_col] = edited_df[_col].values
st.caption(
    f"**{len(combined_df)} total rows** from {len(all_output_dfs)} file(s) ready for upload. "
    "Click any dropdown column cell to change it inline."
)


# ===========================================================================
# STEP 4 -- VALIDATION SUMMARY
# ===========================================================================

# Run validation checks on the combined output and merge new issues into all_warnings.
# validate_output() checks for blank unitid, DUPLICATE unitid, blank staffname,
# blank shift dates, and unknown staffrank values. See staff_transformer.py for details.
val_warnings = validate_output(combined_df)
all_warnings += val_warnings

st.divider()
st.subheader("Step 4 — Validation Summary")
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


# ===========================================================================
# STEP 5 -- DOWNLOAD ARCGIS UPLOAD WORKBOOK
# ===========================================================================

st.divider()
st.subheader("Step 5 — Download ArcGIS Upload Workbook")
st.markdown(
    "Click the button below to download the completed Excel workbook. "
    "Open the downloaded file and import it into ArcGIS using the "
    "Special Event Solution batch upload process."
)

# Guard: re-check template existence here in case the file was removed after startup.
# The sidebar already shows the status, but we block the download explicitly
# to prevent a confusing error from write_to_template() reaching the user.
if not template_exists:
    st.error(
        "Cannot generate download — ArcGIS template not found at "
        f"`{TEMPLATE_PATH}`. Place the file in the `templates/` folder and restart the app."
    )
    st.stop()

# Write the combined output DataFrame into the ArcGIS template workbook.
# write_to_template() returns an io.BytesIO object at position 0.
# We wrap in try/except so that template write errors show a clear message
# with a collapsible traceback rather than crashing the app with a raw exception.
try:
    workbook_bytes = write_to_template(combined_df, TEMPLATE_PATH)
except Exception as e:
    st.error(f"Failed to write template: {e}")
    with st.expander("Traceback"):
        st.code(traceback.format_exc())
    st.stop()

# Generate a timestamp at the moment the download button is rendered so each
# downloaded file has a unique, sortable name.
# Format: YYYYMMDD_HHMMSS  (e.g., 20240615_143022)
_ts = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"{OUTPUT_FILENAME_BASE}_{_ts}.xlsx"

# Download button: sends workbook_bytes as a file download when clicked.
# type="primary" renders as a blue button to draw the analyst's eye to this action.
# The MIME type tells the browser to treat this as an Excel file (.xlsx).
st.download_button(
    label="⬇️ Download Batch Load Event Staff Template",
    data=workbook_bytes,
    file_name=output_filename,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption(
    f"Downloaded file: `{output_filename}` — "
    "contains all converted rows in the ArcGIS Staff List format."
)
