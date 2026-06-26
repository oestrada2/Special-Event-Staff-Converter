"""
staff_transformer.py
All transformation, cleaning, validation, date handling, shift derivation,
and template writing logic for the ArcGIS Special Event Staff Converter.
"""

import re
import io
import copy
import warnings
from datetime import datetime, timedelta
from typing import Optional, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from dateutil import parser as dateutil_parser


# ---------------------------------------------------------------------------
# Column order required by the ArcGIS template
# ---------------------------------------------------------------------------
ARCGIS_COLUMNS = [
    "unitid",
    "unitshift",
    "unitloc",
    "unittype",
    "unitsquad",
    "unitradio",
    "unitduties",
    "payrollnum",
    "staffrank",
    "staffname",
    "staffphone",
    "staffemail",
    "staffskills",
    "staffpay",
    "staffstatus",
    "staffduty",
    "staffagency",
    "unitshiftstart",
    "unitshiftend",
    "eventstatus",
]

# ---------------------------------------------------------------------------
# Mapping: SpecialEventWorkupforGIS.csv -> ArcGIS template
# Keys are ArcGIS column names, values are source column names (or None).
# ---------------------------------------------------------------------------
WORKUP_COLUMN_MAP = {
    "unitid":         "UnitId",
    "unitshift":      "UnitShift",     # fallback: derived from ShiftStart/ShiftEnd
    "unitloc":        None,            # always blank
    "unittype":       "UnitType",      # default: Vehicle
    "unitsquad":      "UnitSquad",
    "unitradio":      "UnitRadio",
    "unitduties":     "UnitDuties",
    "payrollnum":     "Payroll",
    "staffrank":      "StaffRank",
    "staffname":      "StaffName",
    "staffphone":     "StaffPhone",
    "staffemail":     "Staffemail",
    "staffskills":    "StaffSkills",
    "staffpay":       "StaffPay",
    "staffstatus":    "StaffStatus",   # default: On Duty
    "staffduty":      "StaffDuty",
    "staffagency":    "StaffAgency",   # default: HPD
    "unitshiftstart": "ShiftStart",    # +5h offset applied
    "unitshiftend":   "ShiftEnd",      # +5h offset applied
    "eventstatus":    None,            # always "Event Active"
}

# Columns in SpecialEventWorkupforGIS.csv that map to datetime fields
WORKUP_DATETIME_FIELDS = {
    "unitshiftstart": "ShiftStart",
    "unitshiftend":   "ShiftEnd",
}

# ---------------------------------------------------------------------------
# Mapping: CurrentStaffingReport.csv -> ArcGIS template
# ---------------------------------------------------------------------------
STAFFING_COLUMN_MAP = {
    "unitid":         "EmpID",
    "unitshift":      None,            # derived from shift/shiftStart/ShiftEnd
    "unitloc":        None,            # always blank
    "unittype":       None,            # default: Traffic Control
    "unitsquad":      "UnitNo",
    "unitradio":      "RadioCallNumber",
    "unitduties":     None,            # always blank
    "payrollnum":     "EmpID",
    "staffrank":      "RankDescription",
    "staffname":      None,            # derived: LastName, FirstName
    "staffphone":     "CellPhone",
    "staffemail":     None,            # always blank
    "staffskills":    None,            # always blank
    "staffpay":       None,            # always blank
    "staffstatus":    None,            # default: On Duty
    "staffduty":      "Division",
    "staffagency":    None,            # default: HPD
    "unitshiftstart": "shiftStart",    # +5h offset applied
    "unitshiftend":   "ShiftEnd",      # +5h offset applied
    "eventstatus":    None,            # always "Event Active"
}

STAFFING_DATETIME_FIELDS = {
    "unitshiftstart": "shiftStart",
    "unitshiftend":   "ShiftEnd",
}

# Columns that identify the real header row in CurrentStaffingReport.csv
STAFFING_HEADER_IDENTIFIERS = {
    "Division", "RankDescription", "LastName", "FirstName",
    "EmpID", "RadioCallNumber", "UnitNo", "shiftStart", "ShiftEnd",
}

# ID-style columns to preserve as text (no float conversion)
TEXT_ID_COLUMNS = {
    "empid", "payroll", "unitid", "unitno", "radiocallnumber", "unitradio",
}


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def clean_text(value) -> str:
    """Strip, collapse whitespace, return empty string for null/nan."""
    if pd.isna(value) or value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r" {2,}", " ", s)
    return s


def _series_clean(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.replace(r" {2,}", " ", regex=True)


def _case_insensitive_col(df: pd.DataFrame, name: str) -> Optional[str]:
    """Return the actual column name in df matching `name` (case-insensitive)."""
    for col in df.columns:
        if str(col).strip().lower() == name.lower():
            return col
    return None


# ---------------------------------------------------------------------------
# Staff rank normalization
# ---------------------------------------------------------------------------

# Approved ArcGIS staffrank coded values — exact spelling, capitalization,
# punctuation, and spacing required by the ArcGIS Special Event Solution.
APPROVED_RANKS = [
    "POLICE CHIEF",
    "EXECUTIVE CHIEF",
    "EXECUTIVE ASSISTANT POLICE CHIEF",
    "ASSISTANT POLICE CHIEF",
    "POLICE CAPTAIN",
    "POLICE LIEUTENANT",
    "POLICE SERGEANT",
    "Senior Police Officer",
    "POLICE OFFICER",
    "POLICE OFFICER,PROBATIONARY",
    "Civilian",
    "OLEA",
]

# Lookup: lowercase canonical form -> approved rank string
_RANK_CANONICAL = {r.lower(): r for r in APPROVED_RANKS}

# Common alternate source values -> approved rank (lowercase keys)
_RANK_ALIASES: dict = {
    # Police Chief
    "chief of police":                      "POLICE CHIEF",
    "police chief":                         "POLICE CHIEF",
    "chief":                                "POLICE CHIEF",

    # Executive Chief
    "executive chief":                      "EXECUTIVE CHIEF",

    # Executive Assistant Police Chief
    "executive ac":                         "EXECUTIVE ASSISTANT POLICE CHIEF",
    "executive assistant chief":            "EXECUTIVE ASSISTANT POLICE CHIEF",
    "executive assistant police chief":     "EXECUTIVE ASSISTANT POLICE CHIEF",
    "exec ac":                              "EXECUTIVE ASSISTANT POLICE CHIEF",

    # Assistant Police Chief
    "assistant chief":                      "ASSISTANT POLICE CHIEF",
    "assistant police chief":               "ASSISTANT POLICE CHIEF",
    "asst chief":                           "ASSISTANT POLICE CHIEF",

    # Police Captain
    "captain":                              "POLICE CAPTAIN",
    "police captain":                       "POLICE CAPTAIN",
    "capt":                                 "POLICE CAPTAIN",

    # Police Lieutenant
    "lieutenant":                           "POLICE LIEUTENANT",
    "police lieutenant":                    "POLICE LIEUTENANT",
    "lt":                                   "POLICE LIEUTENANT",

    # Police Sergeant
    "sergeant":                             "POLICE SERGEANT",
    "police sergeant":                      "POLICE SERGEANT",
    "sgt":                                  "POLICE SERGEANT",

    # Senior Police Officer
    "senior police officer":                "Senior Police Officer",
    "sr police officer":                    "Senior Police Officer",
    "senior officer":                       "Senior Police Officer",

    # Police Officer
    "police officer":                       "POLICE OFFICER",
    "officer":                              "POLICE OFFICER",
    "po":                                   "POLICE OFFICER",
    "police ofc":                           "POLICE OFFICER",

    # Police Officer Probationary
    "police officer probationary":          "POLICE OFFICER,PROBATIONARY",
    "police officer, probationary":         "POLICE OFFICER,PROBATIONARY",
    "probationary police officer":          "POLICE OFFICER,PROBATIONARY",
    "po probationary":                      "POLICE OFFICER,PROBATIONARY",

    # Civilian
    "civilian":                             "Civilian",
    "hpd civilian":                         "Civilian",

    # OLEA
    "outside leo":                          "OLEA",
    "olea":                                 "OLEA",
}


def normalize_staff_rank(value: str) -> Tuple[str, bool]:
    """
    Normalize a source rank value to the approved ArcGIS staffrank coded value.

    Returns (normalized_rank, matched) where matched=False means the value
    was not found in the approved list or aliases — caller should warn.
    Blank input returns ("", True) — no warning needed for blank.
    """
    cleaned = clean_text(value)
    if not cleaned:
        return "", True

    key = cleaned.lower()

    # Alias match first (covers both alternates and canonical forms)
    if key in _RANK_ALIASES:
        return _RANK_ALIASES[key], True

    # Direct canonical match (case-insensitive fallback)
    if key in _RANK_CANONICAL:
        return _RANK_CANONICAL[key], True

    # No match — return cleaned original so row is not dropped
    return cleaned, False


def _get(df: pd.DataFrame, row: pd.Series, source_col: str) -> str:
    """Case-insensitive get from a row, returning cleaned string."""
    actual = _case_insensitive_col(df, source_col)
    if actual is None:
        return ""
    return clean_text(row.get(actual, ""))


# ---------------------------------------------------------------------------
# File loading
# ---------------------------------------------------------------------------

def load_source_file(file_obj) -> Tuple[pd.DataFrame, str]:
    """
    Load CSV or XLSX from a file-like object.
    Returns (dataframe, file_extension).
    All ID-style columns are read as strings to preserve leading zeros.

    For CSVs, uses Python's csv module to handle files where title rows
    have fewer fields than data rows (e.g. Textbox23 header with 1 field,
    data rows with 13 fields). Pads short rows with empty strings.
    """
    import csv as _csv
    import io as _io

    name = getattr(file_obj, "name", "")
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""

    if ext == "csv":
        raw = file_obj.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="replace")

        reader = _csv.reader(_io.StringIO(raw))
        rows = list(reader)

        if not rows:
            raise ValueError("CSV file is empty.")

        max_cols = max(len(r) for r in rows)
        padded = [r + [""] * (max_cols - len(r)) for r in rows]
        df = pd.DataFrame(padded, dtype=str)

    elif ext in ("xlsx", "xls"):
        df = pd.read_excel(file_obj, dtype=str, keep_default_na=False)
    else:
        raise ValueError(f"Unsupported file type: .{ext}. Expecting .csv or .xlsx")

    return df, ext


# ---------------------------------------------------------------------------
# Format detection
# ---------------------------------------------------------------------------

def detect_source_format(df: pd.DataFrame) -> str:
    """
    Return 'workup', 'staffing', or 'unknown'.
    Checks column names case-insensitively.
    """
    cols_lower = {str(c).strip().lower() for c in df.columns}

    workup_signals = {"unitid", "staffrank", "staffname", "shiftstart", "shiftend"}
    staffing_signals = {"rankdescription", "lastname", "firstname", "empid", "shiftstart", "shiftend"}

    workup_score = len(workup_signals & cols_lower)
    staffing_score = len(staffing_signals & cols_lower)

    if workup_score > staffing_score and workup_score >= 3:
        return "workup"
    if staffing_score > workup_score and staffing_score >= 3:
        return "staffing"
    return "unknown"


# ---------------------------------------------------------------------------
# CurrentStaffingReport: find real header row
# ---------------------------------------------------------------------------

def find_current_staffing_header_row(raw_df: pd.DataFrame) -> Optional[int]:
    """
    Scan rows for the real header (containing STAFFING_HEADER_IDENTIFIERS).
    Returns the 0-based row index, or None if not found.
    """
    for idx, row in raw_df.iterrows():
        cell_values = {str(v).strip() for v in row.values}
        if len(STAFFING_HEADER_IDENTIFIERS & cell_values) >= 5:
            return idx
    return None


def reload_staffing_with_real_header(file_obj, header_row_index: int) -> pd.DataFrame:
    """Re-read the CSV using the detected header row (file-row index, 0-based)."""
    file_obj.seek(0)
    df = pd.read_csv(
        file_obj,
        dtype=str,
        keep_default_na=False,
        header=header_row_index,
    )
    return df


def reparse_staffing_from_raw(raw_df: pd.DataFrame, header_row_idx: int) -> pd.DataFrame:
    """
    Reinterpret an already-loaded raw_df using row `header_row_idx` as the header.
    Avoids file re-reads and off-by-one issues with pandas header= parameter.
    """
    new_cols = [str(v).strip() for v in raw_df.iloc[header_row_idx].tolist()]
    data = raw_df.iloc[header_row_idx + 1:].copy()
    data.columns = new_cols
    data = data.reset_index(drop=True)
    return data


# ---------------------------------------------------------------------------
# DateTime parsing and offset
# ---------------------------------------------------------------------------

_EXCEL_DATE_FORMAT = "yyyy/m/dd h:mm:ss AM/PM"

# openpyxl number format that matches the required Excel display
_OPENPYXL_DATE_FORMAT = "yyyy/m/dd h:mm:ss AM/PM"


def parse_and_offset_datetime(
    value: str,
    offset_hours: float = 5,
) -> Tuple[Optional[datetime], bool]:
    """
    Parse a date/time string and add `offset_hours`.
    Returns (datetime_object, success_bool).
    On failure returns (None, False).
    """
    if not value or value.strip() == "":
        return None, False
    try:
        dt = dateutil_parser.parse(value, dayfirst=False)
        dt_offset = dt + timedelta(hours=offset_hours)
        return dt_offset, True
    except Exception:
        return None, False


# ---------------------------------------------------------------------------
# Unit shift derivation
# ---------------------------------------------------------------------------

def derive_unit_shift(start_str: str, end_str: str) -> str:
    """
    Derive a readable shift label from start/end time strings.
    Handles shifts crossing midnight.
    """
    try:
        start_dt = dateutil_parser.parse(start_str, dayfirst=False)
        end_dt = dateutil_parser.parse(end_str, dayfirst=False)
    except Exception:
        return ""

    # Handle midnight crossing
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    duration_hours = (end_dt - start_dt).total_seconds() / 3600

    # Round to nearest common shift length
    if duration_hours <= 9:
        hrs_label = "8 hr shift"
    elif duration_hours <= 11:
        hrs_label = "10 hr shift"
    else:
        hrs_label = "12 hr shift"

    # Daypart by start hour
    start_hour = start_dt.hour
    if 5 <= start_hour <= 14:
        daypart = "Days"
    elif 15 <= start_hour <= 21:
        daypart = "Evenings"
    else:
        daypart = "Nights"

    return f"{daypart} - {hrs_label}"


# ---------------------------------------------------------------------------
# Transform: SpecialEventWorkupforGIS.csv
# ---------------------------------------------------------------------------

def transform_special_event_workup(
    df: pd.DataFrame,
    offset_hours: float = 5,
    default_unit_type: str = "Vehicle",
    default_staff_status: str = "On Duty",
    default_staff_agency: str = "HPD",
    default_event_status: str = "Event Active",
) -> Tuple[pd.DataFrame, list]:
    """
    Map SpecialEventWorkupforGIS.csv columns to ARCGIS_COLUMNS.
    Returns (output_df, warnings_list).
    """
    warnings_list = []
    rows = []

    # Remove fully blank rows
    df = df.replace("", pd.NA).dropna(how="all").fillna("").reset_index(drop=True)

    for i, row in df.iterrows():
        out = {}

        # unitid
        out["unitid"] = _get(df, row, "UnitId")

        # unitshift: use UnitShift if populated, else derive
        unit_shift_raw = _get(df, row, "UnitShift")
        if unit_shift_raw:
            out["unitshift"] = unit_shift_raw
        else:
            ss = _get(df, row, "ShiftStart")
            se = _get(df, row, "ShiftEnd")
            out["unitshift"] = derive_unit_shift(ss, se)

        # unitloc: always blank
        out["unitloc"] = ""

        # unittype
        ut = _get(df, row, "UnitType")
        out["unittype"] = ut if ut else default_unit_type

        out["unitsquad"]  = _get(df, row, "UnitSquad")
        out["unitradio"]  = _get(df, row, "UnitRadio")
        out["unitduties"] = _get(df, row, "UnitDuties")
        out["payrollnum"] = _get(df, row, "Payroll")

        # staffrank: normalize to approved ArcGIS values
        raw_rank = _get(df, row, "StaffRank")
        norm_rank, rank_matched = normalize_staff_rank(raw_rank)
        out["staffrank"] = norm_rank
        if not rank_matched:
            warnings_list.append(
                f"Row {i + 2}: Unknown staffrank '{raw_rank}' — kept as-is, verify against approved list."
            )

        out["staffname"]  = _get(df, row, "StaffName")
        out["staffphone"] = _get(df, row, "StaffPhone")
        out["staffemail"] = _get(df, row, "Staffemail")
        out["staffskills"]= _get(df, row, "StaffSkills")
        out["staffpay"]   = _get(df, row, "StaffPay")

        # staffstatus default
        ss_val = _get(df, row, "StaffStatus")
        out["staffstatus"] = ss_val if ss_val else default_staff_status

        out["staffduty"]  = _get(df, row, "StaffDuty")

        out["staffagency"] = "HPD"

        # datetime fields
        for arcgis_col, src_col in WORKUP_DATETIME_FIELDS.items():
            raw_val = _get(df, row, src_col)
            dt_obj, ok = parse_and_offset_datetime(raw_val, offset_hours)
            if ok:
                out[arcgis_col] = dt_obj
            else:
                out[arcgis_col] = raw_val
                if raw_val:
                    warnings_list.append(
                        f"Row {i + 2}: Could not parse {arcgis_col} value '{raw_val}'"
                    )

        out["eventstatus"] = default_event_status

        rows.append(out)

    result = pd.DataFrame(rows, columns=ARCGIS_COLUMNS)
    return result, warnings_list


# ---------------------------------------------------------------------------
# Transform: CurrentStaffingReport.csv
# ---------------------------------------------------------------------------

def transform_current_staffing_report(
    df: pd.DataFrame,
    offset_hours: float = 5,
    default_unit_type: str = "Traffic Control",
    default_staff_status: str = "On Duty",
    default_staff_agency: str = "HPD",
    default_event_status: str = "Event Active",
) -> Tuple[pd.DataFrame, list]:
    """
    Map CurrentStaffingReport.csv columns to ARCGIS_COLUMNS.
    Returns (output_df, warnings_list).
    """
    warnings_list = []
    rows = []

    # Remove fully blank rows
    df = df.replace("", pd.NA).dropna(how="all").fillna("").reset_index(drop=True)

    for i, row in df.iterrows():
        out = {}

        emp_id = _get(df, row, "EmpID")

        out["unitid"]    = emp_id
        out["unitloc"]   = ""
        out["unittype"]  = default_unit_type
        out["unitsquad"] = _get(df, row, "UnitNo")
        out["unitradio"] = _get(df, row, "RadioCallNumber")
        out["unitduties"]= ""
        out["payrollnum"]= emp_id

        # staffrank: normalize to approved ArcGIS values
        raw_rank = _get(df, row, "RankDescription")
        norm_rank, rank_matched = normalize_staff_rank(raw_rank)
        out["staffrank"] = norm_rank
        if not rank_matched:
            warnings_list.append(
                f"Row {i + 2}: Unknown staffrank '{raw_rank}' — kept as-is, verify against approved list."
            )

        # staffname: LastName, FirstName
        last  = _get(df, row, "LastName")
        first = _get(df, row, "FirstName")
        if last and first:
            full = f"{last}, {first}"
        elif last:
            full = last
        elif first:
            full = first
        else:
            full = ""
        out["staffname"] = full

        out["staffphone"] = _get(df, row, "CellPhone")
        out["staffemail"] = ""
        out["staffskills"]= ""
        out["staffpay"]   = ""
        out["staffstatus"]= default_staff_status
        out["staffduty"]  = _get(df, row, "Division")
        out["staffagency"] = "HPD"

        # unitshift: derive from shift/shiftStart/ShiftEnd
        shift_label = _get(df, row, "shift")
        ss_raw = _get(df, row, "shiftStart")
        se_raw = _get(df, row, "ShiftEnd")

        if shift_label:
            out["unitshift"] = shift_label
        else:
            out["unitshift"] = derive_unit_shift(ss_raw, se_raw)

        # datetime fields
        for arcgis_col, src_col in STAFFING_DATETIME_FIELDS.items():
            raw_val = _get(df, row, src_col)
            dt_obj, ok = parse_and_offset_datetime(raw_val, offset_hours)
            if ok:
                out[arcgis_col] = dt_obj
            else:
                out[arcgis_col] = raw_val
                if raw_val:
                    warnings_list.append(
                        f"Row {i + 2}: Could not parse {arcgis_col} value '{raw_val}'"
                    )

        out["eventstatus"] = default_event_status

        rows.append(out)

    result = pd.DataFrame(rows, columns=ARCGIS_COLUMNS)
    return result, warnings_list


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def validate_output(df: pd.DataFrame, source_df: pd.DataFrame = None) -> list:
    """
    Run validation checks. Returns list of warning strings.
    Does not raise; caller decides what to do.
    """
    issues = []

    if df.empty:
        issues.append("Output is empty — no rows were produced.")
        return issues

    # Missing required ArcGIS columns
    missing_cols = [c for c in ARCGIS_COLUMNS if c not in df.columns]
    if missing_cols:
        issues.append(f"Missing required output columns: {missing_cols}")

    # Blank unitid
    blank_unit = df["unitid"].apply(lambda v: not str(v).strip()).sum()
    if blank_unit:
        issues.append(f"{blank_unit} row(s) have blank unitid.")

    # Duplicate unitid
    uid_series = df["unitid"].astype(str).str.strip()
    non_blank = uid_series[uid_series != ""]
    dups = non_blank[non_blank.duplicated()].unique().tolist()
    if dups:
        issues.append(f"Duplicate unitid value(s) found: {dups[:10]}")

    # Blank staffname
    blank_name = df["staffname"].apply(lambda v: not str(v).strip()).sum()
    if blank_name:
        issues.append(f"{blank_name} row(s) have blank staffname.")

    # Blank unitshiftstart / unitshiftend
    for col in ("unitshiftstart", "unitshiftend"):
        if col in df.columns:
            blank = df[col].apply(lambda v: not str(v).strip()).sum()
            if blank:
                issues.append(f"{blank} row(s) have blank {col}.")

    # Unknown staffrank values (not in approved ArcGIS list)
    if "staffrank" in df.columns:
        approved_lower = {r.lower() for r in APPROVED_RANKS}
        bad_ranks = df[
            df["staffrank"].apply(
                lambda v: bool(str(v).strip()) and str(v).strip().lower() not in approved_lower
            )
        ]["staffrank"].unique().tolist()
        if bad_ranks:
            issues.append(
                f"staffrank value(s) not in approved ArcGIS list: {bad_ranks}. "
                f"These rows are included but may fail ArcGIS validation."
            )

    return issues


# ---------------------------------------------------------------------------
# Template writing
# ---------------------------------------------------------------------------

def write_to_template(
    output_df: pd.DataFrame,
    template_path: str,
) -> io.BytesIO:
    """
    Load the ArcGIS template workbook, clear data rows in 'Staff List',
    write output_df starting at row 2, preserve formatting,
    return BytesIO of the saved workbook.
    """
    wb = load_workbook(template_path)

    if "Staff List" not in wb.sheetnames:
        raise ValueError(
            "Template workbook does not contain a sheet named 'Staff List'. "
            f"Found sheets: {wb.sheetnames}"
        )

    ws = wb["Staff List"]

    # -----------------------------------------------------------------------
    # Build a map: column_name_lower -> column_index (1-based) from header row
    # -----------------------------------------------------------------------
    header_map = {}
    for cell in ws[1]:
        if cell.value is not None:
            header_map[str(cell.value).strip().lower()] = cell.column

    # Collect formatting from row 2 as a template (if data exists)
    row2_formats = {}
    for cell in ws[2]:
        row2_formats[cell.column] = {
            "font":      copy.copy(cell.font),
            "fill":      copy.copy(cell.fill),
            "border":    copy.copy(cell.border),
            "alignment": copy.copy(cell.alignment),
            "number_format": cell.number_format,
        }

    # -----------------------------------------------------------------------
    # Clear existing data rows (row 2 onward), only in used columns
    # -----------------------------------------------------------------------
    max_row = ws.max_row
    max_col = ws.max_column
    if max_row >= 2:
        for r in range(2, max_row + 1):
            for c in range(1, max_col + 1):
                ws.cell(row=r, column=c).value = None

    # -----------------------------------------------------------------------
    # Write rows
    # -----------------------------------------------------------------------
    datetime_arcgis_cols = {"unitshiftstart", "unitshiftend"}

    for row_idx, (_, data_row) in enumerate(output_df.iterrows(), start=2):
        for arcgis_col in ARCGIS_COLUMNS:
            col_idx = header_map.get(arcgis_col.lower())
            if col_idx is None:
                continue  # column not in template header

            cell = ws.cell(row=row_idx, column=col_idx)
            value = data_row.get(arcgis_col, "")

            # Apply value
            if isinstance(value, datetime):
                cell.value = value
                cell.number_format = _OPENPYXL_DATE_FORMAT
            elif value == "" or (isinstance(value, float) and pd.isna(value)):
                cell.value = None
            else:
                cell.value = value

            # Apply formatting from row 2 template
            fmt = row2_formats.get(col_idx)
            if fmt:
                if fmt["font"]:
                    cell.font = fmt["font"]
                if fmt["fill"]:
                    cell.fill = fmt["fill"]
                if fmt["border"]:
                    cell.border = fmt["border"]
                if fmt["alignment"]:
                    cell.alignment = fmt["alignment"]
                # Only override number format for non-datetime cells
                # if the template row had a format set
                if arcgis_col not in datetime_arcgis_cols and fmt["number_format"] and fmt["number_format"] != "General":
                    cell.number_format = fmt["number_format"]

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
