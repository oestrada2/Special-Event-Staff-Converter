# ArcGIS Special Event Staff Converter

Local Streamlit app that converts HPD police staffing CSVs into the ArcGIS Special Event Solution batch upload Excel format.

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Place the Excel template

Copy the real ArcGIS batch upload template to:

```
templates/Sample_Batch_Load_Event_Staff_Template.xlsx
```

The app will not generate output without this file.

### 3. Run the app

```bash
streamlit run app.py
```

Open the URL shown in the terminal (usually `http://localhost:8501`).

---

## Supported Input Files

| File | Format |
|---|---|
| `SpecialEventWorkupforGIS.csv` | Special Event Workup export |
| `CurrentStaffingReport.csv` | Current Staffing Report (may have title rows above the header) |

Both CSV and XLSX uploads are accepted.

---

## Usage

1. Upload your staffing file using the drag-and-drop area.
2. The app auto-detects the source format.
3. Review the raw data preview and validation summary.
4. Adjust sidebar settings if needed (unit type defaults, agency, offset hours).
5. Click **Download ArcGIS Upload Workbook** to get the completed Excel file.

---

## Sidebar Settings

| Setting | Default | Description |
|---|---|---|
| Default Unit Type (Workup) | Vehicle | Fills blank `unittype` in Workup files |
| Default Unit Type (Staffing) | Traffic Control | Fills `unittype` in Staffing files |
| Default Staff Status | On Duty | Fills blank `staffstatus` |
| Default Staff Agency | HPD | Fills blank `staffagency` |
| Default Event Status | Event Active | Written to all rows |
| Time Offset Hours | 5 | Hours added to shift start/end (UTC → CST/CDT) |

---

## Output

Downloaded file: `ArcGIS_Special_Event_Staff_Upload.xlsx`

Written to the `Staff List` sheet in the ArcGIS template workbook. All other sheets (dropdowns, data dictionary) are preserved.

---

## Troubleshooting

### Template not found
```
Error: Template NOT found
```
Place `Sample_Batch_Load_Event_Staff_Template.xlsx` in the `templates/` folder and reload the app.

---

### Unsupported input format
```
Error: Could not identify the source format.
```
Verify the uploaded file is one of the two supported formats. Check that expected column names are present and spelled correctly.

---

### Date/time parsing warnings
```
Warning: Row N: Could not parse unitshiftstart value '...'
```
The source value could not be parsed as a date. The original value is kept as-is. Check the source data for non-standard date formats.

---

### Missing columns
```
Warning: Missing expected source columns: [...]
```
One or more expected source columns were not found. The output may have blank values for those fields. Check the source file column names.

---

### Empty output
```
Error: Transformed output is empty — nothing to write.
```
All rows were removed (fully blank rows are dropped). Verify the source file has data rows.

---

## Project Structure

```
arcgis-special-event-staff-converter/
├── app.py                    # Streamlit UI
├── staff_transformer.py      # All transformation logic
├── requirements.txt
├── README.md
├── .gitignore
└── templates/
    └── Sample_Batch_Load_Event_Staff_Template.xlsx  ← place your template here
```
