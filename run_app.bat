@echo off
call "C:\Program Files\ArcGIS\Pro\bin\Python\Scripts\activate.bat" staffconv
streamlit run "C:\Users\oestrada\arcgis-special-event-staff-converter\app.py"
pause
