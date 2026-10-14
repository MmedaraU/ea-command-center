@echo off
set PATH=C:\msys64\ucrt64\bin;%PATH%
set WEASYPRINT_DLL_DIRECTORIES=C:\msys64\ucrt64\bin
venv\Scripts\python.exe -m streamlit run app.py