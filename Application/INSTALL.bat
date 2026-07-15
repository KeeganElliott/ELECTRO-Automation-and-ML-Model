@echo off
cd /d "%~dp0"
py -m pip install --upgrade numpy pandas scipy matplotlib PyWavelets openpyxl pywin32 pdfplumber pymupdf pillow pytesseract scikit-learn joblib xlsxwriter
pause
