ELECTRO SHORT-PATH FIXED BUNDLE

1. Extract this folder to a short path, preferably C:\ELECTRO_App.
2. Keep every included .py file together.
3. Run INSTALL.bat once.
4. Run RUN.bat to start the application.

Fixes in this build:
- All companion modules use short filenames, preventing Windows Explorer from skipping files because of path length.
- main.py calls the other scripts internally in the same Python interpreter.
- PDF keyword highlighting now tries embedded search, embedded words, and OCR bounding boxes.
- Per-design artifacts are stored in ELECTRO_Data_Collected\<design_id>.
- Audit and separate model-ready CSV/XLSX files are retained.

Tesseract OCR must be installed separately for OCR highlighting/fallback. If it is not on PATH, set TESSERACT_EXE inside creepage.py.
