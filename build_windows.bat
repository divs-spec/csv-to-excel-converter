@echo off
setlocal
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
py -m pip install pyinstaller
py -m PyInstaller --noconfirm --clean --windowed --onedir --name CSVtoExcel ^
  --collect-all charset_normalizer csv_to_excel.py
echo.
echo Build complete. Run dist\CSVtoExcel\CSVtoExcel.exe
pause
