@echo off
echo ========================================
echo  AI Agent - Office Automation Setup
echo ========================================
echo.

echo [1/3] Installing Python dependencies...
pip install PySide6 openai python-dotenv python-docx python-pptx openpyxl pdfplumber pdf2image Pillow docx2pdf
echo.

echo [2/3] Checking installations...
python -c "import PySide6; print('  PySide6 OK')"
python -c "import openai; print('  OpenAI SDK OK')"
python -c "import docx; print('  python-docx OK')"
python -c "import pptx; print('  python-pptx OK')"
python -c "import openpyxl; print('  openpyxl OK')"
python -c "import pdfplumber; print('  pdfplumber OK')"
echo.

echo [3/3] Setup complete!
echo.
echo To run the app:
echo   python main.py
echo.
pause
