@echo off
cd /d "%~dp0"
echo === Gerando ExBird - Conversor ===
python --version >nul 2>&1
if errorlevel 1 (
  echo Python nao encontrado. Instale em python.org marcando Add Python to PATH e rode de novo.
  pause
  exit /b 1
)
python -m venv .venv
call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install openpyxl firebirdsql passlib pyinstaller
pyinstaller --noconfirm --clean --onefile --noconsole --name ExBird --collect-submodules firebirdsql --collect-all passlib app.py
if errorlevel 1 (
  echo Deu erro ao gerar o exe.
  pause
  exit /b 1
)
copy /Y dist\ExBird.exe . >nul
echo.
echo Pronto: ExBird.exe criado nesta pasta.
pause
