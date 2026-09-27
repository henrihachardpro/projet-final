@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
set "LOG=%~dp0build_log.txt"
>"%LOG%" echo ===== Construction Sufix OptimiZer =====
set /p "VERSION=" < version.txt
for /f "tokens=* delims= " %%A in ("!VERSION!") do set "VERSION=%%A"
if not defined VERSION (echo ERREUR version.txt vide & pause & exit /b 1)
set "APP_NAME=Sufix OptimiZer V!VERSION!"
for %%F in ("sufix_optimiseur.py" "sufix_ui.py" "sufix_core.py" "requirements.txt" "version.txt" "app_icon.ico" "splash_screen.png" "welcome_background.png" "sufix_profiles.xlsx" "base_article_supportage.xlsx" "data_versions.json") do if not exist %%F (echo ERREUR %%F manquant & pause & exit /b 1)
set "PYTHON_CMD="
where py >nul 2>nul && set "PYTHON_CMD=py"
if not defined PYTHON_CMD where python >nul 2>nul && set "PYTHON_CMD=python"
if not defined PYTHON_CMD (echo Python introuvable & pause & exit /b 1)
if not exist ".venv\Scripts\python.exe" %PYTHON_CMD% -m venv .venv >>"%LOG%" 2>&1
call ".venv\Scripts\activate.bat"
if errorlevel 1 goto :error
python -m pip install --upgrade pip >>"%LOG%" 2>&1
python -m pip install -r requirements.txt >>"%LOG%" 2>&1
if errorlevel 1 goto :error
echo Execution des tests...
python -m unittest discover -s tests -v >>"%LOG%" 2>&1
if errorlevel 1 (echo ECHEC TESTS - compilation annulee & goto :error)
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
del /q *.spec 2>nul
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "!APP_NAME!" --icon "app_icon.ico" ^
 --add-data "app_icon.ico;." --add-data "splash_screen.png;." --add-data "welcome_background.png;." --add-data "welcome_background_*.png;." --add-data "version.txt;." ^
 --add-data "sufix_profiles.xlsx;." --add-data "base_article_supportage.xlsx;." --add-data "data_versions.json;." ^
 --collect-all pdfplumber --collect-all pymupdf --collect-all tkinterdnd2 --hidden-import fitz sufix_optimiseur.py >>"%LOG%" 2>&1
if errorlevel 1 goto :error
echo SUCCES : dist\!APP_NAME!.exe
pause
exit /b 0
:error
echo ECHEC - voir build_log.txt
powershell -NoProfile -Command "Get-Content '%LOG%' -Tail 50" 2>nul
pause
exit /b 1
