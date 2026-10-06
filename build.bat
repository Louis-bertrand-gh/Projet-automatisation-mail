@echo off
REM =============================================================
REM  Fabrique l'application .exe "Assistant Arrivees Tardives"
REM  A executer sur le PC Windows du camping (une seule fois,
REM  puis a nouveau seulement si le code est modifie).
REM =============================================================

echo.
echo === 1/3 - Verification de Python ===
python --version
if errorlevel 1 (
    echo.
    echo ERREUR : Python n'est pas installe ou n'est pas dans le PATH.
    echo Installez Python depuis https://www.python.org/downloads/
    echo puis cochez "Add Python to PATH" lors de l'installation.
    pause
    exit /b 1
)

echo.
echo === 2/3 - Installation des dependances ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo === 3/3 - Fabrication du fichier .exe ===
pyinstaller --noconfirm --onefile --windowed ^
    --name "AssistantArriveesTardives" ^
    --version-file "version_info.txt" ^
    --add-data "plans;plans" ^
    main.py
if errorlevel 1 (
    echo.
    echo ERREUR : la construction a echoue.
    echo Fermez d'abord toute instance de l'ancien .exe, puis relancez ce script.
    pause
    exit /b 1
)

echo.
echo =========================================================
echo   Termine ! Le fichier .exe se trouve dans le dossier
echo   "dist\AssistantArriveesTardives.exe"
echo.
echo   Copiez ce .exe (et un dossier "plans" a cote de lui)
echo   a l'endroit ou vous voulez l'utiliser sur le PC de
echo   l'accueil.
echo =========================================================
pause
