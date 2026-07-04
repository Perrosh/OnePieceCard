@echo off
cd /d "%~dp0"

echo Avvio dashboard One Piece Card Collection...
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERRORE: ambiente virtuale non trovato.
    echo Crea prima il venv oppure installa le dipendenze.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" -m streamlit run src\one_piece_app.py

pause