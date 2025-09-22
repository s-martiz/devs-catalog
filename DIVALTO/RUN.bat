@echo off
setlocal

REM Active l'environnement "mdu" avec micromamba
call micromamba activate mdu

REM Vérifie si un argument est donné
if "%~1"=="" (
    echo Aucun fichier spécifié, on prendra le dernier Excel du dossier.
    python convert_temps.py
) else (
    echo Fichier indiqué : %~1
    python convert_temps.py "%~1"
)

pause
