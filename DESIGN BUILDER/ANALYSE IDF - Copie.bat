@echo off
REM Activation de l'environnement Python
call activate mdu

REM Exécution du script Python
python parse_idf_to_excel.py

REM Pause pour vérifier si tout a fonctionné
pause