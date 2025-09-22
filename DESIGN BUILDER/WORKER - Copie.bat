@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM === 1) Chemin d'EnergyPlus (adapter si besoin) ===========================
set "EPLUS_ROOT=C:\EnergyPlusV9-4-0"
set "EPLUS_EXE=%EPLUS_ROOT%\energyplus.exe"

if not exist "%EPLUS_EXE%" (
  echo [ERREUR] EnergyPlus introuvable : "%EPLUS_EXE%"
  pause & exit /b 1
)

REM === 2) Se placer dans le dossier du .bat =================================
pushd "%~dp0"

REM === 3) Trouver l'unique IDF ==============================================
set "IDF="
set /a IDFCOUNT=0
for %%I in ("*.idf") do (
  set "IDF=%%~fI"
  set /a IDFCOUNT+=1
)
if %IDFCOUNT%==0 (
  echo [ERREUR] Aucun fichier .idf trouve.
  popd & pause & exit /b 1
)
if not %IDFCOUNT%==1 (
  echo [ERREUR] Plusieurs .idf trouves :
  for %%I in ("*.idf") do echo   - %%~nxI
  popd & pause & exit /b 1
)

REM === 4) Trouver l'unique EPW ==============================================
set "EPW="
set /a EPWCOUNT=0
for %%W in ("*.epw") do (
  set "EPW=%%~fW"
  set /a EPWCOUNT+=1
)
if %EPWCOUNT%==0 (
  echo [ERREUR] Aucun fichier .epw trouve.
  popd & pause & exit /b 1
)
if not %EPWCOUNT%==1 (
  echo [ERREUR] Plusieurs .epw trouves :
  for %%W in ("*.epw") do echo   - %%~nxW
  popd & pause & exit /b 1
)

REM === 5) Dossier de resultats horodate =====================================
for /f "usebackq delims=" %%T in (`
  powershell -NoProfile -Command "(Get-Date).ToString('yyyy-MM-dd_HH-mm-ss')"
`) do set "STAMP=%%T"

set "OUTDIR=results-%STAMP%"
mkdir "%OUTDIR%" >nul 2>&1

REM === 6) Copier IDF et EPW dans OUTDIR =====================================
copy /Y "%IDF%" "%OUTDIR%\" >nul
copy /Y "%EPW%" "%OUTDIR%\" >nul

REM === 6b) Copier le script eso_to_csv - Copie.bat ========================== [NOUVEAU]
if exist "eso_to_csv - Copie.bat" (
  copy /Y "eso_to_csv - Copie.bat" "%OUTDIR%\" >nul
)

REM === 7) Lancer EnergyPlus avec CSV (-x) et logger =========================
set "LOG=%OUTDIR%\log.txt"
echo --------------------------------------------------------------- >  "%LOG%"
echo EnergyPlus log - %DATE% %TIME% >> "%LOG%"
echo IDF: %IDF%                         >> "%LOG%"
echo EPW: %EPW%                         >> "%LOG%"
echo OUT: %OUTDIR%                      >> "%LOG%"
echo --------------------------------------------------------------- >> "%LOG%"

REM -d : dossier de sortie
REM -w : fichier EPW
REM -x : lance ReadVarsESO pour generer eplusout.csv
"%EPLUS_EXE%" -d "%OUTDIR%" -w "%EPW%" -x "%IDF%" >> "%LOG%" 2>&1

set "ERR=%ERRORLEVEL%"
echo.                                      >> "%LOG%"
echo Code de sortie: %ERR%                 >> "%LOG%"

echo.
if not "%ERR%"=="0" (
  echo [FINI] Code de sortie %ERR%. Voir "%LOG%".
) else (
  echo [OK] Simulation terminee. Resultats (dont eplusout.csv) dans: "%OUTDIR%"
)


REM === 8) Attendre la fin d'E+ et stabilisation de eplusout.eso =============
set "END=%OUTDIR%\eplusout.end"
set "ESO=%OUTDIR%\eplusout.eso"

REM -- 8a) Attendre que eplusout.end existe et signale la fin (max ~300 s)
set /a _tries=0
:WAIT_END
if exist "%END%" (
  findstr /I "success" "%END%" >nul 2>&1 && goto WAIT_STABLE
)
set /a _tries+=1
if %_tries% GEQ 300 goto WAIT_STABLE
timeout /t 1 >nul
goto WAIT_END

REM -- 8b) Attendre que eplusout.eso existe et que sa taille se stabilise
:WAIT_STABLE
if not exist "%ESO%" (
  REM si pas d'ESO, on continue quand même (peut-être que ton batch lit le CSV)
  goto RUN_CONVERTER
)

set "prev=-1"
set /a same=0
REM Jusqu’à ~120 s ou 3 tailles identiques d’affilée (stabilité)
for /l %%# in (1,1,120) do (
  for %%A in ("%ESO%") do set "size=%%~zA"
  if "!size!"=="!prev!" (
    set /a same+=1
  ) else (
    set "prev=!size!"
    set /a same=0
  )
  if !same! GEQ 3 goto RUN_CONVERTER
  timeout /t 1 >nul
)

:RUN_CONVERTER
REM === 9) Exécuter le script eso_to_csv - Copie.bat ==========================
if exist "%OUTDIR%\eso_to_csv - Copie.bat" (
  echo [INFO] Lancement de "eso_to_csv - Copie.bat"...
  pushd "%OUTDIR%"
  call "eso_to_csv - Copie.bat"
  set "CONV_ERR=%ERRORLEVEL%"
  popd
  if not "%CONV_ERR%"=="0" (
    echo [AVERTISSEMENT] Le convertisseur a retourne %CONV_ERR%.
  ) else (
    echo [OK] Conversion eso->csv terminee.
  )
) else (
  echo [ATTENTION] Aucun fichier "eso_to_csv - Copie.bat" trouve dans "%OUTDIR%".
)

popd
pause
