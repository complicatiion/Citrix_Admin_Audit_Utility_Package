@echo off
setlocal EnableExtensions

set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell.exe"

set "SCRIPT_DIR=%~dp0"
set "HELPER=%SCRIPT_DIR%Citrix_Admin_Audit_Helper.ps1"

title Citrix Admin Audit Utility
color 0B
chcp 65001 >nul

net session >nul 2>&1
if errorlevel 1 (
  set "ISADMIN=0"
) else (
  set "ISADMIN=1"
)

set "REPORTROOT=%USERPROFILE%\Desktop\CitrixReports"
if not exist "%REPORTROOT%" md "%REPORTROOT%" >nul 2>&1

if not exist "%HELPER%" (
  cls
  echo ============================================================
  echo Citrix Admin Audit Utility by complicatiion
  echo ============================================================
  echo.
  echo Helper file not found:
  echo %HELPER%
  echo.
  echo Make sure the BAT and PS1 files are kept in the same folder.
  echo.
  pause
  goto END
)

set "ADMINADDR="

:MAIN
cls
echo ============================================================
echo.
echo    CCCCCC   IIIIIII TTTTT  RRRRRR   IIIIIII X   X
echo   CC          III     T    RR   RR    III    X X
echo   CC          III     T    RRRRRR     III     X
echo   CC          III     T    RR  RR     III    X X
echo    CCCCCC   IIIIIII   T    RR   RR  IIIIIII X   X
echo.
echo     Citrix Admin Audit Utility by complicatiion
echo.
echo ============================================================
echo.
if "%ISADMIN%"=="1" (
  echo Admin status : YES
) else (
  echo Admin status : NO
)
if defined ADMINADDR (
  echo AdminAddress : %ADMINADDR%
) else (
  echo AdminAddress : Localhost default
)
echo Report folder : %REPORTROOT%
echo.
echo [1] Quick audit (Instances, Services, Tools)
echo [2] Site and license overview
echo [3] Controllers and core service status
echo [4] Machine catalogs and delivery groups
echo [5] Machine registration, maintenance and power state
echo [6] Sessions and connected users
echo [7] Policy rules and entitlements
echo [8] MCS, provisioning and hosting
echo [9] Local Citrix, VDA and App Layering checks
echo [A] Local Citrix services and event review
echo [B] Create full report
echo [C] Open report folder
echo [S] Set or clear AdminAddress target
echo [0] Exit
echo.
set "CHO="
set /p CHO="Selection: "

if "%CHO%"=="1" call :RUNACTION QuickAudit
if "%CHO%"=="2" call :RUNACTION SiteOverview
if "%CHO%"=="3" call :RUNACTION Controllers
if "%CHO%"=="4" call :RUNACTION Catalogs
if "%CHO%"=="5" call :RUNACTION Machines
if "%CHO%"=="6" call :RUNACTION Sessions
if "%CHO%"=="7" call :RUNACTION Policies
if "%CHO%"=="8" call :RUNACTION MCS
if "%CHO%"=="9" call :RUNACTION LocalChecks
if /I "%CHO%"=="A" call :RUNACTION LocalEvents
if /I "%CHO%"=="B" goto REPORT
if /I "%CHO%"=="C" goto OPENFOLDER
if /I "%CHO%"=="S" goto SETADMIN
if "%CHO%"=="0" goto END
goto MAIN

:RUNACTION
cls
echo ============================================================
echo Running %~1 ...
echo ============================================================
echo.
if defined ADMINADDR (
  "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%HELPER%" -Action "%~1" -AdminAddress "%ADMINADDR%"
) else (
  "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%HELPER%" -Action "%~1"
)
echo.
pause
goto MAIN

:SETADMIN
cls
echo ============================================================
echo Set or clear AdminAddress
echo ============================================================
echo.
echo Leave blank to use the local host default.
echo Example: ctxddc01.domain.local
echo.
if defined ADMINADDR echo Current value: %ADMINADDR%
if not defined ADMINADDR echo Current value: Localhost default
echo.
set "NEWADMIN="
set /p NEWADMIN="Enter AdminAddress: "
set "ADMINADDR=%NEWADMIN%"
echo.
if defined ADMINADDR (
  echo AdminAddress set to: %ADMINADDR%
) else (
  echo AdminAddress cleared. Localhost default will be used.
)
echo.
pause
goto MAIN

:REPORT
cls
echo ============================================================
echo Creating report ...
echo ============================================================
echo.
set "STAMP=%DATE%_%TIME%"
set "STAMP=%STAMP:/=-%"
set "STAMP=%STAMP:\=-%"
set "STAMP=%STAMP::=-%"
set "STAMP=%STAMP:.=-%"
set "STAMP=%STAMP:,=-%"
set "STAMP=%STAMP: =0%"
set "OUTFILE=%REPORTROOT%\Citrix_Admin_Audit_Report_%STAMP%.txt"

if defined ADMINADDR (
  "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%HELPER%" -Action "FullReport" -AdminAddress "%ADMINADDR%" -ReportPath "%OUTFILE%"
) else (
  "%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -File "%HELPER%" -Action "FullReport" -ReportPath "%OUTFILE%"
)
echo.
echo Report file:
echo %OUTFILE%
echo.
pause
goto MAIN

:OPENFOLDER
start "" explorer.exe "%REPORTROOT%"
goto MAIN

:END
endlocal
exit /b 0
