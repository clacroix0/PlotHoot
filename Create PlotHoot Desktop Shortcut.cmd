@echo off
setlocal

set "APP_FOLDER="
set "SCRIPT="
call :TryFolder "%~dp0PlotHoot App Files"
for /d %%D in ("%~dp0*App Files") do call :TryFolder "%%~fD"
call :TryFolder "%~dp0"

set "PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS%" (
    set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

if not defined SCRIPT (
    echo Could not find the PlotHoot app file.
    echo Keep this file beside the PlotHoot App Files folder, or inside the app files folder.
    pause
    exit /b 1
)

pushd "%APP_FOLDER%"
if errorlevel 1 (
    echo Could not open "%APP_FOLDER%".
    pause
    exit /b 1
)

set "PLOTHOOT_SCRIPT=%SCRIPT%"
set "PLOTHOOT_APPROOT=%APP_FOLDER%"
"%PS%" -NoProfile -ExecutionPolicy Bypass -STA -Command "Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $p=$env:PLOTHOOT_SCRIPT; $s=[System.IO.File]::ReadAllText($p); . ([scriptblock]::Create($s)) @args" -RepairShortcuts
set "APP_EXIT=%ERRORLEVEL%"
popd

echo.
if "%APP_EXIT%"=="0" (
    echo PlotHoot shortcuts were checked for this Windows account.
    echo If the desktop icon does not refresh right away, sign out and back in or refresh the desktop.
) else (
    echo PlotHoot could not create the shortcut for this account.
    echo If Windows blocks this, IT may be blocking PowerShell or Windows Script Host for this account.
)
echo.
pause

endlocal
exit /b %APP_EXIT%

:TryFolder
if defined SCRIPT exit /b
set "CANDIDATE=%~1"
if not exist "%CANDIDATE%" exit /b
if exist "%CANDIDATE%\PlotHoot.ps1" (
    set "APP_FOLDER=%CANDIDATE%"
    set "SCRIPT=%CANDIDATE%\PlotHoot.ps1"
    exit /b
)
if exist "%CANDIDATE%\Plot*.ps1" (
    for %%S in ("%CANDIDATE%\Plot*.ps1") do (
        if not defined SCRIPT (
            set "APP_FOLDER=%CANDIDATE%"
            set "SCRIPT=%%~fS"
        )
    )
)
exit /b
