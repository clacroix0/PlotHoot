@echo off
setlocal

set "APP_FOLDER="
set "SCRIPT="
call :TryFolder "%~dp0"

set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS32%" (
    set "PS32=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

if not defined SCRIPT (
    echo Could not find the PlotHoot app file.
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
"%PS32%" -NoProfile -ExecutionPolicy Bypass -STA -Command "Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $p=$env:PLOTHOOT_SCRIPT; $s=[System.IO.File]::ReadAllText($p); . ([scriptblock]::Create($s)) @args" %*
set "APP_EXIT=%ERRORLEVEL%"
popd

if not "%APP_EXIT%"=="0" (
    echo.
    echo PlotHoot closed because of an error.
    echo If Windows still blocks this app, IT may be blocking PowerShell itself for this account.
    echo.
    pause
)

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
