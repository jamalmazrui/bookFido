@echo off
setlocal EnableExtensions
rem createbookFidoRepo.cmd -- thin wrapper that runs the PowerShell
rem worker and captures everything to createbookFidoRepo.log, in the
rem manner of the createUrlFido pipeline.  PowerShell does the git and gh
rem work, because invoking gh from a batch file falls into a batch trap
rem when gh is installed as a .cmd shim: without "call", control transfers
rem to the shim and never returns, ending the script silently.

if not defined sLogging (
    set "sLogging=1"
    cmd /d /c ""%~f0"" %* > "%~dp0createbookFidoRepo.log" 2>&1
    type "%~dp0createbookFidoRepo.log"
    exit /b
)

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0createbookFidoRepo.ps1"
exit /b %errorlevel%
