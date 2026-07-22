@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
rem ====================================================================
rem buildbookFido.cmd -- builds bookFido.exe as a single
rem 64-bit GUI exe (no console window) for .NET Framework 4.8.
rem
rem All output is captured to buildbookFido.log beside this
rem script: the script relaunches itself once with output redirected,
rem then types the log to the console when the build ends.
rem
rem The PDF-to-HTML feature uses the UglyToad.PdfPig library
rem (Apache-2.0, pure managed code, no Office dependency).  On first
rem build, the PdfPig assemblies and their System helper assemblies
rem are fetched from nuget.org (internet needed that first time only),
rem then embedded into the exe as manifest resources with csc
rem /resource, the same single-file technique 2htm uses for Markdig.
rem Each nupkg is saved with a .zip extension because PowerShell's
rem Expand-Archive refuses any other extension -- the exact detail
rem that made the first version of this script fail silently.
rem The resulting bookFido.exe needs NO dll files at runtime.
rem
rem A Roslyn csc is preferred when present; the in-box .NET Framework
rem compiler is accepted as a fallback because the source is kept
rem compatible with C# 5.
rem ====================================================================

rem ---- Log wrapper: relaunch once with all output going to the log ----
if not defined bookFidoBuildLogging (
    set "bookFidoBuildLogging=1"
    cmd /c ""%~f0"" > "buildbookFido.log" 2>&1
    set "iExitCode=!errorlevel!"
    type "buildbookFido.log"
    echo [INFO] This output was also saved to buildbookFido.log
    exit /b !iExitCode!
)
echo [INFO] Build started %date% %time%

rem ---- Locate a C# compiler (Roslyn preferred, Framework fallback) ----
set "sCsc="
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe" set "sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe"
if not defined sCsc for %%E in (Community Professional Enterprise) do (
    if not defined sCsc if exist "C:\Program Files\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" set "sCsc=C:\Program Files\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
    if not defined sCsc if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" set "sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2022\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
)
if not defined sCsc if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe" set "sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe"
if not defined sCsc for %%E in (Community Professional Enterprise) do (
    if not defined sCsc if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\%%E\MSBuild\Current\Bin\Roslyn\csc.exe" set "sCsc=C:\Program Files (x86)\Microsoft Visual Studio\2019\%%E\MSBuild\Current\Bin\Roslyn\csc.exe"
)
if not defined sCsc if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
    set "sCsc=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    echo [INFO] Roslyn csc.exe was not found; using the in-box Framework compiler, which is fine for this source.
)
if not defined sCsc (
    echo [ERROR] No C# compiler was found.  Install .NET Framework 4.8 or Visual Studio Build Tools.
    exit /b 2
)
echo [INFO] Using compiler: %sCsc%

rem ---- Locate the netstandard facade (needed for netstandard2.0 dlls) ----
set "sNetstd="
if exist "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8.1\Facades\netstandard.dll" set "sNetstd=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8.1\Facades\netstandard.dll"
if not defined sNetstd if exist "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\Facades\netstandard.dll" set "sNetstd=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\Facades\netstandard.dll"
if not defined sNetstd if exist "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\Facades\netstandard.dll" set "sNetstd=C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.7.2\Facades\netstandard.dll"
if not defined sNetstd if exist "%SystemRoot%\Microsoft.NET\assembly\GAC_MSIL\netstandard\v4.0_2.0.0.0__cc7b13ffcd2ddd51\netstandard.dll" set "sNetstd=%SystemRoot%\Microsoft.NET\assembly\GAC_MSIL\netstandard\v4.0_2.0.0.0__cc7b13ffcd2ddd51\netstandard.dll"
if not defined sNetstd (
    echo [ERROR] netstandard.dll facade not found.  Repair or install .NET Framework 4.8.
    exit /b 2
)
echo [INFO] netstandard facade: %sNetstd%

rem ---- Fetch the PdfPig assemblies on first build ---------------------
rem The official package id is plain "pdfpig" -- the project's README
rem says Install-Package PdfPig -- and that one package contains all
rem the UglyToad assemblies.  (The NuGet id "UglyToad.PdfPig" is an
rem unrelated third-party upload and must NOT be used.)  PdfPig 0.1.14
rem stable declares net462 dependencies on Microsoft.Bcl.HashCode,
rem System.Memory, and System.ValueTuple, so those helpers are fetched
rem too.  System.Memory must be 4.6.0 or later: PdfPig 0.1.14 was
rem compiled against System.Memory assembly version 4.0.2.0, and the
rem 4.5.x packages carry only 4.0.1.x, which csc rejects with CS1705.
rem A nugetPins.txt stamp records the fetched set, so editing any pin
rem above makes the next build refetch instead of reusing stale dlls.  Packages come from the official NuGet flat-container url (the
rem DbDo build's approach) and are unpacked with PowerShell Expand-Archive,
rem which requires the archive to carry a .zip extension, so each
rem downloaded nupkg is saved as <package>.zip.  From each package the
rem best lib target for .NET Framework 4.8 is copied: net462 first,
rem then net461, net46, net45, then netstandard2.0.  The System helper
rem packages cover PdfPig's possible dependencies; any that a given
rem PdfPig version does not need are simply absent after extraction
rem and are skipped at compile time.
set "sPins=pdfpig@0.1.14 epplus@4.5.3.3 microsoft.bcl.hashcode@6.0.0 system.text.encoding.codepages@4.5.1 system.buffers@4.5.1 system.memory@4.6.0 system.numerics.vectors@4.5.0 system.runtime.compilerservices.unsafe@6.0.0 system.valuetuple@4.5.0"
set "sStamp="
set "sStampWant=%sPins% layout-net40"
if exist "nugetPins.txt" set /p sStamp=<"nugetPins.txt"
if "%sStamp%"=="%sStampWant%" (
    echo [INFO] The pinned packages are already present, so the fetch step is skipped.
    goto :dllsReady
)
echo [INFO] Fetching the PdfPig assemblies from nuget.org ...
del /q UglyToad.*.dll EPPlus.dll Microsoft.Bcl.HashCode.dll System.Text.Encoding.CodePages.dll System.Buffers.dll System.Memory.dll System.Numerics.Vectors.dll System.Runtime.CompilerServices.Unsafe.dll System.ValueTuple.dll 2>nul
set "sTempDir=%TEMP%\bookFido_nuget_%RANDOM%%RANDOM%"
mkdir "%sTempDir%" 2>nul
if not exist "%sTempDir%" (
    echo [ERROR] Could not create temp directory %sTempDir%.
    exit /b 2
)
for %%K in (%sPins%) do (
    for /f "tokens=1,2 delims=@" %%A in ("%%K") do (
        echo [INFO] Fetching %%A %%B ...
        curl -sL -o "!sTempDir!\%%A.zip" "https://api.nuget.org/v3-flatcontainer/%%A/%%B/%%A.%%B.nupkg"
        if exist "!sTempDir!\%%A.zip" (
            powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Path '!sTempDir!\%%A.zip' -DestinationPath '!sTempDir!\%%A' -Force"
            set "sLibDir="
            if exist "!sTempDir!\%%A\lib\net462" set "sLibDir=!sTempDir!\%%A\lib\net462"
            if not defined sLibDir if exist "!sTempDir!\%%A\lib\net461" set "sLibDir=!sTempDir!\%%A\lib\net461"
            if not defined sLibDir if exist "!sTempDir!\%%A\lib\net46" set "sLibDir=!sTempDir!\%%A\lib\net46"
            if not defined sLibDir if exist "!sTempDir!\%%A\lib\net45" set "sLibDir=!sTempDir!\%%A\lib\net45"
            if not defined sLibDir if exist "!sTempDir!\%%A\lib\net40" set "sLibDir=!sTempDir!\%%A\lib\net40"
            if not defined sLibDir if exist "!sTempDir!\%%A\lib\netstandard2.0" set "sLibDir=!sTempDir!\%%A\lib\netstandard2.0"
            if defined sLibDir (
                echo [INFO] %%A: taking dlls from !sLibDir!
                copy /y "!sLibDir!\*.dll" . >nul
            )
            if not defined sLibDir (
                echo [WARN] No usable lib folder found inside %%A %%B.  Package layout:
                dir /s /b "!sTempDir!\%%A\lib" 2>nul
            )
        ) else (
            echo [WARN] Could not download %%A %%B.
        )
    )
)
rmdir /s /q "%sTempDir%" 2>nul
(echo %sStampWant%)>"nugetPins.txt"
:dllsReady
if not exist "UglyToad.PdfPig.dll" (
    echo [ERROR] UglyToad.PdfPig.dll could not be obtained, so the build cannot proceed.
    echo         Check your internet connection, or place the PdfPig dlls beside this script.
    exit /b 2
)
echo [INFO] Assemblies present for embedding:
dir /b UglyToad.*.dll 2>nul
dir /b EPPlus.dll Microsoft.Bcl.HashCode.dll System.Buffers.dll System.Text.Encoding.CodePages.dll System.Memory.dll System.Numerics.Vectors.dll System.Runtime.CompilerServices.Unsafe.dll System.ValueTuple.dll 2>nul

rem ---- Assemble the reference and embed switches for each dll present ----
set "sExtra="
for %%D in (UglyToad.*.dll) do set "sExtra=!sExtra! /reference:%%D /resource:%%D,%%D"
for %%D in (EPPlus.dll Microsoft.Bcl.HashCode.dll System.Buffers.dll System.Memory.dll System.Numerics.Vectors.dll System.Runtime.CompilerServices.Unsafe.dll System.Text.Encoding.CodePages.dll System.ValueTuple.dll) do (
    if exist "%%D" set "sExtra=!sExtra! /reference:%%D /resource:%%D,%%D"
)

rem ---- Compile --------------------------------------------------------
if exist bookFido.exe del bookFido.exe
set "sIcon="
if exist "bookFido.ico" set "sIcon=/win32icon:bookFido.ico"
"%sCsc%" /nologo %sIcon% /target:winexe /platform:x64 /optimize+ /out:bookFido.exe /reference:System.dll /reference:System.Core.dll /reference:System.Web.Extensions.dll /reference:System.Windows.Forms.dll /reference:"%sNetstd%" %sExtra% bookFido.cs
if not exist bookFido.exe (
    echo [ERROR] Build failed.
    exit /b 1
)
echo [INFO] Built bookFido.exe successfully as a single-file exe with PdfPig and EPPlus embedded.
echo [INFO] Build finished %date% %time%
endlocal
