# createbookFidoRepo.ps1 -- establishes the bookFido repo on github.com
# as JamalMazrui/bookFido from the folder this script lives in (normally
# C:\bookFido).  If a JamalMazrui/GetAudibleInfo repo exists on github.com,
# it is RENAMED to bookFido (github redirects the old name), since the
# GetAudibleInfo name is retired; otherwise a fresh public repo is created.
# Requires git and an authenticated gh; run "gh auth login" once beforehand.
# Idempotent: safe to rerun.

# ErrorActionPreference stays Continue: under Windows PowerShell 5.1 with
# the Stop preference, redirecting a native command's error stream (as the
# probes below do with *>) wraps any stderr line in a terminating
# NativeCommandError, so a mere "repo not found" probe kills the script.
# Success is judged by $LASTEXITCODE after every call instead.
$ErrorActionPreference = "Continue"
Set-Location -Path $PSScriptRoot
Write-Output ("[INFO] Repo setup started " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
Write-Output ("[INFO] Working directory: " + (Get-Location).Path)

foreach ($sTool in @("git", "gh")) {
    if (-not (Get-Command $sTool -ErrorAction SilentlyContinue)) {
        Write-Output ("[ERROR] " + $sTool + " was not found on the PATH.")
        exit 1
    }
}

& gh auth status *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Output "[ERROR] gh is not authenticated; run: gh auth login"
    exit 1
}

# If the retired GetAudibleInfo repo exists on github.com, rename it to
# bookFido; the rename keeps stars, issues, and history, and github
# forwards the old url.  If bookFido already exists, nothing to rename.
& gh repo view JamalMazrui/bookFido *> $null
if ($LASTEXITCODE -eq 0) {
    Write-Output "[INFO] The bookFido repo already exists on github.com"
} else {
    & gh repo view JamalMazrui/GetAudibleInfo *> $null
    if ($LASTEXITCODE -eq 0) {
        Write-Output "[INFO] Renaming the retired GetAudibleInfo repo to bookFido"
        & gh repo rename bookFido --repo JamalMazrui/GetAudibleInfo --yes
        if ($LASTEXITCODE -ne 0) { exit 1 }
    } else {
        Write-Output "[INFO] No existing repo found; a fresh bookFido repo will be created at push time"
    }
}

if (-not (Test-Path ".gitignore")) {
    Write-Output "[INFO] Writing .gitignore"
    $lIgnore = @(
        "bookFido.log",
        "bookFido.json",
        "buildbookFido.log",
        "createbookFidoRepo.log",
        "GetAudibleInfo.json",
        "GetAudibleInfo_state.json",
        "bookFido_fail_*.html",
        "GetAudibleInfo_fail_*.html",
        "nugetPins.txt",
        "bookFido.db",
        "e_sqlite3.dll",
        "Version.cs",
        "*.dll",
        "tagRelease.cmd",
        "tagRelease.ps1",
        "tagRelease.log"
    )
    Set-Content -Path ".gitignore" -Value ($lIgnore -join "`r`n") -Encoding Ascii
}

if (-not (Test-Path ".git")) {
    Write-Output "[INFO] Initializing the local repository"
    & git init -b main
    if ($LASTEXITCODE -ne 0) { exit 1 }
} else {
    Write-Output "[INFO] The local repository already exists"
}

& git add -A
if ($LASTEXITCODE -ne 0) { exit 1 }

& git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Output "[INFO] Nothing new to commit"
} else {
    & git commit -m "bookFido, renamed from GetAudibleInfo"
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

& git remote get-url origin *> $null
if ($LASTEXITCODE -ne 0) {
    & gh repo view JamalMazrui/bookFido *> $null
    if ($LASTEXITCODE -eq 0) {
        Write-Output "[INFO] Wiring origin to the existing bookFido repo and pushing"
        & git remote add origin https://github.com/JamalMazrui/bookFido.git
        if ($LASTEXITCODE -ne 0) { exit 1 }
        & git push -u origin main
        if ($LASTEXITCODE -ne 0) { exit 1 }
    } else {
        Write-Output "[INFO] Creating https://github.com/JamalMazrui/bookFido and pushing"
        & gh repo create JamalMazrui/bookFido --public --source . --remote origin --push --description "Catalogs an Audible library into accessible HTML, Markdown, and Excel, downloading companion PDFs and converting them to accessible HTML.  Part of the Homer Tools series."
        if ($LASTEXITCODE -ne 0) { exit 1 }
    }
} else {
    Write-Output "[INFO] The origin remote already exists; pointing it at bookFido and pushing"
    & git remote set-url origin https://github.com/JamalMazrui/bookFido.git
    if ($LASTEXITCODE -ne 0) { exit 1 }
    & git push -u origin main
    if ($LASTEXITCODE -ne 0) { exit 1 }
}

Write-Output "[INFO] Done.  The repo is at https://github.com/JamalMazrui/bookFido"
