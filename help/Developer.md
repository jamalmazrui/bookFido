---
title: "bookFido Developer Notes"
author: "Jamal Mazrui"
---

# bookFido Developer Notes

How bookFido is built, released and laid out, with the facts that cost the most to learn. Since the move to the Homer Development Kit it is built on HomerDev, kit 1.43.4 or later, in `C:\HomerDev`.

## The project folder

`C:\bookFido` mirrors the installed tree:

- At the top: `bookFido.cs` (the program), `build.cmd`, `bookFido_setup.iss`, `bookFido.ico`, `accept.inix`, `RepoFiles.txt`, `LocalFiles.txt`, `ReadMe` and `License`.
- `exec` — the built `bookFido.exe`, and beside it copies of the managed libraries it embeds (see below). Never in git.
- `help` — this document and the others: `bookFido` (the guide), `Announce`, `Developer`, `History`, `Hotkeys`, each as `.md` and `.htm`.
- `logs` — one log per run of the build or any tool.
- `scripts` — the kit's tools, refreshed from `C:\HomerDev\scripts` by every build.
- `work\nuget` — the pinned NuGet libraries, fetched once. Never in git.

`version.txt` holds the version and lives only on this machine; never ship it in a zip, which once reset the number on every unzip. The build writes `Version.cs` from it, and the installer reads it directly.

The program's own data is NOT in the project: `bookFido.db`, `bookFido.json` and diagnostic pages live in `%LOCALAPPDATA%\bookFido\data` (Homer's `Paths.data()`), and each run's log in `%LOCALAPPDATA%\bookFido\logs`. The first run of the kit version moves the database and the state file there from beside the program, the project folder, or `%LOCALAPPDATA%\bookFido`, taking the newest copy.

## The four steps

1. `build` — steps the version (`build nobump` keeps it), fetches the libraries if the pins changed, compiles `exec\bookFido.exe`, writes each `.htm` from its `.md`, puts the project's files in the Homer encoding, checks that the installer ships every file in `help`, and builds `bookFido_setup.exe`. Its log is `logs\bookFido-build-yyyyMMdd-HHmmss.log`.
2. `scripts\push "message"` — rewrites the whitelist `.gitignore` from `RepoFiles.txt`, commits and pushes.
3. `scripts\tidy` and `scripts\tidy --do-it` — the periodic clean.
4. `scripts\release` — runs `scripts\check`, then tags the pushed commit with the version stamped in `bookFido_setup.exe` and publishes the installer. The published installer is always at the releases page's `latest/download/bookFido_setup.exe`.

Try the fresh build with `exec\bookFido.exe`.

## The compiler and the kit's classes

The build finds the Roslyn C# compiler with `vswhere`, or installs the free Visual Studio Build Tools with winget; the Framework's own `csc.exe` stops at C# 5 and cannot compile the kit. `bookFido.cs` itself is still written in C# 5 syntax.

bookFido compiles eight of the kit's classes straight from `C:\HomerDev\CSharp`: Elevate, Inix, Lbc, Log, Paths, Say, Util and Web. The copies of `Lbc.cs` and `Say.cs` that bookFido used to carry are deleted by the build once the kit's are there. Lbc needs Elevate (its Help box offers the update), Inix, Log, Paths, Say and Util; Elevate needs Web.

- `Log.start("bookFido")` opens the session log first thing in `Main`; bookFido's own `log()` writes every line through `Log.line`, after replacing the user-profile path with `%USERPROFILE%`.
- `Paths.data()` is `dataDir()`.
- `Elevate.configure("jamalmazrui", "bookFido", version)` at startup; F11 is claimed through the opening dialog's `commandKey`.

## The embedded libraries

Everything bookFido uses travels inside `bookFido.exe` as a resource, loaded by `resolveEmbeddedAssembly`. The pins, each for a reason:

- **pdfpig 0.1.14.** The official id is plain `pdfpig`; the NuGet id `UglyToad.PdfPig` is an unrelated upload and must never be used.
- **epplus 4.5.3.3.** The last LGPL release. `lib\net40` is taken.
- **system.memory 4.6.0.** PdfPig 0.1.14 wants assembly 4.0.2.0; the 4.5.x packages carry only 4.0.1.x, which csc rejects with CS1705.
- **Helpers:** microsoft.bcl.hashcode 6.0.0, system.buffers 4.5.1, system.numerics.vectors 4.5.0, system.runtime.compilerservices.unsafe 6.0.0, system.text.encoding.codepages 4.5.1, system.valuetuple 4.5.0.
- **The SQLite side:** microsoft.data.sqlite.core 8.0.6 with sqlitepclraw core, provider.e_sqlite3, bundle_e_sqlite3 and lib.e_sqlite3 at 2.1.8. `batteries_v2` is required, because Microsoft.Data.Sqlite looks for it by name. The native `e_sqlite3.dll` is embedded as a plain resource and written out before the first open; the program then calls `SetDllDirectory`, `LoadLibraryW` and `Batteries_V2.Init`.

The general packages give the dlls of their best lib folder for .NET Framework 4.8 (net462, net461, net46, net45, net40, then netstandard2.0). The SQLite packages give each named file found by searching the package, preferring a netstandard2.0 or win-x64 copy, never a guessed path. The netstandard facade is referenced for the netstandard2.0 libraries.

**The assembly-instance trap, which cost about six rounds.** A resolver that calls `Assembly.Load(bytes)` mints a separate instance per request, each with its own statics. The SQLite engine was registered on one copy while Microsoft.Data.Sqlite, asking for a slightly different version, got a second whose provider field was null ("You need to call SQLitePCL.raw.SetProvider()"). The cure: cache by simple name, and when `<name>.dll` exists beside the program, `Assembly.LoadFrom` that path, which returns the instance already in memory. .NET Framework accepts a version-mismatched assembly returned from AssemblyResolve, which is what hid the duplication. That is why the build copies each managed library into `exec` beside the program: every verified database run loaded them from beside the program.

## Other traps

- **PowerShell 5.1:** `$ErrorActionPreference = "Stop"` with a native program's stderr redirected turns any stderr line into a terminating NativeCommandError. Use Continue and `$LASTEXITCODE`.
- **JavaScriptSerializer** round-trips `List<string>` as `object[]`; `MaxJsonLength` is raised to 50,000,000.
- **Kindle** returns all authors as one colon-joined string, names repeated, sometimes with subtitle junk: split on ":", swap each, de-duplicate, discard anything over 50 characters or containing "?".
- **Focus:** MessageBoxTimeoutW plus FindWindowW and forced foreground. Call PeekMessageW first (a winexe thread has no queue until then), then AttachThreadInput, SetForegroundWindow, BringWindowToTop and SetFocus. Plain `MessageBox.Show` boxes need `focusWhenShown()`.
- **In-place sign-in forms:** sites can show a login form without redirecting; `loginFormShowing()` probes for a password field. Kindle answers 401 or 403 when signed out.
- **NLS lives on nlsbard.loc.gov**, not bard.loc.gov, and its reading history is drawn after the page reports loaded, so every scan waits for the book count to settle.
- **Audible** never returns publisher_name, so the publisher is the first Open Library print publisher. Wikimedia wants a contactable user agent.
- **Bookshare's history filter defaults to a rolling one-month window**, so the history url sends startDate=01/01/2000 and today's endDate. Its newest-first list lets a walk stop early at a page of known books; Goodreads rows change, so its walk does not.

## Conventions

- Camel Type for C#, as `C:\HomerDev\help\CamelType_CSharp.md` describes. Constants added since the move to the kit carry the `c_` prefix.
- Titles order ignoring a leading A, An or The; authors by surname. Spreadsheets carry one region at A1, a bold frozen header and a sheet-scoped ColumnTitle01 name.
- Documents are generated by the build with pandoc; the build writes each `.htm` when its `.md` is newer.
