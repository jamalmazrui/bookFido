---
title: "bookFido History"
author: "Jamal Mazrui"
---

# bookFido History

## September 2026: built on the Homer Development Kit

The first version built by the kit's build (the number after 1.1.18).

### What's new

- **Built with HomerDev 1.43.19.** The build refreshes the kit's tools under their current names, and the ones that call each other now find each other; `scripts\tidy`, `scripts\check` and `scripts\release` carry the day's fixes, among them a release that publishes a draft and confirms it is GitHub's latest.
- **Built on the Homer Development Kit.** bookFido compiles the kit's shared classes (Elevate, Inix, Lbc, Log, Paths, Say, Util, Web) from `C:\HomerDev\CSharp` instead of carrying copies of Lbc and Say. It follows the Homer layout: the program in `exec`, the documents in `help`, one log per run in `logs`.
- **F11 checks for a newer version.** In the opening dialog, F11 asks GitHub for the latest release. Yes is the default when a newer version exists and No when this one is current; Yes downloads `bookFido_setup.exe` and starts it. The dialog's Help (F1) ends with the same check.
- **One log per run.** Each run writes `%LOCALAPPDATA%\bookFido\logs\bookFido-yyyyMMdd-HHmmss.log`, with the environment in its header, and the 30 newest are kept. A cancelled run no longer empties the log of the run before, because it never shared a file with it.
- **The data folder.** bookFido.db and bookFido.json live in `%LOCALAPPDATA%\bookFido\data` wherever the program runs from. Before, the state file sat beside the program when that folder was writable and in `%LOCALAPPDATA%\bookFido` otherwise, while the database always sat beside the program -- which under Program Files could not be written, so an installed copy ran without its database. The first run moves both, taking the newest copy.
- **Documents.** The ReadMe is a quick start; the full guide is `help\bookFido.md` and `.htm`. New Developer, Hotkeys and this History; Announce describes the current release.

### Installer

- The program installs to `exec`, the documents other than ReadMe and License to `help`.
- The finish page offers Launch (checked) and Open the user guide (unchecked). A results box says what was installed and where the logs are, and bookFido starts only after that box is closed.
- A reinstall no longer asks for the folder; it goes where the last one went.
- The uninstaller removes the logs and the SQLite engine it wrote out, and keeps the data folder and the Edge profile with your sign-ins.

### For developers

- `buildbookFido.cmd` is the kit's C# build template with two sections of bookFido's own: the pinned NuGet libraries, fetched into `work\nuget` and embedded, and the carry-over from the layout before.
- The compiler is Roslyn, found with vswhere or installed as the Build Tools. The Framework compiler fallback is gone, because the kit's classes need more than C# 5.
- `RepoFiles.txt` and `LocalFiles.txt` decide what git carries; built programs are no longer in the repository.
- The compile is clean: eleven long-standing warnings about variables declared and never used, and one field set and never read, are gone, so a new warning stands out.

## 1.1.x (July 2026)

- **Five libraries.** Audible, then Kindle, Goodreads, Bookshare and, last, NLS (the National Library Service's BARD reading history), each with its own document, a sheet in one bookFido.xlsx workbook, and entries in the combined bookFido.htm catalog that opens at the end.
- **bookFido.db**, in the standard DbDo schema, so nothing already gathered is looked up again; a full five-library run with a warm database takes about five minutes.
- **Twins.** A book held in more than one library is recognized as one work, its details gathered once and shared; an author's biography is fetched once.
- **The opening dialog**, built with Lbc: an introduction and a checkbox for each library, all checked.
- **Politeness and resilience.** Paced requests per service, rate-limit responses honored, a breaker that stops asking a service that declines, state saved every two minutes so a cancelled run resumes.
- **Bookshare's early stop**, NLS sign-in and paging fixes, a faster startup probe, and publisher names stored without their source label.

## 1.0.0 (21 July 2026)

The first release under the name bookFido, formerly GetAudibleInfo: the Audible library walk, companion PDFs downloaded under friendly names and converted to accessible HTML, and an accessible catalog enriched from Audible's catalog service, Open Library and Wikipedia.
