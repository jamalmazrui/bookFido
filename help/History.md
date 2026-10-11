---
title: "bookFido History"
author: "Jamal Mazrui"
---

# bookFido History


## 11 October 2026 -- the start-up fault, closed from both sides

- **Nothing in bookFido's start-up needs an embedded library now.** 1.1.33 still stopped with the same TypeInitializationException, and its log held no line from the resolver at all. The field that caused it, connectionDb, of SqliteConnection's type, lost its "= null": a static field is null without it, and without it bookFido's static set-up no longer touches SQLite, so it cannot fail there whatever the resolver does.
- **The log shows which build ran and what was asked.** A line, "Embedded libraries: the resolver is ready", marks a build with the resolver in its own class; and the resolver logs every request as it arrives, before anything that could fail, so a log tells a library never asked for from one asked for and not found.

## 11 October 2026 -- bookFido starts again

- **The cause, named by the new error handler.** Every run stopped at once with "TypeInitializationException: The type initializer for 'bookFido' threw an exception", over "Could not load file or assembly 'Microsoft.Data.Sqlite'". bookFido embeds its libraries in its program file, and a resolver loads them when .NET asks. But the resolver was part of bookFido and kept what it had loaded in one of bookFido's static fields. bookFido's own static set-up includes a field of SqliteConnection's type, so compiling that set-up asked for Microsoft.Data.Sqlite before any static field existed, the resolver's among them: the resolver locked on an empty field, which threw, its catch returned nothing, and .NET reported the library missing.
- **The fix.** The resolver is now a class of its own, EmbeddedAssemblies, with its own state, logging through Homer's Log; it depends on nothing of bookFido's, so it works while bookFido is still being set up. No other Homer app uses an assembly resolver, so none shares the fault.

## 11 October 2026 -- from an audit and a report

- **No run ends unseen.** Every run of 1.1.29 and 1.1.31 on 11 October stopped right after the log's header, with no error, no message and no "session end", so neither a crash nor a hang could be told apart. An unhandled-error handler is now set first in Main, for any thread, writing the full error to the session log and showing it in a message box; start-up writes a checkpoint at each step, so the last one names where a stopped run stopped; and each log line reaches the file before the console.
- **The last installer survives a failed build, and error 110 is retried.** Inno Setup removes its output before writing, so 1.1.31's failure, "EndUpdateResource failed (110)", left no installer at all. Inno now writes bookFido_setup.building.exe, which replaces bookFido_setup.exe only when complete; error 110 is tried again at most twice, ten seconds apart, and any other failure stops at once.
- **Launch after setup works on a first install.** The installer's Launch box writes a marker file into %LOCALAPPDATA%\\<App>\\logs, but on a first install that folder did not exist yet -- it was made later, when setup kept its log -- so the marker could not be written, and the program did not start after setup. The marker line now makes the folder first. Found by an audit of bookFido by another AI, and fixed in every Homer installer.

## 9 October 2026 -- from the build and release logs

- **The build compiles again.** Line 5332 of bookFido.cs, the statement that fills in a book's missing address, year, title and description, had the second half of its WHERE clause pasted onto its end a second time, after the string had closed; the compiler stopped with "empty character literal". The duplicate is removed and the statement is as intended. A scan of every C# file in every Homer project found no other such damage.

## 9 October 2026 -- spoken tutorials

**Ten spoken walks, in the Homer pattern of ten.** Walk 0 is an overview and table of contents; 1, the user interface, each control shown as the screen reader speaks it; 2, installing; 3 to 8, tasks around concrete wants -- your first catalog, reading it by title, Audible's companion PDFs, Goodreads shelves, NLS and Bookshare history, and the catalog in DbDo; and 9, the conclusion, with a glossary in two voices and every way to get help. Each task walk is predicted at three minutes or more by the kit's measured model, and every key and feature it teaches was checked against this guide. The build speaks them; the installer ships them and their audio; the repository carries both.

## 8 October 2026 -- an audit by another AI

ChatGPT audited bookFido and reported 36 findings. Checked against the code, these held and are fixed:

- **Filling one field no longer erases three others.** When any one of a book's address, first-published year, Wikipedia title or description was blank and a value arrived, all four incoming values were written, blanks included; a year arriving alone erased the other three. Each field is now filled only when it is blank. Reproduced and checked with the program's own SQL.
- **A checkpoint no longer empties saved collections.** The saved Kindle, Goodreads, Bookshare and NLS rows wait until their phase runs, and a checkpoint during the Audible phase wrote the still-empty lists over all four. Each is now saved from its saved rows until its phase has rows of its own.
- **No gap when the state is saved.** The old state file was deleted before the new one moved in, and every run used the same temporary name; the new state now replaces the old in one step, under a name of its own.
- **Audible's sign-in stays with Audible.** Its cookies went to any address downloaded, including a PDF link found in an error page; they now go only to Audible's own sites (audible.com, audible.co.uk and the other Audible stores).
- **Only a whole PDF is kept.** A download went straight to its final name, so a broken transfer, an empty reply or an error body became a file the next run trusted. It is now written to a .part file and kept only when it begins as every PDF does.
- **Authors given as one string are read as names.** A string of authors was walked a letter at a time, and every one-letter "name" was then thrown away, in both the Kindle and Bookshare readers.
- **Kit tools** updated from HomerDev 1.63.3.

Left for later, as larger changes: keys for matching books across libraries that cannot merge two distinct books; one instance at a time; and tabs of a reused browser that are not bookFido's.

## September 2026: built on the Homer Development Kit

The first version built by the kit's build (the number after 1.1.18).

### What's new

- **Setup.** The Results box at the end of setup is titled "bookFido Setup Results", and the finish page uses the Homer wording: the verb first, no "recommended", and "Launch bookFido (desktop hotkey ...)".
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
