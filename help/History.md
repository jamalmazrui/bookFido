---
title: "bookFido History"
author: "Jamal Mazrui"
---

# bookFido History


## 30 October 2026, later -- the build compiles the kit's Schema

- **bookFido builds again.** The kit's Schema, which bookFido uses to add a field BookTrail lacks, was named in the wrong list in build.cmd, so it was never compiled and the build stopped with thirteen errors. It is now one of the kit modules the build compiles (found by ChatGPT's audit of the logs).

## 30 October 2026 -- BookTrail kept as DbDo makes it

- **Catalog documents are written safely.** Each library's .htm and .md catalog was written straight over the old one, so a full disk or an interruption could leave neither a good old one nor a good new one. Each is now written beside its target and swapped in whole, through the Homer kit (from an audit by another AI).
- **A book is found in BookTrail by its fields, not by spelling its prime.** bookFido built the prime itself, author|title|source, but DbDo now joins look and prime with a hidden separator when it opens a database, and BookTrail's prime now includes series and volume. bookFido would have missed every book and tried to add it again. It now finds a copy by its source and catalog number, or by author, title and source (from an audit by another AI).
- **BookTrail is never rebuilt.** An older BookTrail missing a field bookFido writes gains it, empty, through the Homer kit's Schema, the same way DbDo upgrades a database. bookFido no longer rewrites look and prime to its own form, which would have undone DbDo's.
- **Two editions are kept apart.** A second copy with a different catalog number but the same author, title and source is no longer written over the first. It is left out, the log names it, and the run's summary counts it.
- **The build needs HomerDev 1.94.0 or later.**

## 25 October 2026 -- the newest kit, and extCheck's rules

- **The build uses the newest Homer kit it finds,** and logs each one: your HomerDev variable and C:\HomerDev. The variable named an older installed kit in Program Files (1.77.6), which a build took first.
- **Files it makes are judged by extCheck's rules,** from the kit's Quality.cs, now compiled in (HomerDev 1.92.5). A route table the build puts in exec is no longer removed by tidy, so conversions use the kit's current ways.

## 20 October 2026 -- Companion PDFs through the Homer kit

- **Recordings and speech, as in every Homer app:** a recording converts to a transcript, and a document to spoken audio, through the Homer kit. The build needs HomerDev 1.87.5 or later.
- **Web pages in the Homer page form,** as every Homer app makes them: one heading for the title, the page's parts marked for a screen reader, and pictures carried inside the page. The build needs HomerDev 1.87.1 or later.
- **Audible's companion PDFs become web pages through the Homer kit's one conversion,** like every Homer app's. bookFido's own PDF readers had become the kit's PdfReader, which reads a tagged PDF's structure and an untagged one's layout. Pages are made through Pandoc when it is installed and through the kit's own writer when it is not, so nothing more is needed on your computer.
- **Each existing companion page is made again once,** the first time bookFido sees it, so all are made the same way.
- **PdfPig is version 0.1.16,** the one the kit's reader is built and tested on.
- **The build needs HomerDev 1.81.0 or later.**

## 14 October 2026 -- a PDF is read, not only recognized

- **A companion PDF must open in PdfPig and have a page.** A short file holding only a PDF's first and last markers passed as complete. bookFido now opens it with the PDF parser it already embeds; an encrypted PDF is kept, logged as one whose contents could not be checked (from a second audit by another AI).
- **An unrecorded PDF named for the title is no longer adopted** for a book, since nothing in it shows which book it belongs to; the next free name is used, and the file is recorded from then on.

## 14 October 2026 -- the developer guide's kit path

- **The developer guide names the kit's classes where they are,** in C:\HomerDev\exec\CSharp; it still named C:\HomerDev\CSharp, a folder the kit no longer has (from an audit by another AI).

## 14 October 2026 -- progress you hear without losing your place

- **Progress no longer pulls you out of another program.** Each announcement was a message box that fought for focus, since a screen reader speaks only a window that is truly active, and a run takes hours, so every box took the cursor from whatever you were doing. Now the box is shown only while bookFido or the Edge it drives is in front. When another program is in front, the same words are spoken through the shared Homer speech class, which reaches JAWS, NVDA or Narrator without touching focus, and the log says which way each announcement went.
## 14 October 2026 -- contents lists, and plainer guides

- **Every help page has a contents list.** The build ran Pandoc without --toc, so a long guide had no list of its sections. Pages are now built with one, and a page that lacks one is built again once to gain it.
- **The user guide reads at grade 8.3, down from 9.6,** with every fact kept; it now also names bookFido-companions.txt, the record of which book each companion PDF belongs to.

## 14 October 2026 -- the announcement in plainer words

- **Announce.md now reads at grade 8.2, down from 9.3,** inside the ninth-grade target for documents shared publicly. Long sentences were split and dense phrases put plainly; every fact, key and link is kept, and first-person passages keep their author's own words. The section about the companion tools reads the same in each announcement that carries it, with the tools' hotkeys in alphabetical order.

## 14 October 2026 -- the newest Pandoc for the help pages

- **The help pages are made by the newest Pandoc, not the first on the PATH.** The build took the first pandoc on the PATH, which was Pandoc 2.19 in C:\bin, while Pandoc 3.12.1 was installed in Program Files. It now asks the kit's newest.ps1, as the kit's app template has done since kit 1.56.0, and the log names every copy found and its version.

## 14 October 2026 -- each companion PDF known by its book

- **A file named for the title is no longer taken on trust.** Any "Title.pdf" already in the folder was treated as the book's companion -- another edition with the same title, a partial file, a file put there by hand -- and a file under the server's name was renamed into it on its name alone. Each download folder now keeps one record, bookFido-companions.txt, of which book each companion belongs to (its asin), where it came from, its size and fingerprint. Another edition with the same title gets "Title (2).pdf". A complete PDF from before the record is adopted and recorded, so your library is not downloaded again; an incomplete one is left as it is, never deleted, and the book's file takes the next free name. A file under the server's name is claimed only when the record says it is this book's, or its name carries this book's asin (from an audit by another AI).
- **A PDF must be complete, not only begin like one.** It must also end with PDF's end marker and match the length the server announced, so a transfer cut short is refused rather than kept.
- **A web page sent instead of a PDF is read only up to the 500,000 characters kept,** rather than read in full and then cut.

## 14 October 2026 -- started as you

- **Setup starts bookFido as you,** not with the installer's administrator rights. Its comment said starting through cmd did that; it did not (from an audit by another AI).

## 12 October 2026 -- Camel Type constants, and tools before the decision

- **bookFido's constants are Camel Type's**, without the obsolete c_ prefix.
- **The build refreshes the kit's tools before it decides whether anything changed,** so a kit update to them reaches the installer in the same build.

## 12 October 2026 -- one switch to install with every default

- **bookFido_setup.exe /DEFAULTS** installs as pressing Enter at every step would, then launches bookFido; administrator approval is still asked. A /VERYSILENT install no longer stops at a Results box nobody can see.

## 12 October 2026 -- BookTrail holds only your own books

- **The sample books go once yours arrive.** DbDo's BookTrail template comes with six sample books, so a new reader sees what the database holds, and bookFido fills that same file in %LOCALAPPDATA%\\DbDo\\data\\BookTrail. A library of 1,217 gathered books still carried A Study in Scarlet and five more nobody had read. After a run writes your books, each sample is removed -- but only one that has never been edited and that nothing links to, so a sample you made your own stays. The samples are read from DbDo's installed template, so new samples in DbDo need no change to bookFido. The run's summary says how many went.
- **A build with nothing to change builds nothing.** When nothing bookFido ships has changed since its last build, the build keeps exec\\bookFido.exe, its version number and bookFido_setup.exe as they are, rather than taking a new number and building both again; release then finds that version already published. Any change, even to a help page, still makes the full build with a new number. `build fresh` builds regardless.
- **The build finds PdfStructure.cs again.** It is compiled with bookFido.cs, but the repository's list of files did not name it, so tidy took it for a stray, moved it into notes\\other, and the build stopped: "Source file PdfStructure.cs could not be found". The list names it now, the build brings back a copy tidy moved, and the kit's tidy keeps any file the build compiles.
- **A BookTrail bookFido makes comes complete.** When DbDo has not yet made BookTrail, bookFido makes it from DbDo's template, and now brings every file the template has -- its settings, the reading list, the reading-by-year query and the reading log report -- not only its settings. A file already in the folder is kept.

## 11 October 2026 -- your books in DbDo's BookTrail, and better companion pages

- **BookTrail instead of bookFido.db.** When DbDo is installed for you (%LOCALAPPDATA%\\DbDo exists), bookFido writes your Audible, Kindle, Goodreads, Bookshare and BARD libraries into DbDo's BookTrail; without DbDo, no database is written. If DbDo has never opened BookTrail, it is made from DbDo's installed template. Each copy of each book is one row, known by author, title and source, so an Audible copy and a BARD copy keep their own narrator, format and catalog number. New fields hold the publisher, first publication, language, length, pages, ISBN, progress, average rating and its count, the Wikipedia article, the companion files, and the work every copy shares. bookFido refreshes what the source says, fills only blanks in what you may have written, moves a status only forward, and only adds tags and notes. Authors join BookTrail's author list with their biographies. An older BookTrail is brought up to date in one transaction, keeping every row, any column you added, and every link. BookTrail is also where a later run finds what was already gathered.
- **Companion .htm files read as their PDF is laid out.** Pages are read in reading order, so two columns read down each column; a short bold heading in body-sized type is a heading; and a tagged PDF's own headings, lists, tables and figure descriptions are used when its tags reach most of the text (PdfStructure.cs, after HomerScribe's way of reading a PDF). Any failure falls back to the reading before. Each new .htm carries a mark, and one bookFido made earlier is made again once, so the files already there improve; a file anyone else made is left alone.
- **The first build stopped once**, on a PdfPig member HomerScribe uses but version 0.1.14 does not make public (Structure.TokenScanner); PdfStructure.cs now follows references with the public Structure.GetObject. Everything else compiled the first time.
- **The installer was right.** It installs to C:\\Program Files\\bookFido in administrative mode, as its logs show; no UAC prompt appears when the setup program is started from a window that already runs as administrator.

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
