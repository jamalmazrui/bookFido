---
title: "bookFido Guide"
author: "Jamal Mazrui"
---

# bookFido Guide

bookFido is the dog that fetches your books. This is the full guide; for a quick start, see ReadMe.

Give it a moment and it will go through every book you own or have borrowed, across five services, and hand you back a catalog you can actually read: a document, a spreadsheet, and a database, all built for a screen reader.

It was written by a blind developer for blind readers, and it is part of the Homer Tools series, a sibling of [urlFido](https://github.com/jamalmazrui/urlFido).

## What it does

bookFido opens Microsoft Edge, signs in as you, and visits the libraries you choose:

- **Audible** — every title in your library, plus each companion PDF downloaded under a friendly name and converted to accessible HTML.
- **Kindle** — every book on the same Amazon account.
- **Goodreads** — every shelf of your My Books list, with your ratings, dates, and reviews.
- **Bookshare** — your whole download history.
- **NLS** — your reading history from BARD, the National Library Service's talking-book and braille collection.

Then it fills in what the libraries leave out: publication years, publishers, descriptions, and author biographies from Open Library and Wikipedia.

## What you get

In your Downloads folder, after every run:

- **bookFido.htm and bookFido.md** — one combined catalog of every book from every library, in order by title, each with the publication year in parentheses and all the fields its own library offers. This is the document that opens when the run finishes.
- **A document for each library** — Audible_Library, Kindle_Library, Goodreads_Library, Bookshare_Library, and NLS_Library, each in both HTML and Pandoc-flavored Markdown, with a table of contents, an introduction with statistics, every book's details, and appendixes indexing the library by author, narrator, series, genre, publisher, subject, and rating, ending with About the Authors.
- **bookFido.xlsx** — one workbook with a sortable sheet for each library.

And in its data folder, `%LOCALAPPDATA%\bookFido\data`, **bookFido.db**, a SQLite database in the standard [DbDo](https://github.com/jamalmazrui/DbDo) schema, holding your books, authors, and publishers as related records you can browse and query in DbDo itself.

## Why it is pleasant to use

- **It talks.** Timed message boxes announce each page found, each book cataloged, and the estimated minutes remaining, each one focused so your screen reader speaks it in its own voice.
- **It remembers.** Everything gathered is kept in bookFido.db and a state snapshot, so a second run never looks up what it already knows. A first full catalog can take hours; the runs after it take minutes.
- **It is polite to the web.** One request per second per service, rate-limit responses honored, and a circuit breaker that steps back from a service that has closed its door for the day.
- **It respects your conventions.** Titles order by title, ignoring a leading A, An, or The; authors order by surname. Spreadsheets carry one region at cell A1, a bold header row, and a ColumnTitle01 name so JAWS announces column headers, with a truly empty cell wherever a value is unknown.
- **It knows a book from its twin.** A title you own in more than one library is recognized as one work, so its details are gathered once and shared, and an author's biography is fetched only once however many books they wrote.
- **It keeps your business private.** The log is a debugging aid only. In bookFido's own lines, paths under your user profile are replaced by a placeholder before anything is written; the log's opening lines, written by the shared Homer log, name the Windows version, user and computer, as every Homer program's log does.

## How bookFido reaches your libraries

bookFido reads the same web pages you would read yourself. It drives Microsoft Edge, signs in as you, and gathers what is on the pages each library shows to its own account holder — your library, your shelves, your download history. These are pages made available to you once you are logged in; nothing is collected anonymously, and nothing is taken that you could not open by hand a page at a time.

It reads politely. Requests to any one service are spaced about a second apart, with longer pauses for the reference sites, and when a service asks bookFido to slow down it waits for as long as it is asked. If a service declines repeatedly, bookFido stops asking it for the rest of the run. The point is to be an unremarkable visitor: your own reading, fetched at a human pace, without crowding anyone else off the server.

Nothing about your libraries leaves your computer. bookFido sends no book, author, list, or account information anywhere. It fetches pages from the libraries you chose, looks up publication details and author biographies from Open Library and Wikipedia by title and author name, and writes everything it learns to your own Downloads folder and to bookFido.db in its data folder on your computer. There is no account to create, no server behind it, and no telemetry of any kind.

## Installing

Download `bookFido_setup.exe` from the [bookFido releases page](https://github.com/jamalmazrui/bookFido/releases) and run it. It needs administrator rights, because it installs for everyone on the computer. The setup wizard asks for the folder the first time (default: `C:\Program Files\bookFido`); an update goes to the same folder without asking. It adds a Start-menu group and a desktop shortcut whose hotkey is **Alt+Control+Shift+B**.

The last page offers two checkboxes. **Launch bookFido now** is checked; **Open the user guide** is not. When you press Finish, a short results box says what was installed and where the logs are. bookFido starts after you close that box, so the box is never hidden behind it.

The program itself is in the `exec` folder of the installation, and this guide and the other documents are in the `help` folder. `ReadMe.htm` and `License.htm` are at the top.

## Updating

In the opening dialog, press **F11**. bookFido asks GitHub whether a newer version exists. If one does, Yes is the default: press Enter and the new setup program is downloaded and started. If you already have the newest version, No is the default. If the check fails, for example with no internet connection, a message says so. The dialog's Help (F1) ends with the same check.

## Using it

Press Alt+Control+Shift+B, or run bookFido from the Start menu. An opening dialog explains what will happen and offers a checkbox for each library — Audible, Bookshare, Goodreads, Kindle, and NLS — all checked to begin with. Uncheck any you would rather skip this time; their books stay in your catalog from the last run either way.

Choose OK, and if a library asks you to sign in, do so in the Edge window that opens and then answer the prompt. Each sign-in is usually needed only once, because bookFido keeps its own browser profile.

The rest is automatic. You can cancel at any point: progress is saved every couple of minutes, and the next run continues where this one stopped.

## Where bookFido keeps its files

- **Your catalogs**: in your Downloads folder, as listed under What you get.
- **What it has learned**: `bookFido.db` and the state snapshot `bookFido.json`, in `%LOCALAPPDATA%\bookFido\data`. Versions before the move to the Homer Development Kit kept them beside the program; the first run of this version moves them there, taking the newest copy when there is more than one.
- **Its browser profile**, which holds your sign-ins: `%LOCALAPPDATA%\bookFido\EdgeProfile`.
- **Logs**: one per run, in `%LOCALAPPDATA%\bookFido\logs`, named for the moment the run started; the 30 newest are kept. If something goes wrong, zip that folder and send it.

Uninstalling removes the program and its logs. It leaves the data folder and the browser profile, so hours of gathering and your sign-ins survive a reinstall; delete `%LOCALAPPDATA%\bookFido` by hand to remove them too.

## For developers

How bookFido is built, released and structured is in Developer, in this folder.

## License

MIT. The embedded libraries retain their own licenses, including PdfPig (Apache 2.0) and EPPlus 4 (LGPL).
