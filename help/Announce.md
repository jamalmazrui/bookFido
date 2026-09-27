---
title: "bookFido Announcement"
author: "Jamal Mazrui"
---

# bookFido, built on the Homer Development Kit

bookFido catalogs your book libraries -- Audible, Bookshare, Goodreads, Kindle and NLS -- into one accessible catalog, a workbook and a database. Get it from the [bookFido releases page](https://github.com/jamalmazrui/bookFido/releases).

## What is new

- **Update from inside the program.** Press F11 in the opening dialog to check GitHub for a newer version. If there is one, Enter downloads and starts its installer.
- **A log of every run.** Each run keeps its own log in `%LOCALAPPDATA%\bookFido\logs`, so a cancelled run no longer empties the log of the run before.
- **Your gathered details in one place.** bookFido.db and bookFido.json now live in `%LOCALAPPDATA%\bookFido\data`, wherever the program runs from. The first run moves them there.
- **A tidier installation.** The program is in the `exec` folder and the documents in `help`, as in every Homer Tools program. The installer's results box comes before bookFido starts, and uninstalling keeps what bookFido has learned.

The full list of changes is in History.

## The companion tools

bookFido is part of the Homer Tools series, beside [urlFido](https://github.com/jamalmazrui/urlFido), which downloads files from web pages by extension, and [urlCheck](https://github.com/JamalMazrui/urlCheck), which checks web pages for accessibility.
