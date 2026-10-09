---
title: "bookFido ReadMe"
author: "Jamal Mazrui"
---

# bookFido ReadMe

bookFido is the dog that fetches your books. It goes through every book you own or have borrowed in Audible, Bookshare, Goodreads, Kindle and NLS, and hands you back one catalog you can read: a document, a spreadsheet, and a database, all built for a screen reader.

This is the quick start. The full guide is `help\bookFido.htm`.

## Install

1. Download `bookFido_setup.exe` from the [bookFido releases page](https://github.com/jamalmazrui/bookFido/releases).
2. Run it. It asks for administrator rights, because it installs for everyone on the computer.
3. On the last page, leave **Launch bookFido now** checked and press Finish. Read the results box, then close it; bookFido opens.

You need Windows 10 or 11, 64-bit. Edge is already part of Windows, and nothing else needs installing.

## Catalog your books

1. Press **Alt+Control+Shift+B** from anywhere in Windows. The opening dialog explains what will happen.
2. Leave checked the libraries you want: Audible, Bookshare, Goodreads, Kindle and NLS. All are checked to begin with.
3. Press Enter.
4. If a library asks you to sign in, do it in the Edge window that opens, then answer the prompt. Each sign-in is usually needed once.

bookFido talks you through the run. When it finishes, the combined catalog, `bookFido.htm`, opens from your Downloads folder, beside a document for each library and the `bookFido.xlsx` workbook.

## Keys in the opening dialog

- **Alt+A**, **Alt+B**, **Alt+G**, **Alt+K**, **Alt+N** move to the Audible, Bookshare, Goodreads, Kindle and NLS checkboxes.
- **Enter** starts; **Escape** closes without searching.
- **F1** shows Help.
- **F11** checks the web for a newer version of bookFido and offers to install it.

All the keys are listed in `help\Hotkeys.htm`.

## Learn by listening

Ten short spoken walks teach bookFido, each a few minutes long, in two voices: a host, and a screen reader saying what you would hear. They are installed in the `help\tutorials` folder as mp3 files, with a playlist, and as text in `Tutorials.htm`. Start with walk 0, the overview.

## When something goes wrong

Every run keeps a log in `%LOCALAPPDATA%\bookFido\logs`, one file per run. Zip that folder and send it with a description of what happened.

## License

bookFido is free and open source under the MIT License. See `License.htm`.
