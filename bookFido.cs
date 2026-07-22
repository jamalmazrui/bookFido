// bookFido.cs
// Harvests the signed-in user's Audible library: downloads the companion PDF
// files that publishers attach to titles, and writes an accessible catalog of
// the whole library as Audible_Library.htm and Audible_Library.md in the
// Downloads folder, opened in the default web browser at the end.  Each title
// is an h2 heading linking to its Audible detail page, followed by the fields
// the library page offers under Audible's own labels (By, Narrated by,
// Series, Summary, and listening progress), plus the companion PDF outcome.
// Each title is enriched with details from public web services in the
// manner of the biblio tool: Audible's own catalog service supplies the
// rating average and count, length, release date, publisher, and language;
// and Open Library supplies the year of first publication and, when Audible
// offers no publisher, the publisher of the print editions.  The gathering is pipelined rather than
// sequential: three background lanes, one thread per service, start as soon
// as login succeeds and consume titles while the library pages are still
// being walked.  Each lane keeps the same polite one-request-per-second
// pacing toward its own service that the sequential design had, so no
// service sees a heavier load; the speedup comes purely from overlap.
// Individual request failures are logged and leave fields absent.  Appendixes index the titles by author, narrator,
// series, listening progress, publisher, and rating, each in
// alphabetical order with internal links to the title headings.  Companion
// files are not mentioned in the catalog; they sit beside it in Downloads.  The program drives Microsoft Edge through
// the Chrome DevTools Protocol over a raw WebSocket, the same single-file
// technique used by urlFido, so it builds to one portable 64-bit exe with the
// in-box .NET Framework 4.8 compiler.  Build with buildbookFido.cmd,
// which produces a GUI (winexe) program with no console window.
//
// Flow: an introductory message box explains the program and offers OK or
// Cancel.  The program then prefers the user's main Edge profile, so an
// existing Audible login is reused with no sign-in effort; recent Edge
// versions refuse remote debugging against the main profile, and Edge cannot
// be attached at all while it is already running, so in those cases a
// persistent dedicated profile is used instead, which remembers the Audible
// login after the first run.  The library at
// https://www.audible.com/library/titles is then opened, and only if the user
// is not already logged in does a message box pause for login, with Cancel
// offered as a last chance to bail out.  On each library page, every anchor
// whose href contains /companion-file/ is collected and downloaded to the
// standard Windows Downloads folder by replaying the browser cookies over
// direct HTTP.  The next-page link is followed until the next button carries
// the bc-button-disabled class, which marks the last page.  Guards against
// infinite loops: a visited-url set and a maximum page count.  Every step is
// logged to bookFido.log beside the exe.  Each download is announced
// with a timed message box whose title is the base file name, which JAWS
// speaks automatically.  Because the program launched Edge itself, it closes
// Edge again when it finishes.
//
// Hardening learned from urlFido and from field runs:
//   - Edge signs a brand-new profile into the Windows account by default and
//     immediately offers to sync, producing windows the user has to dismiss.
//     The launch switches and the seeded Preferences file below shut down
//     sign-in, sync, extensions, and the background services that drive them.
//     This does not affect logging in to the audible.com website itself.
//   - The CDP WebSocket can be forcibly closed, for example when a stray Edge
//     window is closed by the user.  Every CDP command therefore retries once
//     after reconnecting to a fresh page target.
//   - PDF remains a hard format for screen reader users, so every companion
//     PDF also gets a same-root-name .htm sibling that preserves headings,
//     paragraphs, and lists, built with the embedded UglyToad.PdfPig library
//     (Apache-2.0, pure managed code) rather than Microsoft Office COM, so
//     there is no Office dependency and the exe stays a single file.  Font
//     size tiers become heading levels, bullet and numbered lines become
//     lists, hyphenated line breaks are rejoined, images are dropped, and
//     any /Alt image descriptions found in the PDF are listed at the end.
//     The PdfPig assemblies are embedded as manifest resources at build time
//     and loaded through an AssemblyResolve handler, the same single-file
//     technique 2htm uses for Markdig.  Guided by the Matterhorn Protocol
//     (the PDF/UA conformance testing model), the generated HTML also:
//     takes its title element and h1 from the PDF's own Title metadata when
//     one is present (checkpoint 06, dc:title); declares the language from
//     the PDF's /Lang entry when one is declared (checkpoint 11); never
//     skips heading levels in sequence (checkpoint 14); wraps the content
//     in a main landmark; and treats repeating page furniture as pagination
//     artifacts (checkpoint 18), dropping bare page numbers and short lines
//     that recur across three or more pages.
//   - Amazon names companion PDFs with warehouse codes like
//     bk_tcco_001380.pdf.  Following the approach of the renTitle utility,
//     each download is given a human-friendly name instead: first choice is
//     the book title shown beside the View PDF button on the library page,
//     second choice is the Title metadata inside the PDF itself (parsed
//     natively from the Info dictionary or XMP, filtered through renTitle's
//     junk-title blacklist), and only as a last resort the server's code
//     name.  renTitle's character cleanup and -001 numbering for duplicates
//     are ported too, and a file downloaded under a code name by an earlier
//     run is renamed in place rather than downloaded again.
//   - Some older titles show a View PDF button whose companion-file url
//     answers with a web page instead of a redirect to the PDF.  For those,
//     the failure page is searched for a direct PDF link, and the title's
//     product page is consulted for a newer asin under Amazon's current
//     naming, which is then tried against companion-file as well.  The
//     failure page body is also saved beside the log for diagnosis.

using Homer;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.WebSockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.DocumentLayoutAnalysis.WordExtractor;

class bookFido
{
    // Constant definitions
    const int iDebugPort = 9222, iDelayApiMs = 800, iDelayDownloadMs = 1500, iDelayWikipediaMs = 1800, iDelayPageMs = 3000, iHttpTimeoutMs = 120000, iJitterMaxMs = 1500, iLaunchWaitMainMs = 15000, iLaunchWaitMs = 30000, iLoginTryMax = 3, iDrainReportMs = 30000, iHtmStatusCreated = 1, iHtmStatusExisted = 0, iHtmStatusFailed = 2, iLanePollMs = 250, iRateLimitStrikesMax = 3, iMessageBoxMs = 2000, iNavigateTimeoutMs = 60000, iSaveDocumentsMs = 360000, iSaveStateMs = 120000, iStallTicks = 3, iPageMax = 500, iSettleMs = 2500, iStatusDownloaded = 0, iStatusFailed = 3, iStatusFailedHtml = 2, iStatusSkipped = 1;
    const string sApiUserAgent = "bookFido/1.0 (https://github.com/JamalMazrui/bookFido; personal library catalog)", sAudibleApiUrl = "https://api.audible.com/1.0/catalog/products/", sEdgePathPrimary = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", sEdgePathSecondary = "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe", sLibraryUrl = "https://www.audible.com/library/titles", sOpenLibraryUrl = "https://openlibrary.org/search.json", sOpenLibraryAuthorSearchUrl = "https://openlibrary.org/search/authors.json?q=", sOpenLibraryAuthorUrl = "https://openlibrary.org/authors/", sWikipediaApiUrl = "https://en.wikipedia.org/w/api.php?action=query&list=search&format=json&srlimit=1&srsearch=", sWikipediaPageUrl = "https://en.wikipedia.org/wiki/", sWikipediaSummaryUrl = "https://en.wikipedia.org/api/rest_v1/page/summary/", sVersionText = BuildVersion.Version, sStartUrl = "https://www.audible.com/", sUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0";
    const uint iMbOk = 0x00000000, iMbSetForeground = 0x00010000, iMbTopmost = 0x00040000;

    // Static variable definitions
    static bool bSummaryWritten = false;
    static ClientWebSocket wsCdp = null;
    static HashSet<string> setSeenUrls = new HashSet<string>(), setVisitedPages = new HashSet<string>();
    static int iNextId = 1;
    static Dictionary<string, string> dFailedTitles = new Dictionary<string, string>(), dPdfNames = new Dictionary<string, string>();
    static HashSet<string> setCatalogAsins = new HashSet<string>();
    const string sKindleLibraryUrl = "https://read.amazon.com/kindle-library", sKindleSearchUrl = "https://read.amazon.com/kindle-library/search?query=&libraryType=BOOKS&sortType=acquisition_desc&querySize=50";
    static bool bKindleHarvested = false;
    static object[] aSavedKindleRows = null;
    static List<Dictionary<string, object>> lCatalog = new List<Dictionary<string, object>>();
    static List<Dictionary<string, object>> lKindleCatalog = new List<Dictionary<string, object>>();
    static List<Dictionary<string, object>> lGoodreadsCatalog = new List<Dictionary<string, object>>();
    static List<Dictionary<string, object>> lBookshareCatalog = new List<Dictionary<string, object>>();
    static object[] aSavedGoodreadsRows = null;
    static object[] aSavedBookshareRows = null;
    static Dictionary<string, string> dAuthorOlBio = new Dictionary<string, string>(), dAuthorWikiBio = new Dictionary<string, string>(), dAuthorWikiUrl = new Dictionary<string, string>();
    static HashSet<string> setAuthorsQueued = new HashSet<string>();
    static int iAudibleDone = 0, iAuthorsDone = 0, iHtmCount = 0, iOpenLibraryDone = 0, iWikipediaDone = 0, iWorkTotal = 0;
    static object oAuthorLock = new object();
    [ThreadStatic] static bool bLastFetchRateLimited;
    static bool bSearchAudible = true, bSearchBookshare = true, bSearchGoodreads = true, bSearchKindle = true;
    static bool[] aLaneSkipping = new bool[3];
    static volatile string sProgressText = "";
    static string sConsolidatedHtmPath = "";
    static volatile bool bStopLanes = false;
    static DateTime dtLastDocumentsSave = DateTime.MinValue, dtLastStateSave = DateTime.MinValue, dtRunStart = DateTime.Now;
    static volatile bool bWalkComplete = false;
    static bool bSavedWalkComplete = false;
    static string sDownloadDirShared = "";
    static Dictionary<string, string[]> dCompanions = new Dictionary<string, string[]>();
    static HashSet<string> setAuthorsCheckedOl = new HashSet<string>(), setAuthorsCheckedWiki = new HashSet<string>(), setCompanionsHandled = new HashSet<string>(), setDeadCompanions = new HashSet<string>(), setSavedAsinsCache = null;
    static int iSavedTotalPages = 0;
    static List<string> lSavedPage1Asins = null;
    static object[] aSavedRows = null;
    static string sSavedAtText = "";
    static int iTotalPagesSeen = 0;
    static List<string> lPage1AsinsSeen = new List<string>();
    static object oLogLock = new object();
    static Queue<Dictionary<string, object>> queueAudible = new Queue<Dictionary<string, object>>(), queueOpenLibrary = new Queue<Dictionary<string, object>>(), queueWikipedia = new Queue<Dictionary<string, object>>();
    static Queue<string> queueAuthorOpenLibrary = new Queue<string>(), queueAuthorWikipedia = new Queue<string>();
    static string sLastSavedName = "", sLogFilePath = "", sUserProfileDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    static string sDataDir = "";
    static Thread threadAudible = null, threadOpenLibrary = null, threadWikipedia = null;
    static volatile bool bHarvestDone = false;
    static JavaScriptSerializer jsonCodec = new JavaScriptSerializer() { MaxJsonLength = int.MaxValue };
    static List<string> lDownloaded = new List<string>(), lFailed = new List<string>(), lSkipped = new List<string>();
    static Process oEdgeProcess = null;
    static Random oRandom = new Random();
    static UTF8Encoding utf8NoBom = new UTF8Encoding(false);
    static string sLastFailureBody = "";

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid oFolderId, uint iFlags, IntPtr hToken, out IntPtr hPath);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern int MessageBoxTimeoutW(IntPtr hWnd, string sText, string sCaption, uint iType, ushort iLanguageId, uint iMilliseconds);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern IntPtr FindWindowW(string sClassName, string sWindowName);

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern bool SetForegroundWindow(IntPtr hWindow);

    [DllImport("user32.dll")]
    static extern bool BringWindowToTop(IntPtr hWindow);

    [DllImport("user32.dll")]
    static extern uint GetWindowThreadProcessId(IntPtr hWindow, IntPtr hProcessId);

    [DllImport("user32.dll")]
    static extern bool AttachThreadInput(uint iAttachThread, uint iAttachToThread, bool bAttach);

    [DllImport("kernel32.dll")]
    static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    static extern bool SetFocus(IntPtr hWindow);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    static extern bool PeekMessageW(out MSG oMessage, IntPtr hWindow, uint iFilterMin, uint iFilterMax, uint iRemove);

    [StructLayout(LayoutKind.Sequential)]
    struct MSG
    {
        public IntPtr hWindow;
        public uint iMessage;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint iTime;
        public int iPointX;
        public int iPointY;
    }

    [STAThread]
    static int Main(string[] aArgs)
    {
        int iExitCode;

        AppDomain.CurrentDomain.AssemblyResolve += resolveEmbeddedAssembly;
        iExitCode = 0;
        try { iExitCode = mainAsync().GetAwaiter().GetResult(); }
        catch (Exception oException)
        {
            log("Fatal error: " + oException.ToString());
            focusWhenShown("bookFido error");
            MessageBox.Show("bookFido stopped with a fatal error.  See bookFido.log for details.", "bookFido error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            iExitCode = 1;
        }
        shutdownEdge();
        writeSummarySections();
        return iExitCode;
    }

    static async Task<int> mainAsync()
    {
        bool bReachable, bSignedIn;
        int iFound, iNewTitles, iPage, iTry;
        string sAsin, sBookTitle, sCatalogJson, sCookieHeader, sCurrentUrl, sDownloadDir, sIntro, sLibraryHtmPath, sMainProfileDir, sNextHref, sOldProfileDir, sPdfUrl, sProfileDir, sScanJson, sSummary;
        Dictionary<string, object> dItem, dRow, dScan;
        int iLastDecile, iPercent, iTotalPages;
        List<string> lPageAsins;
        DialogResult dialogResultAnswer;
        List<string> lRetry;

        sLogFilePath = Path.Combine(dataDir(), "bookFido.log");
        File.WriteAllText(sLogFilePath, "", new UTF8Encoding(true));
        dtRunStart = DateTime.Now;
        log("bookFido started, version " + sVersionText);
        log("Embedded assemblies in this exe: " + string.Join(", ", Assembly.GetExecutingAssembly().GetManifestResourceNames()));
        ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol | SecurityProtocolType.Tls12;
        jsonCodec.MaxJsonLength = 50000000;
        sIntro = "This is bookFido version " + sVersionText + ".  bookFido visits every page of your Audible library and gathers what it finds there.  " +
            "It downloads the companion PDF files that publishers attach to audiobooks, naming each by its book title, creates a screen-reader-friendly .htm version of every PDF, and it builds an accessible catalog of your whole library, " +
            "with a heading for every title, its details enriched from Audible's catalog service, Open Library, and Wikipedia (this gathering step takes a few minutes for a large library), and appendixes indexed by author, narrator, series, publisher, rating, and more, saved as Audible_Library.htm and Audible_Library.md in your Downloads folder, with a sortable spreadsheet for every library as its own sheet of one bookFido.xlsx workbook.  Kindle books on the same Amazon account, the Goodreads My Books shelves, and the Bookshare My History list are cataloged alongside, as Kindle_Library, Goodreads_Library, and Bookshare_Library files, and a title present in more than one library shares its gathered details without extra requests.  " +
            "A note on announcements: this program works to keep each announcement window focused so your screen reader speaks it, but Windows can occasionally withhold focus from a background program; the complete play-by-play is always in bookFido.log.  " +
            "It opens Microsoft Edge at audible.com and uses your existing Audible login when possible.  " +
            "Progress is spoken through brief message boxes, and a full record is written to bookFido.log beside the program.  " +
            "When it finishes, it reports totals, opens the catalog in your web browser, and closes the Edge window it opened.  " +
            "Check the libraries to search below; all are checked to begin with.  Then choose OK to begin, or Cancel to exit.";
        if (!showOpeningDialog(sIntro)) { log("The user chose Cancel at the introduction, so exiting"); return 0; }
        sDownloadDir = downloadsFolder();
        log("Download directory: " + sDownloadDir);
        sDownloadDirShared = sDownloadDir;
        sProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "bookFido", "EdgeProfile");
        sOldProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GetAudibleInfo", "EdgeProfile");
        if (!Directory.Exists(sOldProfileDir)) sOldProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GetAudiblePDFs", "EdgeProfile");
        if (!Directory.Exists(sProfileDir) && Directory.Exists(sOldProfileDir))
        {
            Directory.CreateDirectory(Path.GetDirectoryName(sProfileDir));
            Directory.Move(sOldProfileDir, sProfileDir);
            log("Migrated the saved Edge profile from its previous folder, so the Audible login carries over");
        }
        Directory.CreateDirectory(sProfileDir);
        log("Dedicated Edge profile directory: " + sProfileDir);
        sMainProfileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Edge", "User Data");
        bReachable = debugPortReachable();
        if (bReachable) log("Edge is already listening on the debugging port, so reusing that instance");
        if (!bReachable && isEdgeRunning())
        {
            log("Edge is currently running, and a running browser cannot be attached for remote debugging, so the main profile is skipped");
        }
        else if (!bReachable && Directory.Exists(sMainProfileDir))
        {
            log("Trying the main Edge profile first, so an existing Audible login can be reused: " + sMainProfileDir);
            launchEdge(sMainProfileDir, sDownloadDir, true);
            bReachable = waitForDebugPort(iLaunchWaitMainMs);
            if (!bReachable)
            {
                log("Edge did not open its debugging channel against the main profile; recent Edge versions refuse this for security, so falling back to the dedicated profile");
                shutdownEdge();
            }
        }
        if (!bReachable)
        {
            launchEdge(sProfileDir, sDownloadDir, false);
            bReachable = waitForDebugPort(iLaunchWaitMs);
        }
        if (!bReachable)
        {
            log("The Edge debugging port never became reachable");
            focusWhenShown("bookFido error");
            MessageBox.Show("Microsoft Edge did not start with its debugging port.  See bookFido.log for details.", "bookFido error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
        await connectCdp();
        await enableDomains();
        bSignedIn = !bSearchAudible;
        for (iTry = 1; bSearchAudible && iTry <= iLoginTryMax; iTry++)
        {
            await navigate(sLibraryUrl);
            Thread.Sleep(iSettleMs);
            sCurrentUrl = await evaluate("location.href");
            log("Current url after opening the library: " + sCurrentUrl);
            if (sCurrentUrl.Contains("/library")) { bSignedIn = true; break; }
            if (iTry == 1) log("Not logged in yet, so pausing for the user to log in");
            focusWhenShown("bookFido: log in to Audible");
            dialogResultAnswer = MessageBox.Show("You are not logged in to Audible yet.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to exit without downloading anything.", "bookFido: log in to Audible", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (dialogResultAnswer == DialogResult.Cancel) { log("The user chose Cancel at the login prompt, so exiting"); await closeEdgeAsync(); return 0; }
        }
        if (!bSignedIn)
        {
            log("The library page could not be reached after " + iLoginTryMax + " attempts");
            focusWhenShown("bookFido error");
            MessageBox.Show("The Audible library page could not be reached.  See bookFido.log for details.", "bookFido error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            await closeEdgeAsync();
            return 1;
        }
        log("Logged in; beginning the library walk");
        loadState();
        startEnrichmentLanes();
        iLastDecile = 0;
        iPage = 1;
        iPercent = 0;
        iTotalPages = 0;
        sCookieHeader = "";
        while (bSearchAudible && iPage <= iPageMax)
        {
            sCurrentUrl = await evaluate("location.href");
            if (setVisitedPages.Contains(sCurrentUrl)) { log("Loop guard: this page url was already visited, so stopping: " + sCurrentUrl); break; }
            setVisitedPages.Add(sCurrentUrl);
            log("Scanning library page " + iPage + ": " + sCurrentUrl);
        savePeriodically(false);
            sScanJson = await evaluate(scanScript());
            if (sScanJson == "") { log("The page scan returned no result, so stopping"); break; }
            log("Scan result: " + sScanJson);
            dScan = (Dictionary<string, object>) jsonCodec.DeserializeObject(sScanJson);
            if (iTotalPages == 0 && dScan.ContainsKey("totalPages") && Convert.ToInt32(dScan["totalPages"]) > 1)
            {
                iTotalPages = Convert.ToInt32(dScan["totalPages"]);
                log("The library pagination reports " + iTotalPages + " pages, so progress will be announced at each ten percent");
            }
            if (iTotalPages > 0)
            {
                // The walk's percentage gets no dialog of its own; it rides
                // in the body of whichever announcement appears next, so the
                // screen reader speaks it as a trailing detail.  It is
                // computed here, after the page scan, once the page total is
                // known, so even page one's announcements carry it.
                iPercent = iPage * 100 / iTotalPages;
                sProgressText = iPercent + "%";
                if (iPercent / 10 > iLastDecile && iPercent < 100)
                {
                    iLastDecile = iPercent / 10;
                    log("Library walk " + (iLastDecile * 10) + " percent complete");
                }
            }
            sCookieHeader = await cookieHeader();
            iFound = 0;
            foreach (object oItem in (IEnumerable) dScan["pdfs"])
            {
                dItem = (Dictionary<string, object>) oItem;
                sPdfUrl = Convert.ToString(dItem["href"]);
                sBookTitle = dItem.ContainsKey("title") ? Convert.ToString(dItem["title"]) : "";
                dCompanions[asinFromUrl(sPdfUrl)] = new string[] { sPdfUrl, sBookTitle };
                if (setSeenUrls.Contains(sPdfUrl)) continue;
                setSeenUrls.Add(sPdfUrl);
                setCompanionsHandled.Add(asinFromUrl(sPdfUrl));
                iFound = iFound + 1;
                Thread.Sleep(iDelayDownloadMs + oRandom.Next(0, iJitterMaxMs));
                downloadCompanion(sPdfUrl, sBookTitle, sCookieHeader, sDownloadDir);
            }
            log("New companion PDF links found on this page: " + iFound);
            sCatalogJson = await evaluate(catalogScript());
            lPageAsins = new List<string>();
            if (sCatalogJson != "")
            {
                iNewTitles = 0;
                foreach (object oRow in (object[]) jsonCodec.DeserializeObject(sCatalogJson))
                {
                    dRow = (Dictionary<string, object>) oRow;
                    sAsin = Convert.ToString(dRow["asin"]);
                    lPageAsins.Add(sAsin);
                    if (iPage == 1) lPage1AsinsSeen.Add(sAsin);
                    if (setCatalogAsins.Contains(sAsin)) continue;
                    setCatalogAsins.Add(sAsin);
                    lCatalog.Add(dRow);
                    enqueueForEnrichment(dRow);
                    iNewTitles = iNewTitles + 1;
                }
                log("Catalog rows harvested on this page: " + iNewTitles);
            }
            iTotalPagesSeen = iTotalPages > 0 ? iTotalPages : iPage;
            if (iPage == 1 && libraryUnchanged())
            {
                log("The library is unchanged since " + sSavedAtText + ", so the remaining pages are loaded from the saved catalog");
                showTimedMessageBox("Library unchanged since " + sSavedAtText + "; " + aSavedRows.Length + " titles loaded");
                mergeSavedRows();
                break;
            }
            if (iPage > 1 && aSavedRows != null && pageFullyKnown(lPageAsins))
            {
                log("Reached previously cataloged titles on page " + iPage + ", so the rest is loaded from the saved catalog");
                showTimedMessageBox("Reached known titles on page " + iPage + "; loading the rest from the saved catalog");
                mergeSavedRows();
                break;
            }
            sNextHref = Convert.ToString(dScan["nextHref"]);
            if (Convert.ToBoolean(dScan["nextDisabled"])) { log("The next button is disabled, so this is the last page"); break; }
            if (sNextHref == "") { log("No next link was found, so treating this as the last page"); break; }
            if (setVisitedPages.Contains(sNextHref)) { log("Loop guard: the next link points to a page already visited, so stopping"); break; }
            Thread.Sleep(iDelayPageMs + oRandom.Next(0, iJitterMaxMs));
            await navigate(sNextHref);
            Thread.Sleep(iSettleMs);
            iPage = iPage + 1;
        }
        if (iPage > iPageMax) log("Loop guard: the maximum page count of " + iPageMax + " was reached");
        if (!bSearchAudible)
        {
            // Audible was not selected, so its saved world stands as it is:
            // the catalog rows come from the state snapshot without any
            // downloads or lookups, and the saved walk record is preserved
            // so the state file keeps its resume information.
            iTotalPagesSeen = iSavedTotalPages;
            lPage1AsinsSeen = lSavedPage1Asins != null ? new List<string>(lSavedPage1Asins) : new List<string>();
            materializeSavedRows(aSavedRows, lCatalog);
            log("Audible was not selected this run, so its " + lCatalog.Count + " saved titles stand as they are");
        }
        if (lFailed.Count > 0)
        {
            log("Retrying " + lFailed.Count + " failed downloads once, in case the failures were transient");
            lRetry = new List<string>(lFailed);
            lFailed.Clear();
            sCookieHeader = await cookieHeader();
            foreach (string sRetryUrl in lRetry)
            {
                sBookTitle = dFailedTitles.ContainsKey(sRetryUrl) ? dFailedTitles[sRetryUrl] : "";
                Thread.Sleep(iDelayDownloadMs + oRandom.Next(0, iJitterMaxMs));
                downloadCompanion(sRetryUrl, sBookTitle, sCookieHeader, sDownloadDir);
            }
            foreach (string sDeadUrl in lFailed) setDeadCompanions.Add(asinFromUrl(sDeadUrl));
        }
        bWalkComplete = bSearchAudible ? true : bSavedWalkComplete;
        if (bSearchAudible) sweepCompanions(sCookieHeader, sDownloadDir);
        // The libraries are processed in alphabetical order; a library left
        // unchecked is filled from the saved state first, so its books can
        // still be matched by any library processed after it.
        if (!bSearchBookshare) materializeSavedRows(aSavedBookshareRows, lBookshareCatalog);
        if (!bSearchGoodreads) materializeSavedRows(aSavedGoodreadsRows, lGoodreadsCatalog);
        if (!bSearchKindle) materializeSavedRows(aSavedKindleRows, lKindleCatalog);
        if (bSearchBookshare) await harvestBookshareAsync();
        if (bSearchGoodreads) await harvestGoodreadsAsync();
        if (bSearchKindle) await harvestKindleAsync();
        waitForEnrichment();
        saveState();
        sLibraryHtmPath = lCatalog.Count > 0 ? buildLibraryFiles(sDownloadDir) : "";
        try { writeWorkbook(sDownloadDir); }
        catch (Exception oException) { log("The spreadsheet workbook could not be written: " + oException.Message); }
        resolveTwins();
        if (lKindleCatalog.Count > 0)
        {
            try { buildKindleFiles(sDownloadDir); }
            catch (Exception oException) { log("The Kindle catalog documents could not be written: " + oException.Message); }
        }
        if (lGoodreadsCatalog.Count > 0)
        {
            try { buildGoodreadsFiles(sDownloadDir); }
            catch (Exception oException) { log("The Goodreads catalog documents could not be written: " + oException.Message); }
        }
        if (lBookshareCatalog.Count > 0)
        {
            try { buildBookshareFiles(sDownloadDir); }
            catch (Exception oException) { log("The Bookshare catalog documents could not be written: " + oException.Message); }
        }
        try { sConsolidatedHtmPath = buildConsolidatedFiles(sDownloadDir); }
        catch (Exception oException) { log("The consolidated catalog could not be written: " + oException.Message); }
        if (sConsolidatedHtmPath != "") sLibraryHtmPath = sConsolidatedHtmPath;
        log("HTML versions of PDFs created this run: " + iHtmCount);
        sSummary = "Downloaded " + lDownloaded.Count + ", skipped " + lSkipped.Count + ", failed " + lFailed.Count + (setDeadCompanions.Count > lFailed.Count ? " (" + setDeadCompanions.Count + " companions are known to be permanently unavailable from Audible)" : "") + ", with " + iHtmCount + " HTML versions of PDFs created and " + lCatalog.Count + " titles cataloged.  " +
            (lFailed.Count > 0 ? "The failed companions are remembered as unavailable and will not be retried on future runs.  " : "") +
            (skippedLaneNames() != "" ? "Because of rate limiting, remaining lookups from " + skippedLaneNames() + " were deferred and will be gathered on a future run.  " : "") +
            (lKindleCatalog.Count > 0 ? "Kindle library: " + lKindleCatalog.Count + " books cataloged as Kindle_Library.htm and Kindle_Library.md.  " : "") +
            (lGoodreadsCatalog.Count > 0 ? "Goodreads library: " + lGoodreadsCatalog.Count + " books cataloged as Goodreads_Library.htm and Goodreads_Library.md.  " : "") +
            (lBookshareCatalog.Count > 0 ? "Bookshare history: " + lBookshareCatalog.Count + " books cataloged as Bookshare_Library.htm and Bookshare_Library.md.  " : "") +
            "The combined catalog of every library was saved as bookFido.htm and bookFido.md.  " +
            "The catalog was saved as Audible_Library.htm and Audible_Library.md in Downloads, every library's spreadsheet is a sheet of the bookFido.xlsx workbook there, and the combined bookFido catalog opens in your web browser when you choose OK.  See bookFido.log for details.";
        log(sSummary);
        writeSummarySections();
        // Edge is closed before the results box appears, so the box faces no
        // foreground contention from the browser, and the box itself gets the
        // same focus treatment as the timed announcements.
        await closeEdgeAsync();
        focusWhenShown("bookFido finished");
        MessageBox.Show(sSummary, "bookFido finished", MessageBoxButtons.OK, MessageBoxIcon.Information);
        if (sLibraryHtmPath != "") openInDefaultBrowser(sLibraryHtmPath);
        return 0;
    }

    // Writes a time-stamped line to the console and to bookFido.log.
    // Writes a time-stamped line to the console and appends it to
    // bookFido.log, opening and closing the file for each line, so
    // the log on disk is always complete even if the program is canceled
    // mid-run, and its size is always current in File Explorer.
    // The log exists only as a debugging aid, so paths under the user's
    // profile are scrubbed to a placeholder, keeping the account name out
    // of a file the user may share; book information stays, since books
    // are the program's whole subject.
    static void log(string sMessage)
    {
        string sLine;

        if (sUserProfileDir != "") sMessage = Regex.Replace(sMessage, Regex.Escape(sUserProfileDir), "%USERPROFILE%", RegexOptions.IgnoreCase);
        sLine = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + sMessage;
        lock (oLogLock)
        {
            Console.WriteLine(sLine);
            if (sLogFilePath != "")
            {
                try { File.AppendAllText(sLogFilePath, sLine + "\r\n", utf8NoBom); }
                catch (Exception) { }
            }
        }
    }

    // Thread-safe jitter, because the Random class is not safe to share.
    static int jitterMs(int iMax)
    {
        lock (oRandom) { return oRandom.Next(0, iMax); }
    }

    // Appends the three-section results summary to the log, once only.
    static void writeSummarySections()
    {
        if (bSummaryWritten) return;
        bSummaryWritten = true;
        log("");
        log("Downloaded: " + lDownloaded.Count);
        foreach (string sItem in lDownloaded) log("  " + sItem);
        log("Failed: " + lFailed.Count);
        foreach (string sItem in lFailed) log("  " + sItem);
        log("Skipped: " + lSkipped.Count);
        foreach (string sItem in lSkipped) log("  " + sItem);
    }

    // Returns the standard Windows Downloads folder through the known-folder api.
    static string downloadsFolder()
    {
        string sPath;
        Guid oFolderId;
        IntPtr hPath;

        oFolderId = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        if (SHGetKnownFolderPath(oFolderId, 0, IntPtr.Zero, out hPath) == 0)
        {
            sPath = Marshal.PtrToStringUni(hPath);
            Marshal.FreeCoTaskMem(hPath);
            return sPath;
        }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
    }

    // Returns true if any Edge process is currently running.
    static bool isEdgeRunning()
    {
        try { return Process.GetProcessesByName("msedge").Length > 0; }
        catch (Exception) { return false; }
    }

    // Launches Edge with remote debugging.  The sign-in, sync, and extension
    // switches come straight from urlFido, which learned the hard way that a
    // fresh profile otherwise signs itself into the Windows account, offers
    // to sync the profile, and updates extensions, producing windows the user
    // has to dismiss.  Several spellings of Edge's implicit sign-in feature
    // are listed because the name has varied across versions; unrecognized
    // feature names are ignored rather than rejected.  The Preferences file
    // is seeded only for the dedicated profile, never for the user's own.
    static void launchEdge(string sUserDataDir, string sDownloadDir, bool bMainProfile)
    {
        string sEdgePath;
        List<string> lArgs;
        ProcessStartInfo oStartInfo;

        sEdgePath = File.Exists(sEdgePathPrimary) ? sEdgePathPrimary : sEdgePathSecondary;
        if (!File.Exists(sEdgePath)) throw new Exception("Microsoft Edge was not found at either standard location");
        if (!bMainProfile) seedPreferences(sUserDataDir, sDownloadDir);
        lArgs = new List<string>();
        lArgs.Add("--disable-component-update");
        lArgs.Add("--disable-extensions");
        lArgs.Add("--disable-sync");
        lArgs.Add("--disable-background-networking");
        lArgs.Add("--disable-client-side-phishing-detection");
        lArgs.Add("--disable-default-apps");
        lArgs.Add("--no-service-autorun");
        lArgs.Add("--metrics-recording-only");
        lArgs.Add("--disable-features=msImplicitSignin,msEdgeImplicitSignin,EdgeAutoSignIn,SyncPromo,SigninPromo,PrivacySandboxSettings4,SearchEngineChoiceScreen");
        lArgs.Add("--mute-audio");
        lArgs.Add("--no-default-browser-check");
        lArgs.Add("--no-first-run");
        lArgs.Add("--remote-debugging-port=" + iDebugPort);
        lArgs.Add("--user-data-dir=\"" + sUserDataDir + "\"");
        lArgs.Add("\"" + sStartUrl + "\"");
        log("Launching Edge (" + (bMainProfile ? "main profile" : "dedicated profile") + "): " + sEdgePath + " " + string.Join(" ", lArgs));
        oStartInfo = new ProcessStartInfo();
        oStartInfo.FileName = sEdgePath;
        oStartInfo.Arguments = string.Join(" ", lArgs);
        oStartInfo.UseShellExecute = false;
        oStartInfo.CreateNoWindow = true;
        oEdgeProcess = Process.Start(oStartInfo);
    }

    // Seeds the dedicated profile's Default\Preferences file before Edge first
    // opens the profile, urlFido's belt-and-braces companion to the launch
    // switches: with sign-in and sync disallowed in the preferences, Edge
    // never reaches the state where it would offer to sync the profile at
    // all.  Runs only when the Preferences file does not exist yet, so a
    // profile that already holds the Audible login is never touched.  Written
    // as UTF-8 without a byte order mark, because Chromium rejects one there.
    static void seedPreferences(string sProfileDir, string sDownloadDir)
    {
        string sDefaultDir, sPreferencesPath;
        Dictionary<string, object> dBrowser, dDownload, dPreferences, dProfile, dSignin, dSync;

        sDefaultDir = Path.Combine(sProfileDir, "Default");
        sPreferencesPath = Path.Combine(sDefaultDir, "Preferences");
        if (File.Exists(sPreferencesPath)) { log("The profile's Preferences file already exists, so leaving it alone"); return; }
        dBrowser = new Dictionary<string, object>();
        dBrowser["has_seen_welcome_page"] = true;
        dDownload = new Dictionary<string, object>();
        dDownload["prompt_for_download"] = false;
        dDownload["default_directory"] = sDownloadDir;
        dProfile = new Dictionary<string, object>();
        dProfile["password_manager_enabled"] = false;
        dProfile["exit_type"] = "Normal";
        dProfile["exited_cleanly"] = true;
        dSignin = new Dictionary<string, object>();
        dSignin["allowed"] = false;
        dSignin["allowed_on_next_startup"] = false;
        dSync = new Dictionary<string, object>();
        dSync["requested"] = false;
        dSync["has_setup_completed"] = false;
        dPreferences = new Dictionary<string, object>();
        dPreferences["browser"] = dBrowser;
        dPreferences["credentials_enable_service"] = false;
        dPreferences["download"] = dDownload;
        dPreferences["profile"] = dProfile;
        dPreferences["signin"] = dSignin;
        dPreferences["sync"] = dSync;
        Directory.CreateDirectory(sDefaultDir);
        File.WriteAllText(sPreferencesPath, jsonCodec.Serialize(dPreferences), new UTF8Encoding(false));
        log("Seeded a fresh profile Preferences file that disallows sign-in and sync");
    }

    // Returns true if the local Edge debugging port answers.
    static bool debugPortReachable()
    {
        try
        {
            using (WebClient webClientProbe = new WebClient()) { webClientProbe.DownloadString("http://127.0.0.1:" + iDebugPort + "/json/version"); }
            return true;
        }
        catch (Exception) { return false; }
    }

    // Polls the debugging port after launching Edge, up to the given timeout.
    static bool waitForDebugPort(int iTimeoutMs)
    {
        int iElapsedMs;

        iElapsedMs = 0;
        while (iElapsedMs < iTimeoutMs)
        {
            Thread.Sleep(500);
            iElapsedMs = iElapsedMs + 500;
            if (debugPortReachable()) return true;
        }
        return false;
    }

    // Finds the WebSocket url of a page target, preferring one already on Audible.
    static string pageWebSocketUrl()
    {
        object[] aTargets;
        string sFallback, sJson, sType, sUrl;
        Dictionary<string, object> dTarget;

        using (WebClient webClientJson = new WebClient()) { sJson = webClientJson.DownloadString("http://127.0.0.1:" + iDebugPort + "/json"); }
        aTargets = (object[]) jsonCodec.DeserializeObject(sJson);
        sFallback = "";
        foreach (object oTarget in aTargets)
        {
            dTarget = (Dictionary<string, object>) oTarget;
            sType = Convert.ToString(dTarget["type"]);
            if (sType != "page") continue;
            sUrl = dTarget.ContainsKey("url") ? Convert.ToString(dTarget["url"]) : "";
            if (sFallback == "") sFallback = Convert.ToString(dTarget["webSocketDebuggerUrl"]);
            if (sUrl.Contains("audible")) return Convert.ToString(dTarget["webSocketDebuggerUrl"]);
        }
        if (sFallback == "") throw new Exception("No page target was found on the Edge debugging port");
        return sFallback;
    }

    // Connects the raw CDP WebSocket.
    static async Task connectCdp()
    {
        string sWsUrl;

        sWsUrl = pageWebSocketUrl();
        log("Connecting to CDP target: " + sWsUrl);
        wsCdp = new ClientWebSocket();
        await wsCdp.ConnectAsync(new Uri(sWsUrl), CancellationToken.None);
    }

    // Enables the CDP domains this program uses; also called after reconnects.
    static async Task enableDomains()
    {
        await cdpSendOnce("Page.enable", null);
        await cdpSendOnce("Runtime.enable", null);
        await cdpSendOnce("Network.enable", null);
        log("CDP connected with the Page, Runtime, and Network domains enabled");
    }

    // Drops the broken WebSocket, connects to a fresh page target, and
    // re-enables the domains.  This recovers when the socket is forcibly
    // closed, for example after the user closes an Edge window.
    static async Task reconnectCdp()
    {
        try { if (wsCdp != null) wsCdp.Dispose(); }
        catch (Exception) { }
        wsCdp = null;
        await connectCdp();
        await enableDomains();
    }

    // Sends one CDP command and waits for its matching reply.  If the
    // WebSocket has died, reconnects to a fresh page target once and retries
    // the command, so a closed stray window does not end the whole run.
    static async Task<Dictionary<string, object>> cdpSend(string sMethod, Dictionary<string, object> dParams)
    {
        bool bReconnect;

        bReconnect = false;
        try { return await cdpSendOnce(sMethod, dParams); }
        catch (WebSocketException oException)
        {
            log("CDP connection lost during " + sMethod + ": " + oException.Message);
            bReconnect = true;
        }
        catch (ObjectDisposedException)
        {
            log("CDP connection was already closed before " + sMethod);
            bReconnect = true;
        }
        catch (InvalidOperationException oException)
        {
            log("CDP connection was unusable during " + sMethod + ": " + oException.Message);
            bReconnect = true;
        }
        if (bReconnect) { log("Reconnecting to Edge and retrying " + sMethod); await reconnectCdp(); }
        return await cdpSendOnce(sMethod, dParams);
    }

    // Sends one CDP command over the current WebSocket, skipping events.
    static async Task<Dictionary<string, object>> cdpSendOnce(string sMethod, Dictionary<string, object> dParams)
    {
        byte[] aOutBytes;
        int iId;
        string sMessage, sReply;
        Dictionary<string, object> dMessage, dReply;

        iId = iNextId;
        iNextId = iNextId + 1;
        dMessage = new Dictionary<string, object>();
        dMessage["id"] = iId;
        dMessage["method"] = sMethod;
        if (dParams != null) dMessage["params"] = dParams;
        sMessage = jsonCodec.Serialize(dMessage);
        aOutBytes = Encoding.UTF8.GetBytes(sMessage);
        await wsCdp.SendAsync(new ArraySegment<byte>(aOutBytes), WebSocketMessageType.Text, true, CancellationToken.None);
        while (true)
        {
            sReply = await receiveMessage();
            dReply = (Dictionary<string, object>) jsonCodec.DeserializeObject(sReply);
            if (!dReply.ContainsKey("id")) continue;
            if (Convert.ToInt32(dReply["id"]) != iId) continue;
            if (dReply.ContainsKey("error")) throw new Exception("CDP error for " + sMethod + ": " + jsonCodec.Serialize(dReply["error"]));
            if (dReply.ContainsKey("result")) return (Dictionary<string, object>) dReply["result"];
            return new Dictionary<string, object>();
        }
    }

    // Receives one complete WebSocket text message, however many frames it takes.
    static async Task<string> receiveMessage()
    {
        byte[] aBuffer;
        MemoryStream msWhole;
        WebSocketReceiveResult oChunk;

        aBuffer = new byte[65536];
        msWhole = new MemoryStream();
        while (true)
        {
            oChunk = await wsCdp.ReceiveAsync(new ArraySegment<byte>(aBuffer), CancellationToken.None);
            msWhole.Write(aBuffer, 0, oChunk.Count);
            if (oChunk.EndOfMessage) break;
        }
        return Encoding.UTF8.GetString(msWhole.ToArray());
    }

    // Evaluates a JavaScript expression in the page and returns its string value.
    static async Task<string> evaluate(string sExpression)
    {
        Dictionary<string, object> dParams, dResult, dValue;

        dParams = new Dictionary<string, object>();
        dParams["expression"] = sExpression;
        dParams["returnByValue"] = true;
        dParams["awaitPromise"] = true;
        dResult = await cdpSend("Runtime.evaluate", dParams);
        if (!dResult.ContainsKey("result")) return "";
        dValue = (Dictionary<string, object>) dResult["result"];
        if (!dValue.ContainsKey("value")) return "";
        return Convert.ToString(dValue["value"]);
    }

    // Navigates the CDP page to a url and waits for the document to load.
    static async Task navigate(string sUrl)
    {
        Dictionary<string, object> dParams;

        dParams = new Dictionary<string, object>();
        dParams["url"] = sUrl;
        log("Navigating to: " + sUrl);
        await cdpSend("Page.navigate", dParams);
        await waitForLoad();
    }

    // Polls document.readyState until complete, up to iNavigateTimeoutMs.
    static async Task waitForLoad()
    {
        int iElapsedMs;
        string sState;

        iElapsedMs = 0;
        while (iElapsedMs < iNavigateTimeoutMs)
        {
            Thread.Sleep(500);
            iElapsedMs = iElapsedMs + 500;
            sState = await evaluate("document.readyState");
            if (sState == "complete") { log("Page load complete after about " + iElapsedMs + " ms"); return; }
        }
        log("Warning: the page did not reach readyState complete within " + iNavigateTimeoutMs + " ms, so continuing anyway");
    }

    // Returns the JavaScript that scans one library page.  It reports the
    // companion PDF links, the next-page link, and whether the next button is
    // disabled, which Audible marks with the bc-button-disabled class on the
    // last page.
    static string scanScript()
    {
        return @"(function() {
    var oResult = { pdfs: [], nextHref: """", nextDisabled: false };
    document.querySelectorAll(""a[href*='/companion-file/']"").forEach(function(oLink) {
        var bKnown = oResult.pdfs.some(function(oItem) { return oItem.href == oLink.href; });
        if (bKnown) return;
        var sTitle = """";
        var oRow = oLink.closest(""[id^='adbl-library-content-row-']"");
        if (oRow) {
            var oHeadline = oRow.querySelector(""[class*='bc-size-headline']"");
            if (oHeadline) sTitle = oHeadline.textContent.replace(/\s+/g, "" "").trim();
        }
        oResult.pdfs.push({ href: oLink.href, title: sTitle });
    });
    var oNext = document.querySelector("".nextButton"");
    if (oNext) {
        if (oNext.className.indexOf(""bc-button-disabled"") >= 0) oResult.nextDisabled = true;
        var oAnchor = oNext.querySelector(""a[href]"");
        if (oAnchor) oResult.nextHref = oAnchor.href;
    }
    oResult.totalPages = 0;
    document.querySelectorAll("".pagingElements a, .pagingElements span, ul.pageNumberElement a, ul.pageNumberElement span"").forEach(function(oOne) {
        var iValue = parseInt(oOne.textContent.replace(/[^0-9]/g, """"), 10);
        if (iValue > oResult.totalPages) oResult.totalPages = iValue;
    });
    return JSON.stringify(oResult);
})()";
    }

    // Returns the JavaScript that harvests one library page's catalog rows:
    // for each title, its asin, detail-page link, authors, narrators, series,
    // merchandising summary, and listening progress, using Audible's own
    // markup markers.  Names are deduplicated because the responsive layout
    // repeats some markup for desktop and mobile widths.
    static string catalogScript()
    {
        return @"(function() {
    var aRows = [];
    var pushUnique = function(aList, sName, sUrl) {
        if (sName == """") return;
        var bKnown = aList.some(function(oItem) { return oItem.name == sName; });
        if (!bKnown) aList.push({ name: sName, url: sUrl });
    };
    document.querySelectorAll(""[id^='adbl-library-content-row-']"").forEach(function(oRow) {
        var sAsin = oRow.id.substring(""adbl-library-content-row-"".length);
        var oItem = { asin: sAsin, title: """", url: """", authors: [], narrators: [], series: [], summary: """", finished: false, timeLeft: """" };
        var oHead = oRow.querySelector(""[class*='bc-size-headline']"");
        if (oHead) oItem.title = oHead.textContent.replace(/\s+/g, "" "").trim();
        var oTitleLink = oRow.querySelector(""a[href*='/pd/']"");
        if (oTitleLink) oItem.url = oTitleLink.href;
        oRow.querySelectorAll(""a[href*='/author/']"").forEach(function(oA) { pushUnique(oItem.authors, oA.textContent.replace(/\s+/g, "" "").trim(), oA.href); });
        oRow.querySelectorAll(""a[href*='searchNarrator=']"").forEach(function(oA) { pushUnique(oItem.narrators, oA.textContent.replace(/\s+/g, "" "").trim(), oA.href); });
        oRow.querySelectorAll(""a[href*='/series/']"").forEach(function(oA) { pushUnique(oItem.series, oA.textContent.replace(/\s+/g, "" "").trim(), oA.href); });
        var oSummary = oRow.querySelector("".merchandisingSummary"");
        if (oSummary) oItem.summary = oSummary.textContent.replace(/\s+/g, "" "").trim();
        var oFinished = document.getElementById(""time-remaining-finished-"" + sAsin);
        if (oFinished && oFinished.className.indexOf(""bc-pub-hidden"") < 0) oItem.finished = true;
        if (!oItem.finished) {
            oRow.querySelectorAll(""span"").forEach(function(oSpan) {
                if (oItem.timeLeft != """") return;
                var sText = oSpan.textContent.replace(/\s+/g, "" "").trim();
                if (sText.length < 30 && sText.indexOf("" left"") == sText.length - 5 && sText.length > 5) oItem.timeLeft = sText;
            });
        }
        aRows.push(oItem);
    });
    return JSON.stringify(aRows);
})()";
    }

    // Builds a Cookie header from the browser's cookies for audible.com, so the
    // direct HTTP downloads ride the user's logged-in session.
    static async Task<string> cookieHeader()
    {
        int iCount;
        string sName, sValue;
        Dictionary<string, object> dCookie, dParams, dResult;
        StringBuilder sbHeader;

        dParams = new Dictionary<string, object>();
        dParams["urls"] = new string[] { sLibraryUrl };
        dResult = await cdpSend("Network.getCookies", dParams);
        iCount = 0;
        sbHeader = new StringBuilder();
        foreach (object oCookie in (IEnumerable) dResult["cookies"])
        {
            dCookie = (Dictionary<string, object>) oCookie;
            sName = Convert.ToString(dCookie["name"]);
            sValue = Convert.ToString(dCookie["value"]);
            if (sbHeader.Length > 0) sbHeader.Append("; ");
            sbHeader.Append(sName + "=" + sValue);
            iCount = iCount + 1;
        }
        log("Collected " + iCount + " browser cookies for the download requests");
        return sbHeader.ToString();
    }

    // Downloads one companion PDF, trying fallbacks when the companion-file
    // url answers with a web page: first any direct PDF link inside that
    // page, then the newer asin that the title's product page redirect
    // reveals, in case Amazon renamed the title since its original release.
    static void downloadCompanion(string sUrl, string sBookTitle, string sCookieHeader, string sDownloadDir)
    {
        int iStatus;
        string sAlternateUrl, sDirectUrl, sModernAsin, sOldAsin;

        iStatus = attemptDownload(sUrl, sBookTitle, sCookieHeader, sDownloadDir);
        if (iStatus == iStatusDownloaded || iStatus == iStatusSkipped) { dPdfNames[asinFromUrl(sUrl)] = sLastSavedName; return; }
        if (iStatus == iStatusFailedHtml)
        {
            sDirectUrl = pdfUrlFromText(sLastFailureBody);
            if (sDirectUrl != "")
            {
                log("The failure page contains a direct PDF link, so trying it: " + sDirectUrl);
                iStatus = attemptDownload(sDirectUrl, sBookTitle, sCookieHeader, sDownloadDir);
                if (iStatus == iStatusDownloaded || iStatus == iStatusSkipped) { dPdfNames[asinFromUrl(sUrl)] = sLastSavedName; return; }
            }
            sOldAsin = asinFromUrl(sUrl);
            sModernAsin = modernAsin(sOldAsin, sCookieHeader);
            if (sModernAsin != "" && sModernAsin != sOldAsin)
            {
                sAlternateUrl = "https://www.audible.com/companion-file/" + sModernAsin;
                log("Trying the newer asin naming for this title: " + sAlternateUrl);
                iStatus = attemptDownload(sAlternateUrl, sBookTitle, sCookieHeader, sDownloadDir);
                if (iStatus == iStatusDownloaded || iStatus == iStatusSkipped) { dPdfNames[asinFromUrl(sUrl)] = sLastSavedName; return; }
            }
        }
        lFailed.Add(sUrl);
        dFailedTitles[sUrl] = sBookTitle;
        dPdfNames[asinFromUrl(sUrl)] = "";
    }

    // Makes one direct HTTP attempt at a url and returns a status constant.
    // A web-page response saves its body beside the log for diagnosis and
    // leaves it in sLastFailureBody for the fallback logic.  The saved file
    // gets a human-friendly name: the book title from the library page when
    // one is available, else the Title metadata inside the PDF, else the
    // server's code name.  A file downloaded under a code name by an earlier
    // run is renamed in place instead of downloaded again, and a timed
    // message box announces each real download by its base file name.
    static int attemptDownload(string sUrl, string sBookTitle, string sCookieHeader, string sDownloadDir)
    {
        string sContentType, sDiagnosticPath, sFriendlyPath, sMetaPath, sMetaRoot, sRoot, sServerName, sServerPath, sTargetPath;
        HttpWebRequest httpRequest;
        StreamReader fReader;

        log("Requesting: " + sUrl + (sBookTitle == "" ? "" : " for the title " + sBookTitle));
        try
        {
            httpRequest = (HttpWebRequest) WebRequest.Create(sUrl);
            if (sCookieHeader != "") httpRequest.Headers.Add("Cookie", sCookieHeader);
            httpRequest.UserAgent = sUserAgent;
            httpRequest.Referer = sLibraryUrl;
            httpRequest.AllowAutoRedirect = true;
            httpRequest.Timeout = iHttpTimeoutMs;
            using (HttpWebResponse httpResponse = (HttpWebResponse) httpRequest.GetResponse())
            {
                sContentType = httpResponse.ContentType == null ? "" : httpResponse.ContentType.ToLower();
                sServerName = fileNameFromResponse(httpResponse, sUrl);
                sServerPath = Path.Combine(sDownloadDir, sServerName);
                sRoot = cleanRoot(sBookTitle);
                sFriendlyPath = sRoot == "" ? "" : Path.Combine(sDownloadDir, sRoot + ".pdf");
                log("Response url: " + httpResponse.ResponseUri.ToString() + ", content type: " + sContentType + ", server file name: " + sServerName);
                if (sContentType.Contains("html"))
                {
                    fReader = new StreamReader(httpResponse.GetResponseStream());
                    sLastFailureBody = fReader.ReadToEnd();
                    fReader.Close();
                    if (sLastFailureBody.Length > 500000) sLastFailureBody = sLastFailureBody.Substring(0, 500000);
                    sDiagnosticPath = saveFailureBody(sLastFailureBody, sUrl);
                    log("Failed, because the response was a web page rather than a file: " + sUrl);
                    if (sDiagnosticPath != "") log("Saved the web page body for diagnosis: " + sDiagnosticPath);
                    return iStatusFailedHtml;
                }
                if (sFriendlyPath != "" && File.Exists(sFriendlyPath))
                {
                    log("Skipped, because the file already exists: " + sFriendlyPath);
                    lSkipped.Add(Path.GetFileName(sFriendlyPath));
                    sLastSavedName = Path.GetFileName(sFriendlyPath);
                    showTimedMessageBox(foundFileAnnouncement("Found", Path.GetFileName(sFriendlyPath), ensureHtmVersion(sFriendlyPath)));
                    return iStatusSkipped;
                }
                if (File.Exists(sServerPath))
                {
                    if (sFriendlyPath == "")
                    {
                        log("Skipped, because the file already exists: " + sServerPath);
                        lSkipped.Add(sServerName);
                        sLastSavedName = sServerName;
                        showTimedMessageBox(foundFileAnnouncement("Found", sServerName, ensureHtmVersion(sServerPath)));
                        return iStatusSkipped;
                    }
                    File.Move(sServerPath, sFriendlyPath);
                    log("Renamed the earlier download " + sServerName + " to " + Path.GetFileName(sFriendlyPath));
                    lSkipped.Add(sServerName + " renamed to " + Path.GetFileName(sFriendlyPath));
                    sLastSavedName = Path.GetFileName(sFriendlyPath);
                    showTimedMessageBox(foundFileAnnouncement("Renamed to", Path.GetFileName(sFriendlyPath), ensureHtmVersion(sFriendlyPath)));
                    return iStatusSkipped;
                }
                sTargetPath = sFriendlyPath != "" ? sFriendlyPath : sServerPath;
                showTimedMessageBox(Path.GetFileNameWithoutExtension(sTargetPath));
                using (FileStream fOut = new FileStream(sTargetPath, FileMode.Create, FileAccess.Write)) { httpResponse.GetResponseStream().CopyTo(fOut); }
                if (sFriendlyPath == "")
                {
                    sMetaRoot = cleanRoot(pdfTitleFromFile(sTargetPath));
                    if (sMetaRoot != "")
                    {
                        sMetaPath = uniquePdfPath(sDownloadDir, sMetaRoot);
                        File.Move(sTargetPath, sMetaPath);
                        log("Renamed to " + Path.GetFileName(sMetaPath) + " using the PDF's own title metadata");
                        sTargetPath = sMetaPath;
                    }
                }
                log("Downloaded: " + sTargetPath);
                lDownloaded.Add(Path.GetFileName(sTargetPath));
                sLastSavedName = Path.GetFileName(sTargetPath);
                ensureHtmVersion(sTargetPath);
                return iStatusDownloaded;
            }
        }
        catch (Exception oException)
        {
            log("Failed: " + sUrl + " -- " + oException.Message);
            return iStatusFailed;
        }
    }

    // Saves the HTML body of a failed companion-file response beside the log,
    // so the reason a title's PDF was refused can be inspected afterward.
    static string saveFailureBody(string sBody, string sUrl)
    {
        string sName, sPath;

        try
        {
            sName = asinFromUrl(sUrl);
            if (sName == "") sName = "unknown";
            sPath = Path.Combine(dataDir(), "bookFido_fail_" + sName + ".html");
            File.WriteAllText(sPath, sBody, new UTF8Encoding(true));
            return sPath;
        }
        catch (Exception oException)
        {
            log("Could not save the failure body: " + oException.Message);
            return "";
        }
    }

    // Returns the last path segment of a url, which for companion-file and
    // product urls is the asin.
    static string asinFromUrl(string sUrl)
    {
        int iPos;
        string sTrimmed;

        sTrimmed = sUrl.TrimEnd('/');
        iPos = sTrimmed.IndexOf("?");
        if (iPos >= 0) sTrimmed = sTrimmed.Substring(0, iPos);
        return sTrimmed.Substring(sTrimmed.LastIndexOf('/') + 1);
    }

    // Returns true if a string has the ten-character alphanumeric shape of an
    // Amazon asin.
    static bool looksLikeAsin(string sCandidate)
    {
        if (sCandidate.Length != 10) return false;
        foreach (char chOne in sCandidate) { if (!char.IsLetterOrDigit(chOne)) return false; }
        return true;
    }

    // Asks the title's product page for the asin Amazon uses today.  The
    // request to /pd/<old asin> redirects to the canonical product url, whose
    // last path segment is the current asin under the newer naming scheme.
    static string modernAsin(string sOldAsin, string sCookieHeader)
    {
        string sCandidate, sProductUrl;
        HttpWebRequest httpRequest;

        if (!looksLikeAsin(sOldAsin)) return "";
        sProductUrl = "https://www.audible.com/pd/" + sOldAsin;
        log("Consulting the product page for a newer asin: " + sProductUrl);
        try
        {
            httpRequest = (HttpWebRequest) WebRequest.Create(sProductUrl);
            if (sCookieHeader != "") httpRequest.Headers.Add("Cookie", sCookieHeader);
            httpRequest.UserAgent = sUserAgent;
            httpRequest.AllowAutoRedirect = true;
            httpRequest.Timeout = iHttpTimeoutMs;
            using (HttpWebResponse httpResponse = (HttpWebResponse) httpRequest.GetResponse())
            {
                sCandidate = asinFromUrl(httpResponse.ResponseUri.AbsolutePath);
                log("The product page resolved to: " + httpResponse.ResponseUri.ToString());
                if (looksLikeAsin(sCandidate)) return sCandidate;
            }
        }
        catch (Exception oException)
        {
            log("The product page lookup failed: " + oException.Message);
        }
        return "";
    }

    // Finds the first https url ending in .pdf inside a block of text, such as
    // the body of a failure page.
    static string pdfUrlFromText(string sText)
    {
        int iEnd, iPos, iStart;
        string sCandidate;

        iPos = 0;
        while (true)
        {
            iPos = sText.IndexOf(".pdf", iPos, StringComparison.OrdinalIgnoreCase);
            if (iPos < 0) return "";
            iStart = sText.LastIndexOf("https://", iPos, StringComparison.OrdinalIgnoreCase);
            if (iStart >= 0 && iPos - iStart < 2000)
            {
                iEnd = sText.IndexOfAny(new char[] { '"', (char) 39, '<', '>', ' ', (char) 13, (char) 10 }, iPos);
                if (iEnd < 0) iEnd = sText.Length;
                sCandidate = sText.Substring(iStart, iEnd - iStart);
                if (!sCandidate.Contains("\"")) return sCandidate.Replace("&amp;", "&");
            }
            iPos = iPos + 4;
        }
    }

    // Fetches a JSON document from a public web service and returns it as a
    // dictionary, or null on any failure, which is logged and tolerated in
    // the manner of the biblio tool's safeFetchJson.
    static Dictionary<string, object> fetchJson(string sUrl)
    {
        int iAttempt, iWaitSeconds;
        string sBody, sRetryAfter;
        HttpWebRequest httpRequest;
        HttpWebResponse httpErrorResponse;
        StreamReader fReader;
        WebException webException;

        bLastFetchRateLimited = false;
        for (iAttempt = 1; iAttempt <= 2; iAttempt = iAttempt + 1)
        {
            try
            {
                httpRequest = (HttpWebRequest) WebRequest.Create(sUrl);
                httpRequest.UserAgent = sApiUserAgent;
                httpRequest.Accept = "application/json";
                httpRequest.Timeout = 15000;
                httpRequest.ReadWriteTimeout = 15000;
                using (HttpWebResponse httpResponse = (HttpWebResponse) httpRequest.GetResponse())
                {
                    fReader = new StreamReader(httpResponse.GetResponseStream());
                    sBody = fReader.ReadToEnd();
                    fReader.Close();
                }
                return (Dictionary<string, object>) jsonCodec.DeserializeObject(sBody);
            }
            catch (Exception oException)
            {
                webException = oException as WebException;
                httpErrorResponse = webException == null ? null : webException.Response as HttpWebResponse;
                if (httpErrorResponse != null && (int) httpErrorResponse.StatusCode == 429)
                {
                    bLastFetchRateLimited = true;
                    if (iAttempt < 2 && !bStopLanes)
                    {
                        sRetryAfter = httpErrorResponse.Headers["Retry-After"];
                        iWaitSeconds = 30;
                        if (sRetryAfter != null && int.TryParse(sRetryAfter, out iWaitSeconds)) iWaitSeconds = Math.Min(Math.Max(iWaitSeconds, 10), 60);
                        else iWaitSeconds = 30;
                        log("The service asked for a pause (429); waiting " + iWaitSeconds + " seconds before retrying " + sUrl);
                        Thread.Sleep(iWaitSeconds * 1000);
                        continue;
                    }
                }
                log("Web request failed for " + sUrl + " -- " + oException.Message);
                return null;
            }
        }
        return null;
    }

    // Returns a nested dictionary member, or null.
    static Dictionary<string, object> dictAt(Dictionary<string, object> dParent, string sKey)
    {
        if (dParent == null || !dParent.ContainsKey(sKey)) return null;
        return dParent[sKey] as Dictionary<string, object>;
    }

    // Returns the first element of a JSON array member as a dictionary, or null.
    static Dictionary<string, object> firstAt(Dictionary<string, object> dParent, string sKey)
    {
        object[] aItems;

        if (dParent == null || !dParent.ContainsKey(sKey)) return null;
        aItems = dParent[sKey] as object[];
        if (aItems == null || aItems.Length == 0) return null;
        return aItems[0] as Dictionary<string, object>;
    }

    // Capitalizes the first letter of a value such as a language name.
    static string capitalizeFirst(string sText)
    {
        if (sText.Length == 0) return sText;
        return char.ToUpper(sText[0]) + sText.Substring(1);
    }

    // Formats a runtime in minutes the way Audible states lengths.
    static string lengthText(int iMinutes)
    {
        if (iMinutes <= 0) return "";
        if (iMinutes < 60) return iMinutes + " mins";
        if (iMinutes % 60 == 0) return (iMinutes / 60) + " hrs";
        return (iMinutes / 60) + " hrs and " + (iMinutes % 60) + " mins";
    }

    // Returns true when the saved snapshot proves the library unchanged:
    // the page count matches and page one shows exactly the saved titles in
    // the saved order.  The library's default sort puts new titles first,
    // so an identical first page is decisive evidence.
    static bool libraryUnchanged()
    {
        int iOne;

        if (!bSavedWalkComplete) return false;
        if (aSavedRows == null || lSavedPage1Asins == null || lSavedPage1Asins.Count == 0) return false;
        if (iSavedTotalPages == 0 || iSavedTotalPages != iTotalPagesSeen) return false;
        if (lPage1AsinsSeen.Count != lSavedPage1Asins.Count) return false;
        for (iOne = 0; iOne < lSavedPage1Asins.Count; iOne = iOne + 1) { if (lPage1AsinsSeen[iOne] != lSavedPage1Asins[iOne]) return false; }
        return true;
    }

    // Returns true when every title on the current page is already in the
    // saved catalog, meaning the walk has caught up with known territory.
    static bool pageFullyKnown(List<string> lPageAsins)
    {
        if (lPageAsins.Count == 0) return false;
        foreach (string sOne in lPageAsins) { if (!savedAsins().Contains(sOne)) return false; }
        return true;
    }

    // The saved catalog's asins, built once on first use.
    static HashSet<string> savedAsins()
    {
        Dictionary<string, object> dOne;

        if (setSavedAsinsCache != null) return setSavedAsinsCache;
        setSavedAsinsCache = new HashSet<string>();
        if (aSavedRows != null) foreach (object oRow in aSavedRows) { dOne = oRow as Dictionary<string, object>; if (dOne != null && dOne.ContainsKey("asin")) setSavedAsinsCache.Add(Convert.ToString(dOne["asin"])); }
        return setSavedAsinsCache;
    }

    // Adds every saved row the walk has not already harvested, in the saved
    // order, feeding each through the same enrichment filters, which only
    // re-fetch what a previous run never successfully checked.
    static void mergeSavedRows()
    {
        string sAsin;
        Dictionary<string, object> dRow;

        foreach (object oRow in aSavedRows)
        {
            dRow = oRow as Dictionary<string, object>;
            if (dRow == null || !dRow.ContainsKey("asin")) continue;
            sAsin = Convert.ToString(dRow["asin"]);
            if (setCatalogAsins.Contains(sAsin)) continue;
            setCatalogAsins.Add(sAsin);
            lCatalog.Add(dRow);
            enqueueForEnrichment(dRow);
        }
        log("Catalog rows now loaded: " + lCatalog.Count);
    }

    // After the walk, quietly settles every companion file the walk did not
    // handle live: files already on disk are checked and converted without
    // per-file announcements, and only genuinely missing files download
    // with the usual spoken announcements.  One summary box then reports
    // the whole sweep, so an unchanged library speaks once instead of a
    // hundred times.
    static void sweepCompanions(string sCookieHeader, string sDownloadDir)
    {
        int iDownloadedBefore, iFailedBefore, iHtmMade, iKnownDead, iMissing, iPresent, iStatus;
        string sAsin, sFilePath;
        string[] aInfo;

        iHtmMade = 0;
        iKnownDead = 0;
        iMissing = 0;
        iPresent = 0;
        iDownloadedBefore = lDownloaded.Count;
        foreach (string sKey in new List<string>(dCompanions.Keys))
        {
            sAsin = sKey;
            if (setCompanionsHandled.Contains(sAsin)) continue;
            if (setDeadCompanions.Contains(sAsin)) { iKnownDead = iKnownDead + 1; continue; }
            aInfo = dCompanions[sAsin];
            sFilePath = dPdfNames.ContainsKey(sAsin) && dPdfNames[sAsin] != "" ? Path.Combine(sDownloadDir, dPdfNames[sAsin]) : "";
            if (sFilePath != "" && File.Exists(sFilePath))
            {
                iPresent = iPresent + 1;
                iStatus = ensureHtmVersion(sFilePath);
                if (iStatus == iHtmStatusCreated) iHtmMade = iHtmMade + 1;
                continue;
            }
            iMissing = iMissing + 1;
            iFailedBefore = lFailed.Count;
            Thread.Sleep(iDelayDownloadMs + jitterMs(iJitterMaxMs));
            downloadCompanion(aInfo[0], aInfo[1], sCookieHeader, sDownloadDir);
            if (lFailed.Count > iFailedBefore) setDeadCompanions.Add(sAsin);
        }
        if (iPresent + iMissing + iKnownDead == 0) return;
        log("Companion sweep: " + iPresent + " already present, " + iMissing + " fetched, " + iHtmMade + " HTML versions made, " + iKnownDead + " known unavailable");
        if (iMissing == 0 && iHtmMade == 0) showTimedMessageBox("All " + iPresent + " companion files and HTML versions are present" + (iKnownDead > 0 ? "; " + iKnownDead + " known unavailable" : ""));
        else showTimedMessageBox("Companions: " + iPresent + " present, " + (lDownloaded.Count - iDownloadedBefore) + " downloaded, " + iHtmMade + " HTML made" + (iKnownDead > 0 ? ", " + iKnownDead + " known unavailable" : ""));
    }

    // The state file lives beside the exe and holds a machine-readable
    // snapshot of everything a run learned: the page count, the first
    // page's title order, the fully enriched catalog rows with their
    // per-service checked markers, the author biographies with their
    // checked lists, and the companion-file map.  A snapshot rewritten
    // atomically after each run was chosen over an append-style log: the
    // reader needs the latest complete picture, not a history to replay,
    // and a rewrite can never leave half-parsed old entries behind.
    static string statePath()
    {
        return Path.Combine(dataDir(), "bookFido.json");
    }

    // Adopts a state snapshot saved under either former name from the
    // program's bookFido era, by renaming it once, so nothing
    // already gathered is lost to the rename.
    static void adoptFormerStateName()
    {
        string sFormerPath;
        string[] aFormerNames;

        aFormerNames = new string[] { "GetAudibleInfo.json", "GetAudibleInfo_state.json" };
        foreach (string sFormerName in aFormerNames)
        {
            if (File.Exists(statePath())) return;
            sFormerPath = Path.Combine(dataDir(), sFormerName);
            if (!File.Exists(sFormerPath)) continue;
            try { File.Move(sFormerPath, statePath()); log("Renamed the saved state file from " + sFormerName + ", so nothing already gathered is lost"); }
            catch (Exception oException) { try { File.Copy(sFormerPath, statePath(), false); } catch (Exception) { } log("The saved state file " + sFormerName + " could not be renamed: " + oException.Message); }
        }
    }

    // The directory for the program's own files: the log, the state
    // snapshot, and diagnostic pages.  The exe's folder is used when it is
    // writable, preserving portable behavior; installed under Program
    // Files, which refuses writes, everything moves to LocalAppData.
    static string dataDir()
    {
        string sFormerDataDir, sProbePath;

        if (sDataDir != "") return sDataDir;
        sDataDir = AppDomain.CurrentDomain.BaseDirectory;
        try
        {
            sProbePath = Path.Combine(sDataDir, "bookFido_probe.tmp");
            File.WriteAllText(sProbePath, "probe");
            File.Delete(sProbePath);
        }
        catch (Exception)
        {
            sDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "bookFido");
            sFormerDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GetAudibleInfo");
            if (!Directory.Exists(sDataDir) && Directory.Exists(sFormerDataDir)) { try { Directory.Move(sFormerDataDir, sDataDir); log("Migrated the program data folder from its previous name"); } catch (Exception) { } }
            Directory.CreateDirectory(sDataDir);
        }
        return sDataDir;
    }

    // Loads the previous run's snapshot when one exists, merging the
    // author knowledge and file names into the live structures and keeping
    // the saved rows aside for the walk to consult.
    static void loadState()
    {
        int iSchema;
        string sJson, sValue;
        Dictionary<string, object> dAuthors, dNames, dSaved, dState;

        try
        {
            adoptFormerStateName();
            if (!File.Exists(statePath())) { log("No saved state file was found, so the whole library will be walked"); return; }
            sJson = File.ReadAllText(statePath(), new UTF8Encoding(false));
            dState = (Dictionary<string, object>) jsonCodec.DeserializeObject(sJson);
            iSavedTotalPages = dState.ContainsKey("totalPages") ? Convert.ToInt32(dState["totalPages"]) : 0;
            bSavedWalkComplete = dState.ContainsKey("walkComplete") && Convert.ToBoolean(dState["walkComplete"]);
            sSavedAtText = dState.ContainsKey("savedAt") ? Convert.ToString(dState["savedAt"]) : "";
            lSavedPage1Asins = new List<string>();
            if (dState.ContainsKey("page1Asins")) foreach (object oOne in (IEnumerable) dState["page1Asins"]) lSavedPage1Asins.Add(Convert.ToString(oOne));
            aSavedRows = dState.ContainsKey("catalog") ? dState["catalog"] as object[] : null;
            aSavedKindleRows = dState.ContainsKey("kindleCatalog") ? dState["kindleCatalog"] as object[] : null;
            aSavedGoodreadsRows = dState.ContainsKey("goodreadsCatalog") ? dState["goodreadsCatalog"] as object[] : null;
            aSavedBookshareRows = dState.ContainsKey("bookshareCatalog") ? dState["bookshareCatalog"] as object[] : null;
            // Schema 2 dropped the Google Books lane, took the publisher
            // from Audible's product_details or Open Library instead, and
            // retired the Categories field.  A state saved under the old
            // schema has its Audible and Open Library checked markers
            // cleared once, so both lanes revisit every title and gather
            // the publisher; Google leftovers are removed outright.
            iSchema = dState.ContainsKey("schema") ? Convert.ToInt32(dState["schema"]) : 1;
            if (aSavedRows != null && iSchema < 2)
            {
                log("The saved state predates schema 2, so the Audible and Open Library details will be gathered afresh for every title");
                foreach (object oRow in aSavedRows)
                {
                    dSaved = oRow as Dictionary<string, object>;
                    if (dSaved == null) continue;
                    dSaved.Remove("audibleChecked");
                    dSaved.Remove("openLibraryChecked");
                    dSaved.Remove("googleChecked");
                    dSaved.Remove("categories");
                    sValue = dSaved.ContainsKey("publisher") ? Convert.ToString(dSaved["publisher"]) : "";
                    if (sValue.EndsWith(" (Google Books)")) dSaved.Remove("publisher");
                }
            }
            dAuthors = dictAt(dState, "authorWikiBio");
            if (dAuthors != null) foreach (string sKey in dAuthors.Keys) dAuthorWikiBio[sKey] = Convert.ToString(dAuthors[sKey]);
            dAuthors = dictAt(dState, "authorOlBio");
            if (dAuthors != null) foreach (string sKey in dAuthors.Keys) dAuthorOlBio[sKey] = Convert.ToString(dAuthors[sKey]);
            dAuthors = dictAt(dState, "authorWikiUrl");
            if (dAuthors != null) foreach (string sKey in dAuthors.Keys) dAuthorWikiUrl[sKey] = Convert.ToString(dAuthors[sKey]);
            if (dState.ContainsKey("authorsCheckedWiki")) foreach (object oOne in (IEnumerable) dState["authorsCheckedWiki"]) setAuthorsCheckedWiki.Add(Convert.ToString(oOne));
            if (dState.ContainsKey("authorsCheckedOl")) foreach (object oOne in (IEnumerable) dState["authorsCheckedOl"]) setAuthorsCheckedOl.Add(Convert.ToString(oOne));
            if (dState.ContainsKey("deadCompanions")) foreach (object oOne in (IEnumerable) dState["deadCompanions"]) setDeadCompanions.Add(Convert.ToString(oOne));
            dNames = dictAt(dState, "pdfNames");
            if (dNames != null) foreach (string sKey in dNames.Keys) { if (!dPdfNames.ContainsKey(sKey)) dPdfNames[sKey] = Convert.ToString(dNames[sKey]); }
            dNames = dictAt(dState, "companions");
            if (dNames != null) foreach (string sKey in dNames.Keys) { if (!dCompanions.ContainsKey(sKey)) { object[] aPair = ((IEnumerable) dNames[sKey]).Cast<object>().ToArray(); dCompanions[sKey] = new string[] { Convert.ToString(aPair[0]), aPair.Length > 1 ? Convert.ToString(aPair[1]) : "" }; } }
            log("Loaded the saved state from " + sSavedAtText + ": " + (aSavedRows == null ? 0 : aSavedRows.Length) + " titles, " + iSavedTotalPages + " pages");
        }
        catch (Exception oException)
        {
            log("The saved state could not be read, so the whole library will be walked: " + oException.Message);
            aSavedRows = null;
            iSavedTotalPages = 0;
            lSavedPage1Asins = null;
        }
    }

    // Saves progress on the running clock: the state snapshot every two
    // minutes, because it alone carries progress that cannot be
    // regenerated and costs a fraction of a second, and the three catalog
    // documents every six minutes, because they are derived views that
    // cost seconds to rebuild and only convenience is lost if a halt
    // catches them stale.
    static void savePeriodically(bool bWithDocuments)
    {
        if ((DateTime.Now - dtLastStateSave).TotalMilliseconds >= iSaveStateMs)
        {
            dtLastStateSave = DateTime.Now;
            saveState();
        }
        if (bWithDocuments && sDownloadDirShared != "" && (DateTime.Now - dtLastDocumentsSave).TotalMilliseconds >= iSaveDocumentsMs)
        {
            dtLastDocumentsSave = DateTime.Now;
            try
            {
                buildLibraryFiles(sDownloadDirShared);
                writeWorkbook(sDownloadDirShared);
                log("Refreshed the catalog documents with the details gathered so far");
            }
            catch (Exception oException) { log("The periodic document refresh failed: " + oException.Message); }
        }
    }

    // Writes the snapshot after enrichment, via a temporary file replaced
    // in one step, so a crash mid-write can never corrupt the state.
    static void saveState()
    {
        string sJson, sTempPath;
        Dictionary<string, object> dState;
        List<string> lPage1;

        try
        {
            dState = new Dictionary<string, object>();
            dState["schema"] = 2;
            dState["savedAt"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            dState["walkComplete"] = bWalkComplete;
            dState["totalPages"] = iTotalPagesSeen;
            lPage1 = new List<string>();
            foreach (string sOne in lPage1AsinsSeen) lPage1.Add(sOne);
            dState["page1Asins"] = lPage1;
            dState["catalog"] = lCatalog;
            dState["kindleCatalog"] = lKindleCatalog;
            dState["goodreadsCatalog"] = lGoodreadsCatalog;
            dState["bookshareCatalog"] = lBookshareCatalog;
            lock (oAuthorLock)
            {
                dState["authorWikiBio"] = dAuthorWikiBio;
                dState["authorOlBio"] = dAuthorOlBio;
                dState["authorWikiUrl"] = dAuthorWikiUrl;
                dState["authorsCheckedWiki"] = new List<string>(setAuthorsCheckedWiki);
                dState["authorsCheckedOl"] = new List<string>(setAuthorsCheckedOl);
            }
            dState["pdfNames"] = dPdfNames;
            dState["deadCompanions"] = new List<string>(setDeadCompanions);
            dState["companions"] = dCompanions;
            sJson = jsonCodec.Serialize(dState);
            sTempPath = statePath() + ".tmp";
            File.WriteAllText(sTempPath, sJson, new UTF8Encoding(false));
            if (File.Exists(statePath())) File.Delete(statePath());
            File.Move(sTempPath, statePath());
            log("Saved the state snapshot to " + statePath() + ": " + lCatalog.Count + " titles");
        }
        catch (Exception oException)
        {
            log("The state snapshot could not be saved: " + oException.Message);
        }
    }

    // Starts the three enrichment lanes, one background thread per web
    // service, so detail gathering overlaps the library walk.  Each lane
    // keeps its own one-request-per-second pacing toward its one service.
    static void startEnrichmentLanes()
    {
        ServicePointManager.DefaultConnectionLimit = 8;
        log("Starting the three detail-gathering lanes: Audible catalog, Open Library, and Wikipedia");
        threadAudible = new Thread(audibleLane);
        threadAudible.IsBackground = true;
        threadAudible.Start();
        threadOpenLibrary = new Thread(openLibraryLane);
        threadOpenLibrary.IsBackground = true;
        threadOpenLibrary.Start();
        threadWikipedia = new Thread(wikipediaLane);
        threadWikipedia.IsBackground = true;
        threadWikipedia.Start();
    }

    // Hands a freshly harvested title to all four lanes, and each newly seen
    // author to the Wikipedia and Open Library lanes for a biography.  The
    // author names are read before the row is published, while this thread
    // still owns it exclusively.
    static void enqueueForEnrichment(Dictionary<string, object> dRow)
    {
        List<string[]> lAuthorPairs;

        lAuthorPairs = catalogLinks(dRow, "authors");
        if (!dRow.ContainsKey("audibleChecked")) { lock (queueAudible) { queueAudible.Enqueue(dRow); } iWorkTotal = iWorkTotal + 1; }
        if (!dRow.ContainsKey("openLibraryChecked")) { lock (queueOpenLibrary) { queueOpenLibrary.Enqueue(dRow); } iWorkTotal = iWorkTotal + 1; }
        if (!dRow.ContainsKey("wikipediaChecked")) { lock (queueWikipedia) { queueWikipedia.Enqueue(dRow); } iWorkTotal = iWorkTotal + 1; }
        foreach (string[] aPair in lAuthorPairs)
        {
            if (aPair[0] == "" || !setAuthorsQueued.Add(aPair[0])) continue;
            if (!setAuthorsCheckedWiki.Contains(aPair[0])) { lock (queueAuthorWikipedia) { queueAuthorWikipedia.Enqueue(aPair[0]); } iWorkTotal = iWorkTotal + 1; }
            if (!setAuthorsCheckedOl.Contains(aPair[0])) { lock (queueAuthorOpenLibrary) { queueAuthorOpenLibrary.Enqueue(aPair[0]); } iWorkTotal = iWorkTotal + 1; }
        }
    }

    // Takes the next waiting author name from a lane's author queue, or null.
    static string dequeueAuthor(Queue<string> queueLane)
    {
        lock (queueLane) { return queueLane.Count > 0 ? queueLane.Dequeue() : null; }
    }

    // Returns how many items a lane still has waiting, counting a lane's
    // author queue together with its title queue.
    static int laneRemaining(int iLane)
    {
        int iCount;

        iCount = 0;
        if (iLane == 0) lock (queueAudible) { iCount = queueAudible.Count; }
        if (iLane == 1) { lock (queueOpenLibrary) { iCount = queueOpenLibrary.Count; } lock (queueAuthorOpenLibrary) { iCount = iCount + queueAuthorOpenLibrary.Count; } }
        if (iLane == 2) { lock (queueWikipedia) { iCount = queueWikipedia.Count; } lock (queueAuthorWikipedia) { iCount = iCount + queueAuthorWikipedia.Count; } }
        return iCount;
    }

    // Names the services whose circuit breakers opened this run, joined for
    // prose, or an empty string when every lane ran to completion.
    static string skippedLaneNames()
    {
        List<string> lNames;

        lNames = new List<string>();
        if (aLaneSkipping[0]) lNames.Add("Audible's catalog service");
        if (aLaneSkipping[1]) lNames.Add("Open Library");
        if (aLaneSkipping[2]) lNames.Add("Wikipedia");
        if (lNames.Count == 0) return "";
        if (lNames.Count == 1) return lNames[0];
        return string.Join(", ", lNames.GetRange(0, lNames.Count - 1).ToArray()) + " and " + lNames[lNames.Count - 1];
    }

    // Waits for the lanes after the library walk ends, watching the actual
    // pace of each lane over a sliding window.  The run continues for as
    // long as steady progress is being made, announcing the overall
    // percentage complete and the estimated minutes remaining every report
    // interval, and saving progress to disk on the periodic policy.  Only
    // a lane that makes no progress across the whole watch window stops
    // the gathering early, with the reason logged and announced.
    static void waitForEnrichment()
    {
        bool bEtaReady, bStalledLane;
        int iDone, iEtaMinutes, iLane, iLaneSeconds, iOne, iPercent, iPredictedSeconds, iRemainingAll, iRemainingLane, iSinceReportMs, iTick, iWindowDone, iWindowDoneAll, iWindowSeconds;
        int[] aPrevRemaining;
        int[,] aWindowDeltas;
        string sBoxText, sMinutesWord;

        bHarvestDone = true;
        log("Library walk complete; waiting for the detail-gathering lanes to finish");
        aPrevRemaining = new int[3];
        aWindowDeltas = new int[3, iStallTicks];
        for (iLane = 0; iLane < 3; iLane = iLane + 1) { aPrevRemaining[iLane] = laneRemaining(iLane); for (iTick = 0; iTick < iStallTicks; iTick = iTick + 1) aWindowDeltas[iLane, iTick] = 1; }
        // The box body carries the overall gathering percentage from the
        // first moment of this phase, so an early announcement is never
        // percent-less and a stale walk percentage is never spoken with a
        // different meaning.
        iRemainingAll = aPrevRemaining[0] + aPrevRemaining[1] + aPrevRemaining[2];
        sProgressText = (iWorkTotal > 0 ? (iWorkTotal - iRemainingAll) * 100 / iWorkTotal : 100) + "%";
        iSinceReportMs = iDrainReportMs;
        iTick = 0;
        while ((threadAudible != null && threadAudible.IsAlive) || (threadOpenLibrary != null && threadOpenLibrary.IsAlive) || (threadWikipedia != null && threadWikipedia.IsAlive))
        {
            Thread.Sleep(1000);
            iSinceReportMs = iSinceReportMs + 1000;
            savePeriodically(true);
            if (iSinceReportMs >= iDrainReportMs)
            {
                iSinceReportMs = 0;
                iWindowSeconds = iDrainReportMs / 1000 * iStallTicks;
                iPredictedSeconds = 0;
                iRemainingAll = 0;
                iWindowDoneAll = 0;
                bStalledLane = false;
                for (iLane = 0; iLane < 3; iLane = iLane + 1)
                {
                    iRemainingLane = laneRemaining(iLane);
                    aWindowDeltas[iLane, iTick % iStallTicks] = aPrevRemaining[iLane] - iRemainingLane;
                    aPrevRemaining[iLane] = iRemainingLane;
                    iWindowDone = 0;
                    for (iOne = 0; iOne < iStallTicks; iOne = iOne + 1) iWindowDone = iWindowDone + aWindowDeltas[iLane, iOne];
                    iRemainingAll = iRemainingAll + iRemainingLane;
                    iWindowDoneAll = iWindowDoneAll + iWindowDone;
                    // A lane whose circuit breaker has opened is fast-draining
                    // its queue without contacting the service, so its backlog
                    // is not real work: it neither stalls the run nor belongs
                    // in the time projection.
                    if (iRemainingLane > 0 && iWindowDone <= 0 && !aLaneSkipping[iLane]) bStalledLane = true;
                    if (iRemainingLane > 0 && iWindowDone > 0 && !aLaneSkipping[iLane])
                    {
                        iLaneSeconds = iRemainingLane * iWindowSeconds / iWindowDone;
                        if (iLaneSeconds > iPredictedSeconds) iPredictedSeconds = iLaneSeconds;
                    }
                }
                iTick = iTick + 1;
                // The window arrives seeded with placeholder progress, so the
                // first estimates extrapolate from almost no evidence and can
                // be wildly high.  The estimate is spoken only once the whole
                // window holds real measurements.
                bEtaReady = iTick >= iStallTicks;
                iDone = iWorkTotal - iRemainingAll;
                iPercent = iWorkTotal > 0 ? iDone * 100 / iWorkTotal : 100;
                iEtaMinutes = (iPredictedSeconds + 59) / 60;
                sMinutesWord = iEtaMinutes == 1 ? " minute" : " minutes";
                log("Details gathered so far: Audible " + iAudibleDone + ", Open Library " + iOpenLibraryDone + ", Wikipedia " + iWikipediaDone + "; author pages checked: " + iAuthorsDone + " of " + setAuthorsQueued.Count + "; overall " + iPercent + " percent" + (bEtaReady && iEtaMinutes > 0 ? ", about " + iEtaMinutes + sMinutesWord + " remaining" : ""));
                sProgressText = iPercent + "%";
                sBoxText = bEtaReady && iEtaMinutes > 0 ? "About " + iEtaMinutes + sMinutesWord + " remain" : "Gathering details";
                showTimedMessageBox(sBoxText);
                if (bStalledLane && iTick >= iStallTicks)
                {
                    bStopLanes = true;
                    log("A detail-gathering lane has made no progress across the whole watch window, so the catalog will be built with the details collected so far");
                    showTimedMessageBox("Detail gathering stalled; building the catalog with what was collected");
                }
            }
        }
        if (threadAudible != null) threadAudible.Join();
        if (threadOpenLibrary != null) threadOpenLibrary.Join();
        if (threadWikipedia != null) threadWikipedia.Join();
        log("Detail gathering finished: Audible " + iAudibleDone + ", Open Library " + iOpenLibraryDone + ", Wikipedia " + iWikipediaDone + "; author pages checked: " + iAuthorsDone + " of " + setAuthorsQueued.Count);
    }

    // Takes the next waiting row from a lane's queue, or null when the queue
    // is empty.
    static Dictionary<string, object> dequeueRow(Queue<Dictionary<string, object>> queueLane)
    {
        lock (queueLane) { return queueLane.Count > 0 ? queueLane.Dequeue() : null; }
    }

    // The Audible catalog lane: ratings, length, release date, publisher,
    // and language for one title at a time, politely paced.
    static void audibleLane()
    {
        bool bSkipRemaining;
        int iStrikes;
        Dictionary<string, object> dRow;

        bSkipRemaining = false;
        iStrikes = 0;

        while (true)
        {
            if (bStopLanes) break;
            dRow = dequeueRow(queueAudible);
            if (dRow == null)
            {
                if (bHarvestDone) break;
                Thread.Sleep(iLanePollMs);
                continue;
            }
            if (!bSkipRemaining)
            {
                try { audibleDetails(dRow); }
                catch (Exception oException) { log("Audible details failed: " + oException.Message); }
                if (bLastFetchRateLimited) iStrikes = iStrikes + 1; else iStrikes = 0;
                if (iStrikes >= iRateLimitStrikesMax)
                {
                    bSkipRemaining = true;
                    aLaneSkipping[0] = true;
                    log("Audible's catalog service is rate limiting this address, so its remaining lookups are skipped and the run continues without them");
                    showTimedMessageBox("Audible's catalog service is rate limiting; skipping its remaining lookups this run");
                }
            }
            Interlocked.Increment(ref iAudibleDone);
            if (!bSkipRemaining) Thread.Sleep(iDelayApiMs + jitterMs(400));
        }
        log("The Audible details lane finished");
    }

    // The Open Library lane: year of first publication for titles, then
    // author biographies, all through one politely paced connection.
    static void openLibraryLane()
    {
        bool bSkipRemaining;
        int iStrikes;
        string sAuthor;
        Dictionary<string, object> dRow;

        bSkipRemaining = false;
        iStrikes = 0;
        while (true)
        {
            if (bStopLanes) break;
            dRow = dequeueRow(queueOpenLibrary);
            if (dRow != null)
            {
                if (!bSkipRemaining)
                {
                    try { openLibraryDetails(dRow); }
                    catch (Exception oException) { log("Open Library details failed: " + oException.Message); }
                    if (bLastFetchRateLimited) iStrikes = iStrikes + 1; else iStrikes = 0;
                    if (iStrikes >= iRateLimitStrikesMax)
                    {
                        bSkipRemaining = true;
                        aLaneSkipping[1] = true;
                        log("Open Library is rate limiting this address, so its remaining lookups are skipped and the run continues without them");
                        showTimedMessageBox("Open Library is rate limiting; skipping its remaining lookups this run");
                    }
                }
                Interlocked.Increment(ref iOpenLibraryDone);
                if (!bSkipRemaining) Thread.Sleep(iDelayApiMs + jitterMs(400));
                continue;
            }
            sAuthor = dequeueAuthor(queueAuthorOpenLibrary);
            if (sAuthor != null)
            {
                if (!bSkipRemaining)
                {
                    try { authorOpenLibrary(sAuthor); }
                    catch (Exception oException) { log("Open Library author lookup failed: " + oException.Message); }
                    if (bLastFetchRateLimited) iStrikes = iStrikes + 1; else iStrikes = 0;
                    if (iStrikes >= iRateLimitStrikesMax)
                    {
                        bSkipRemaining = true;
                        aLaneSkipping[1] = true;
                        log("Open Library is rate limiting this address, so its remaining lookups are skipped and the run continues without them");
                        showTimedMessageBox("Open Library is rate limiting; skipping its remaining lookups this run");
                    }
                    Thread.Sleep(iDelayApiMs + jitterMs(400));
                }
                continue;
            }
            if (bHarvestDone) break;
            Thread.Sleep(iLanePollMs);
        }
        log("The Open Library details lane finished");
    }

    // Reads the identifying values from a shared row under its lock.
    static void rowIdentity(Dictionary<string, object> dRow, out string sAsin, out string sTitle, out string sAuthor)
    {
        lock (dRow)
        {
            sAsin = catalogValue(dRow, "asin");
            sTitle = catalogValue(dRow, "title");
            sAuthor = "";
            foreach (string[] aPair in catalogLinks(dRow, "authors")) { sAuthor = aPair[0]; break; }
        }
    }

    // Fetches one title's details from Audible's catalog service.
    static void audibleDetails(Dictionary<string, object> dRow)
    {
        double nAverage;
        int iMinutes, iRatingsCount;
        string sAsin, sAuthor, sLadderText, sLanguage, sLengthText, sPublisher, sRatingText, sReleaseDate, sStepName, sTitle;
        Dictionary<string, object> dLadder, dProduct, dRating, dReply, dStep;
        List<string> lGenreLadders, lGenresTop;
        object[] aLadders, aSteps;

        rowIdentity(dRow, out sAsin, out sTitle, out sAuthor);
        dReply = fetchJson(sAudibleApiUrl + sAsin + "?response_groups=rating,product_attrs,product_extended_attrs,category_ladders");
        dProduct = dictAt(dReply, "product");
        if (dProduct == null) return;
        lock (dRow) { dRow["audibleChecked"] = true; }
        lGenreLadders = new List<string>();
        lGenresTop = new List<string>();
        if (dProduct.ContainsKey("category_ladders"))
        {
            aLadders = dProduct["category_ladders"] as object[];
            if (aLadders != null)
            {
                foreach (object oLadder in aLadders)
                {
                    dLadder = oLadder as Dictionary<string, object>;
                    aSteps = dLadder != null && dLadder.ContainsKey("ladder") ? dLadder["ladder"] as object[] : null;
                    if (aSteps == null || aSteps.Length == 0) continue;
                    sLadderText = "";
                    foreach (object oStep in aSteps)
                    {
                        dStep = oStep as Dictionary<string, object>;
                        sStepName = dStep != null && dStep.ContainsKey("name") ? Convert.ToString(dStep["name"]).Trim() : "";
                        if (sStepName == "") continue;
                        sLadderText = sLadderText == "" ? sStepName : sLadderText + " > " + sStepName;
                    }
                    if (sLadderText == "" || lGenreLadders.Contains(sLadderText)) continue;
                    lGenreLadders.Add(sLadderText);
                    sStepName = sLadderText.Contains(" > ") ? sLadderText.Substring(0, sLadderText.IndexOf(" > ")) : sLadderText;
                    if (!lGenresTop.Contains(sStepName)) lGenresTop.Add(sStepName);
                }
            }
        }
        nAverage = 0;
        iRatingsCount = 0;
        sRatingText = "";
        dRating = dictAt(dictAt(dProduct, "rating"), "overall_distribution");
        if (dRating != null && dRating.ContainsKey("average_rating"))
        {
            nAverage = Convert.ToDouble(dRating["average_rating"]);
            iRatingsCount = dRating.ContainsKey("num_ratings") ? Convert.ToInt32(dRating["num_ratings"]) : 0;
            if (iRatingsCount > 0) sRatingText = nAverage.ToString("0.0") + " out of 5 stars (" + iRatingsCount.ToString("N0") + " ratings)";
        }
        iMinutes = dProduct.ContainsKey("runtime_length_min") ? Convert.ToInt32(dProduct["runtime_length_min"]) : 0;
        sLengthText = lengthText(iMinutes);
        sReleaseDate = dProduct.ContainsKey("release_date") ? Convert.ToString(dProduct["release_date"]) : "";
        sPublisher = dProduct.ContainsKey("publisher_name") ? Convert.ToString(dProduct["publisher_name"]) : "";
        sLanguage = dProduct.ContainsKey("language") ? capitalizeFirst(Convert.ToString(dProduct["language"])) : "";
        lock (dRow)
        {
            if (sRatingText != "") { dRow["ratingValue"] = nAverage; dRow["ratingText"] = sRatingText; dRow["ratingsCount"] = iRatingsCount; }
            if (iMinutes > 0) { dRow["lengthMinutes"] = iMinutes; dRow["lengthText"] = sLengthText; }
            if (sReleaseDate != "") dRow["releaseDate"] = sReleaseDate;
            if (sPublisher != "") dRow["publisher"] = sPublisher;
            if (sLanguage != "") dRow["language"] = sLanguage;
            if (lGenreLadders.Count > 0) { dRow["genreLadders"] = lGenreLadders; dRow["genresTop"] = lGenresTop; }
        }
        log("Audible details gathered for: " + sTitle);
    }

    // Fetches one title's year of first publication from Open Library.
    static void openLibraryDetails(Dictionary<string, object> dRow)
    {
        string sAsin, sAuthor, sPublisher, sTitle, sYear;
        Dictionary<string, object> dDoc, dReply;

        rowIdentity(dRow, out sAsin, out sTitle, out sAuthor);
        dReply = fetchJson(sOpenLibraryUrl + "?title=" + Uri.EscapeDataString(sTitle) + (sAuthor == "" ? "" : "&author=" + Uri.EscapeDataString(sAuthor)) + "&fields=first_publish_year,publisher&limit=1");
        if (dReply != null) lock (dRow) { dRow["openLibraryChecked"] = true; }
        dDoc = firstAt(dReply, "docs");
        if (dDoc == null) return;
        // Open Library reports the publishers of the print editions.  The
        // first one stands in for the audiobook publisher, clearly labeled,
        // only when Audible's catalog service offered none; a real Audible
        // publisher always overwrites this approximation.
        foreach (string sName in catalogStrings(dDoc, "publisher"))
        {
            sPublisher = sName.Trim();
            if (sPublisher == "") continue;
            lock (dRow) { if (!dRow.ContainsKey("publisher")) dRow["publisher"] = sPublisher + " (Open Library)"; }
            break;
        }
        if (!dDoc.ContainsKey("first_publish_year")) return;
        sYear = Convert.ToString(dDoc["first_publish_year"]);
        lock (dRow) { dRow["firstPublished"] = sYear; }
    }

    // The Wikipedia lane: a page for each book when one exists, then author
    // biographies, all through one politely paced connection.
    static void wikipediaLane()
    {
        bool bSkipRemaining;
        int iStrikes;
        string sAuthor;
        Dictionary<string, object> dRow;

        bSkipRemaining = false;
        iStrikes = 0;
        while (true)
        {
            if (bStopLanes) break;
            dRow = dequeueRow(queueWikipedia);
            if (dRow != null)
            {
                if (!bSkipRemaining)
                {
                    try { wikipediaDetails(dRow); }
                    catch (Exception oException) { log("Wikipedia details failed: " + oException.Message); }
                    if (bLastFetchRateLimited) iStrikes = iStrikes + 1; else iStrikes = 0;
                    if (iStrikes >= iRateLimitStrikesMax)
                    {
                        bSkipRemaining = true;
                        aLaneSkipping[2] = true;
                        log("Wikipedia is rate limiting this address, so its remaining lookups are skipped and the run continues without them");
                        showTimedMessageBox("Wikipedia is rate limiting; skipping its remaining lookups this run");
                    }
                }
                Interlocked.Increment(ref iWikipediaDone);
                if (!bSkipRemaining) Thread.Sleep(iDelayWikipediaMs + jitterMs(600));
                continue;
            }
            sAuthor = dequeueAuthor(queueAuthorWikipedia);
            if (sAuthor != null)
            {
                if (!bSkipRemaining)
                {
                    try { authorWikipedia(sAuthor); }
                    catch (Exception oException) { log("Wikipedia author lookup failed: " + oException.Message); }
                    if (bLastFetchRateLimited) iStrikes = iStrikes + 1; else iStrikes = 0;
                    if (iStrikes >= iRateLimitStrikesMax)
                    {
                        bSkipRemaining = true;
                        aLaneSkipping[2] = true;
                        log("Wikipedia is rate limiting this address, so its remaining lookups are skipped and the run continues without them");
                        showTimedMessageBox("Wikipedia is rate limiting; skipping its remaining lookups this run");
                    }
                    Thread.Sleep(iDelayWikipediaMs + jitterMs(600));
                }
                Interlocked.Increment(ref iAuthorsDone);
                continue;
            }
            if (bHarvestDone) break;
            Thread.Sleep(iLanePollMs);
        }
        log("The Wikipedia details lane finished");
    }

    // Searches Wikipedia for a page whose title matches the book's title,
    // accepting parenthetical disambiguators such as (novel) but never a
    // disambiguation page, so only a confident match produces the field.
    static void wikipediaDetails(Dictionary<string, object> dRow)
    {
        string sAsin, sAuthor, sPageTitle, sTitle;
        Dictionary<string, object> dHit, dReply;

        rowIdentity(dRow, out sAsin, out sTitle, out sAuthor);
        dReply = fetchJson(sWikipediaApiUrl + Uri.EscapeDataString("intitle:\"" + sTitle + "\" " + sAuthor));
        if (dReply != null) lock (dRow) { dRow["wikipediaChecked"] = true; }
        dHit = firstAt(dictAt(dReply, "query"), "search");
        if (dHit == null || !dHit.ContainsKey("title")) return;
        sPageTitle = Convert.ToString(dHit["title"]).Trim();
        if (!wikipediaTitleMatches(sTitle, sPageTitle)) return;
        lock (dRow)
        {
            dRow["wikipediaTitle"] = sPageTitle;
            dRow["wikipediaUrl"] = sWikipediaPageUrl + Uri.EscapeDataString(sPageTitle.Replace(' ', '_'));
        }
    }

    // Returns true when a Wikipedia page title confidently names the book:
    // the titles match after normalization, allowing the page a trailing
    // parenthetical such as (novel) or (book), but never (disambiguation).
    static bool wikipediaTitleMatches(string sBookTitle, string sPageTitle)
    {
        int iParen;
        string sBare;

        if (sPageTitle.ToLower().EndsWith("(disambiguation)")) return false;
        sBare = sPageTitle;
        iParen = sPageTitle.LastIndexOf(" (", StringComparison.Ordinal);
        if (iParen > 0 && sPageTitle.EndsWith(")")) sBare = sPageTitle.Substring(0, iParen);
        return normalizedTitle(sBare) == normalizedTitle(sBookTitle) && normalizedTitle(sBare) != "";
    }

    // Fetches an author's Wikipedia summary and keeps it only when the page
    // is a standard article whose title matches the author's name, whose
    // extract is substantial rather than a stub, and whose text reads like a
    // biography of a writer, guarding against dummy pages and namesakes.
    static void authorWikipedia(string sAuthor)
    {
        string sExtract, sPageTitle, sType;
        Dictionary<string, object> dReply;

        dReply = fetchJson(sWikipediaSummaryUrl + Uri.EscapeDataString(sAuthor.Replace(' ', '_')));
        if (dReply == null && bLastFetchRateLimited) return;
        lock (oAuthorLock) { setAuthorsCheckedWiki.Add(sAuthor); }
        if (dReply == null) return;
        sType = dReply.ContainsKey("type") ? Convert.ToString(dReply["type"]) : "";
        if (sType != "standard") return;
        sPageTitle = dReply.ContainsKey("title") ? Convert.ToString(dReply["title"]).Trim() : "";
        if (normalizedTitle(sPageTitle) != normalizedTitle(sAuthor) || normalizedTitle(sPageTitle) == "") return;
        sExtract = dReply.ContainsKey("extract") ? collapseWhitespace(Convert.ToString(dReply["extract"])) : "";
        if (sExtract.Length < 150) return;
        if (!looksLikeWriterBio(sExtract)) return;
        lock (oAuthorLock)
        {
            dAuthorWikiBio[sAuthor] = sExtract;
            dAuthorWikiUrl[sAuthor] = sWikipediaPageUrl + Uri.EscapeDataString(sPageTitle.Replace(' ', '_'));
        }
    }

    // Returns true when a biography extract plausibly concerns a writer.
    static bool looksLikeWriterBio(string sExtract)
    {
        string sLower;

        sLower = sExtract.ToLower();
        return sLower.Contains("author") || sLower.Contains("writer") || sLower.Contains("novelist") || sLower.Contains("historian") ||
            sLower.Contains("journalist") || sLower.Contains("essayist") || sLower.Contains("poet") || sLower.Contains("biographer") ||
            sLower.Contains("professor") || sLower.Contains("scholar") || sLower.Contains("wrote") || sLower.Contains("book");
    }

    // Fetches an author's biography from Open Library: an author search
    // whose top result must match the name exactly, then the author record,
    // whose bio field may be a plain string or a wrapped value.
    static void authorOpenLibrary(string sAuthor)
    {
        string sBio, sKey, sName;
        Dictionary<string, object> dAuthorDoc, dRecord;
        object oBio;

        dAuthorDoc = firstAt(fetchJson(sOpenLibraryAuthorSearchUrl + Uri.EscapeDataString(sAuthor)), "docs");
        if (dAuthorDoc == null && bLastFetchRateLimited) return;
        lock (oAuthorLock) { setAuthorsCheckedOl.Add(sAuthor); }
        if (dAuthorDoc == null) return;
        sName = dAuthorDoc.ContainsKey("name") ? Convert.ToString(dAuthorDoc["name"]).Trim() : "";
        if (normalizedTitle(sName) != normalizedTitle(sAuthor) || normalizedTitle(sName) == "") return;
        sKey = dAuthorDoc.ContainsKey("key") ? Convert.ToString(dAuthorDoc["key"]).Replace("/authors/", "").Trim() : "";
        if (sKey == "") return;
        Thread.Sleep(iDelayApiMs + jitterMs(400));
        dRecord = fetchJson(sOpenLibraryAuthorUrl + sKey + ".json");
        if (dRecord == null || !dRecord.ContainsKey("bio")) return;
        oBio = dRecord["bio"];
        sBio = oBio is Dictionary<string, object> ? Convert.ToString(((Dictionary<string, object>) oBio).ContainsKey("value") ? ((Dictionary<string, object>) oBio)["value"] : "") : Convert.ToString(oBio);
        sBio = collapseWhitespace(sBio);
        if (sBio.Length < 80 || !meaningfulValue(sBio)) return;
        lock (oAuthorLock) { dAuthorOlBio[sAuthor] = sBio; }
    }

    // Collapses runs of whitespace, including newlines, to single spaces.
    static string collapseWhitespace(string sText)
    {
        StringBuilder sbOut;

        sbOut = new StringBuilder();
        foreach (char chOne in sText)
        {
            if (char.IsWhiteSpace(chOne)) { if (sbOut.Length > 0 && sbOut[sbOut.Length - 1] != ' ') sbOut.Append(' '); }
            else sbOut.Append(chOne);
        }
        return sbOut.ToString().Trim();
    }

    // Lowercases a title and collapses everything but letters and digits, so
    // punctuation and spacing differences never block a match.
    static string normalizedTitle(string sText)
    {
        StringBuilder sbOut;

        sbOut = new StringBuilder();
        foreach (char chOne in sText.ToLower()) { if (char.IsLetterOrDigit(chOne)) sbOut.Append(chOne); }
        return sbOut.ToString();
    }

    // Returns a catalog field that holds a plain list of strings.
    static List<string> catalogStrings(Dictionary<string, object> dRow, string sKey)
    {
        object[] aValues;
        List<string> lValues;

        if (!dRow.ContainsKey(sKey)) return new List<string>();
        lValues = dRow[sKey] as List<string>;
        if (lValues != null) return lValues;
        aValues = dRow[sKey] as object[];
        lValues = new List<string>();
        if (aValues != null) foreach (object oOne in aValues) lValues.Add(Convert.ToString(oOne));
        return lValues;
    }

    // Places a rating average into the named bucket used by the rating appendix.
    static string ratingBucket(double nAverage)
    {
        if (nAverage >= 4.5) return "4.5 stars and up";
        if (nAverage >= 4.0) return "4.0 to 4.4 stars";
        if (nAverage >= 3.5) return "3.5 to 3.9 stars";
        return "Below 3.5 stars";
    }

    // Returns the string value of a catalog field, or an empty string.
    static string catalogValue(Dictionary<string, object> dRow, string sKey)
    {
        if (!dRow.ContainsKey(sKey) || dRow[sKey] == null) return "";
        return Convert.ToString(dRow[sKey]);
    }

    // Returns a catalog field's list of name-and-url pairs.
    static List<string[]> catalogLinks(Dictionary<string, object> dRow, string sKey)
    {
        Dictionary<string, object> dPair;
        List<string[]> lPairs;

        lPairs = new List<string[]>();
        if (!dRow.ContainsKey(sKey) || dRow[sKey] == null) return lPairs;
        foreach (object oPair in (IEnumerable) dRow[sKey])
        {
            dPair = (Dictionary<string, object>) oPair;
            lPairs.Add(new string[] { cleanContributorName(Convert.ToString(dPair["name"])), cleanAudibleUrl(Convert.ToString(dPair["url"])) });
        }
        return lPairs;
    }

    // Strips a trailing role annotation such as "- foreword" from a
    // contributor's name, so Alan Parsons - foreword indexes, links, and
    // looks up as Alan Parsons.
    static string cleanContributorName(string sName)
    {
        int iDash;
        string sRole;

        sName = sName.Trim();
        iDash = sName.LastIndexOfAny(new char[] { '-', (char) 8211, (char) 8212 });
        if (iDash <= 0) return sName;
        sRole = sName.Substring(iDash + 1).Trim().ToLower();
        if (sRole == "foreword" || sRole == "introduction" || sRole == "afterword" || sRole == "preface" || sRole == "editor" || sRole == "translator" ||
            sRole == "narrator" || sRole == "contributor" || sRole == "illustrator" || sRole == "adaptation" || sRole == "adapter" || sRole == "compiler" || sRole == "abridged")
            return sName.Substring(0, iDash).Trim().TrimEnd('-', (char) 8211, (char) 8212).Trim();
        return sName;
    }

    // Strips the tracking query from an Audible url, keeping only the
    // searchNarrator parameter when that is what the link is for.
    static string cleanAudibleUrl(string sUrl)
    {
        int iAmp, iPos;

        iPos = sUrl.IndexOf("searchNarrator=");
        if (iPos >= 0)
        {
            iAmp = sUrl.IndexOf("&", iPos);
            return "https://www.audible.com/search?" + (iAmp >= 0 ? sUrl.Substring(iPos, iAmp - iPos) : sUrl.Substring(iPos));
        }
        iPos = sUrl.IndexOf("?");
        return iPos >= 0 ? sUrl.Substring(0, iPos) : sUrl;
    }

    // Escapes text for HTML output.
    static string htmlText(string sText)
    {
        return WebUtility.HtmlEncode(sText);
    }

    // Escapes link text for Markdown output.
    static string mdText(string sText)
    {
        return sText.Replace("[", "(").Replace("]", ")");
    }

    // Renders a list of name-and-url pairs as comma-separated HTML links.
    static string linksHtml(List<string[]> lPairs)
    {
        StringBuilder sbOut;

        sbOut = new StringBuilder();
        foreach (string[] aPair in lPairs)
        {
            if (sbOut.Length > 0) sbOut.Append(", ");
            sbOut.Append("<a href=\"" + htmlText(aPair[1]) + "\">" + htmlText(aPair[0]) + "</a>");
        }
        return sbOut.ToString();
    }

    // Renders a list of name-and-url pairs as comma-separated Markdown links.
    static string linksMd(List<string[]> lPairs)
    {
        StringBuilder sbOut;

        sbOut = new StringBuilder();
        foreach (string[] aPair in lPairs)
        {
            if (sbOut.Length > 0) sbOut.Append(", ");
            sbOut.Append("[" + mdText(aPair[0]) + "](" + aPair[1] + ")");
        }
        return sbOut.ToString();
    }

    // Adds a title under a category value in an appendix index.
    static void addToIndex(Dictionary<string, List<string[]>> dIndex, string sCategory, string sTitle, string sAsin)
    {
        if (sCategory == "" || sTitle == "") return;
        if (!dIndex.ContainsKey(sCategory)) dIndex[sCategory] = new List<string[]>();
        dIndex[sCategory].Add(new string[] { sTitle, sAsin });
    }

    // Writes Audible_Library.xlsx: one worksheet holding one contiguous
    // region that starts at A1, following the conventions of his xlFormat
    // and xlHeaders tools for screen reader users: a single bold header row
    // of unique names, a workbook-level name ColumnTitle01 on the region's
    // title cell so JAWS announces column headers while navigating, column
    // widths fitted to content but capped at 40 with wrapping on capped
    // columns, and no merged cells.  Unlike the documents, every column is
    // present for every title, and a value that is not known is left as a
    // truly empty cell, which a screen reader reads as blank.
    // Writes bookFido.xlsx: one workbook holding one sheet per library
    // source, each sheet following the screen-reader conventions, with a
    // sheet-scoped ColumnTitle01 name on each sheet's A1 so JAWS announces
    // column headers everywhere in the workbook.
    static void writeWorkbook(string sDownloadDir)
    {
        string sPath;

        sPath = Path.Combine(sDownloadDir, "bookFido.xlsx");
        using (ExcelPackage oPackage = new ExcelPackage())
        {
            fillAudibleSheet(oPackage);
            if (lKindleCatalog.Count > 0) fillKindleSheet(oPackage);
            if (lGoodreadsCatalog.Count > 0) fillGoodreadsSheet(oPackage);
            if (lBookshareCatalog.Count > 0) fillBookshareSheet(oPackage);
            if (File.Exists(sPath)) File.Delete(sPath);
            File.WriteAllBytes(sPath, oPackage.GetAsByteArray());
        }
        log("The spreadsheet workbook was saved as " + sPath + " with one sheet per library");
    }

    static void fillAudibleSheet(ExcelPackage oPackage)
    {
        int iCol, iRow, iWidth;
        int[] aWidths;
        string[] aHeaders;
        List<Dictionary<string, object>> lSorted;
        List<string[]> lPairs;

        aHeaders = new string[] { "Title", "ASIN", "Audible link", "By", "First published", "Genres", "Language", "Length", "Length in minutes", "Narrated by", "Progress", "Publisher", "Rating", "Rating average", "Ratings count", "Release date", "Series", "Summary", "Wikipedia" };
        lSorted = new List<Dictionary<string, object>>(lCatalog);
        lSorted.Sort(compareCatalogRowsByTitle);
        ExcelWorksheet oSheet = oPackage.Workbook.Worksheets.Add("Audible");
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1) oSheet.Cells[1, iCol].Value = aHeaders[iCol - 1];
        oSheet.Cells[1, 1, 1, aHeaders.Length].Style.Font.Bold = true;
        iRow = 1;
        foreach (Dictionary<string, object> dRow in lSorted)
        {
            iRow = iRow + 1;
            Monitor.Enter(dRow);
            oSheet.Cells[iRow, 1].Value = catalogValue(dRow, "title");
            oSheet.Cells[iRow, 2].Value = catalogValue(dRow, "asin");
            oSheet.Cells[iRow, 3].Value = catalogValue(dRow, "url");
            lPairs = catalogLinks(dRow, "authors");
            oSheet.Cells[iRow, 4].Value = joinPairNames(lPairs);
            oSheet.Cells[iRow, 5].Value = catalogValue(dRow, "firstPublished");
            oSheet.Cells[iRow, 6].Value = string.Join("; ", catalogStrings(dRow, "genreLadders").ToArray());
            oSheet.Cells[iRow, 7].Value = catalogValue(dRow, "language");
            oSheet.Cells[iRow, 8].Value = catalogValue(dRow, "lengthText");
            if (dRow.ContainsKey("lengthMinutes")) oSheet.Cells[iRow, 9].Value = Convert.ToInt32(dRow["lengthMinutes"]);
            lPairs = catalogLinks(dRow, "narrators");
            oSheet.Cells[iRow, 10].Value = joinPairNames(lPairs);
            oSheet.Cells[iRow, 11].Value = Convert.ToBoolean(dRow["finished"]) ? "Finished" : catalogValue(dRow, "timeLeft");
            oSheet.Cells[iRow, 12].Value = catalogValue(dRow, "publisher");
            oSheet.Cells[iRow, 13].Value = catalogValue(dRow, "ratingText");
            if (dRow.ContainsKey("ratingValue")) oSheet.Cells[iRow, 14].Value = Convert.ToDouble(dRow["ratingValue"]);
            if (dRow.ContainsKey("ratingsCount")) oSheet.Cells[iRow, 15].Value = Convert.ToInt32(dRow["ratingsCount"]);
            oSheet.Cells[iRow, 16].Value = catalogValue(dRow, "releaseDate");
            lPairs = catalogLinks(dRow, "series");
            oSheet.Cells[iRow, 17].Value = joinPairNames(lPairs);
            oSheet.Cells[iRow, 18].Value = catalogValue(dRow, "summary");
            oSheet.Cells[iRow, 19].Value = catalogValue(dRow, "wikipediaUrl");
            Monitor.Exit(dRow);
        }
        aWidths = new int[aHeaders.Length];
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1)
        {
            aWidths[iCol - 1] = aHeaders[iCol - 1].Length;
            for (iRow = 2; iRow <= lSorted.Count + 1; iRow = iRow + 1)
            {
                iWidth = oSheet.Cells[iRow, iCol].Value == null ? 0 : Convert.ToString(oSheet.Cells[iRow, iCol].Value).Length;
                if (iWidth > aWidths[iCol - 1]) aWidths[iCol - 1] = iWidth;
            }
            if (aWidths[iCol - 1] > 40)
            {
                aWidths[iCol - 1] = 40;
                oSheet.Column(iCol).Style.WrapText = true;
            }
            oSheet.Column(iCol).Width = aWidths[iCol - 1] + 2;
        }
        oSheet.View.FreezePanes(2, 1);
        oSheet.Names.Add("ColumnTitle01", oSheet.Cells[1, 1]);
    }

    // Joins the names of link pairs for a spreadsheet cell.
    static string joinPairNames(List<string[]> lPairs)
    {
        List<string> lNames;

        lNames = new List<string>();
        foreach (string[] aPair in lPairs) lNames.Add(aPair[0]);
        return string.Join("; ", lNames.ToArray());
    }

    // Orders catalog rows alphabetically by title for the spreadsheet.
    static int compareCatalogRowsByTitle(Dictionary<string, object> dFirst, Dictionary<string, object> dSecond)
    {
        return string.Compare(titleSortKey(catalogValue(dFirst, "title")), titleSortKey(catalogValue(dSecond, "title")), StringComparison.OrdinalIgnoreCase);
    }

    // The sort key for a title: case folds, and a leading A, An, or The is
    // ignored, so The Odyssey files under O.
    static string titleSortKey(string sTitle)
    {
        string sKey;

        sKey = sTitle.Trim();
        if (sKey.StartsWith("The ", StringComparison.OrdinalIgnoreCase)) sKey = sKey.Substring(4);
        else if (sKey.StartsWith("An ", StringComparison.OrdinalIgnoreCase)) sKey = sKey.Substring(3);
        else if (sKey.StartsWith("A ", StringComparison.OrdinalIgnoreCase)) sKey = sKey.Substring(2);
        return sKey.Trim();
    }

    // A title decorated with its publication year in parentheses, using
    // the year of first publication when Open Library supplied one and the
    // audio release year otherwise; a title with no known year stands
    // alone.  Callers hold the row's lock.
    static string titleWithYear(Dictionary<string, object> dRow)
    {
        string sTitle, sYear;

        sTitle = catalogValue(dRow, "title");
        sYear = catalogValue(dRow, "firstPublished");
        if (!isYear(sYear)) sYear = yearFrom(catalogValue(dRow, "releaseDate"));
        if (sYear == "") sYear = yearFrom(grField(dRow, "date pub"));
        if (sYear == "") sYear = yearFrom(grField(dRow, "date pub edition"));
        return sYear == "" ? sTitle : sTitle + " (" + sYear + ")";
    }

    static bool isYear(string sText)
    {
        if (sText == null || sText.Length != 4) return false;
        foreach (char chOne in sText) { if (!char.IsDigit(chOne)) return false; }
        return true;
    }

    // The last plausible four-digit year inside a date or free text.
    static string yearFrom(string sText)
    {
        string sFound;

        sFound = "";
        foreach (System.Text.RegularExpressions.Match oMatch in System.Text.RegularExpressions.Regex.Matches(sText == null ? "" : sText, "\\b(1[5-9][0-9][0-9]|20[0-9][0-9])\\b")) sFound = oMatch.Groups[1].Value;
        return sFound;
    }

    // The edition a book represents, from its library and its format field:
    // Kindle, Audible, Print, EPUB, or Generic.
    static string editionText(Dictionary<string, object> dRow, string sLibrary)
    {
        string sFormat;

        if (sLibrary == "Audible") return "Audible";
        if (sLibrary == "Kindle") return "Kindle";
        if (sLibrary == "Goodreads")
        {
            sFormat = grField(dRow, "format").ToLowerInvariant();
            if (sFormat.Contains("kindle")) return "Kindle";
            if (sFormat.Contains("audible") || sFormat.Contains("audio")) return "Audible";
            if (sFormat.Contains("epub") || sFormat.Contains("ebook")) return "EPUB";
            if (sFormat.Contains("paperback") || sFormat.Contains("hardcover") || sFormat.Contains("hardback") || sFormat.Contains("print") || sFormat.Contains("mass market") || sFormat.Contains("library binding") || sFormat.Contains("board book") || sFormat.Contains("spiral")) return "Print";
            return "Generic";
        }
        sFormat = bsField(dRow, "Format").ToLowerInvariant();
        if (sFormat.Contains("epub")) return "EPUB";
        return "Generic";
    }

    // The sort key for a person: the surname first, case folded, with the
    // full name as the tiebreaker; a name that itself begins with The, such
    // as The Great Courses, is an organization and sorts as written.
    static string personSortKey(string sName)
    {
        int iSpace;
        string sKey;

        sKey = sName.Trim();
        if (sKey.StartsWith("The ", StringComparison.OrdinalIgnoreCase)) return sKey;
        iSpace = sKey.LastIndexOf(' ');
        return iSpace < 0 ? sKey : sKey.Substring(iSpace + 1) + " " + sKey;
    }

    // Orders person names by surname, case insensitively.
    static int comparePersonNames(string sFirst, string sSecond)
    {
        return string.Compare(personSortKey(sFirst), personSortKey(sSecond), StringComparison.OrdinalIgnoreCase);
    }

    // Writes the library-highlights paragraph: the most represented authors
    // and narrator, the longest listen, the oldest work, and the most widely
    // rated title.  Every piece is guarded, so nothing appears when the data
    // behind it is absent.
    static void appendHighlights(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, List<string[]>> dByAuthor, Dictionary<string, List<string[]>> dByNarrator, string sLongestTitle, int iLongestMinutes, string sOldestTitle, int iOldestYear, string sMostRatedTitle, int iMostRatings)
    {
        int iShown;
        string sHighlights;
        List<string[]> lAuthorRows, lNarratorRows;

        sHighlights = "";
        lAuthorRows = countRows(dByAuthor, null);
        if (lAuthorRows.Count > 0)
        {
            sHighlights = "The most represented authors are ";
            iShown = 0;
            foreach (string[] aRow in lAuthorRows)
            {
                if (iShown == 3 || int.Parse(aRow[1]) < 2) break;
                sHighlights = sHighlights + (iShown == 0 ? "" : ", ") + aRow[0] + " (" + aRow[1] + " titles)";
                iShown = iShown + 1;
            }
            sHighlights = iShown > 0 ? sHighlights + ".  " : "";
        }
        lNarratorRows = countRows(dByNarrator, null);
        if (lNarratorRows.Count > 0 && int.Parse(lNarratorRows[0][1]) >= 2) sHighlights = sHighlights + "The most frequent narrator is " + lNarratorRows[0][0] + ", heard on " + lNarratorRows[0][1] + " titles.  ";
        if (sLongestTitle != "" && iLongestMinutes > 0) sHighlights = sHighlights + "The longest listen is " + sLongestTitle + " at " + lengthText(iLongestMinutes) + ".  ";
        if (sOldestTitle != "" && iOldestYear > 0) sHighlights = sHighlights + "The oldest work is " + sOldestTitle + ", first published in " + iOldestYear + ".  ";
        if (sMostRatedTitle != "" && iMostRatings > 0) sHighlights = sHighlights + "The most widely rated title is " + sMostRatedTitle + ", with " + iMostRatings.ToString("N0") + " Audible ratings.";
        sHighlights = sHighlights.Trim();
        if (sHighlights == "") return;
        sbHtml.Append("<p>" + htmlText(sHighlights) + "</p>\r\n");
        sbMd.Append(sHighlights + "\r\n\r\n");
    }

    // Writes Appendix I: every author in alphabetical order with the count
    // of their titles, the hours they account for, and links to each title.
    // Biographies are omitted because none of the services already queried
    // return one without an additional per-author request.
    static void appendAuthorsAppendix(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, List<string[]>> dByAuthor, Dictionary<string, int> dAuthorMinutes)
    {
        appendAuthorsAppendix(sbHtml, sbMd, dByAuthor, dAuthorMinutes, "appendix-h", "Appendix H: About the Authors", "Every author in the library, with the number of titles and the listening time they account for, and a biography when a reliable one was found.");
    }

    static void appendAuthorsAppendix(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, List<string[]>> dByAuthor, Dictionary<string, int> dAuthorMinutes, string sId, string sHeading, string sLead)
    {
        int iMinutes;
        List<string> lAuthors;

        sbHtml.Append("<h2 id=\"" + sId + "\">" + htmlText(sHeading) + "</h2>\r\n");
        sbMd.Append("## " + sHeading + " {#" + sId + "}\r\n\r\n");
        sbHtml.Append("<p>" + htmlText(sLead) + "</p>\r\n");
        sbMd.Append(sLead + "\r\n\r\n");
        lAuthors = new List<string>(dByAuthor.Keys);
        lAuthors.Sort(comparePersonNames);
        Monitor.Enter(oAuthorLock);
        foreach (string sAuthor in lAuthors)
        {
            iMinutes = dAuthorMinutes.ContainsKey(sAuthor) ? dAuthorMinutes[sAuthor] : 0;
            dByAuthor[sAuthor].Sort(compareTitlePairs);
            sbHtml.Append("<h3>" + htmlText(sAuthor) + "</h3>\r\n<p>" + dByAuthor[sAuthor].Count + (dByAuthor[sAuthor].Count == 1 ? " title" : " titles") + (iMinutes > 0 ? ", " + lengthText(iMinutes) : "") + "</p>\r\n");
            sbMd.Append("### " + mdText(sAuthor) + "\r\n\r\n" + dByAuthor[sAuthor].Count + (dByAuthor[sAuthor].Count == 1 ? " title" : " titles") + (iMinutes > 0 ? ", " + lengthText(iMinutes) : "") + "\r\n\r\n");
            if (dAuthorWikiBio.ContainsKey(sAuthor))
            {
                sbHtml.Append("<p>" + htmlText(dAuthorWikiBio[sAuthor] + " (Wikipedia)") + "</p>\r\n");
                sbMd.Append(dAuthorWikiBio[sAuthor] + " (Wikipedia)\r\n\r\n");
            }
            if (dAuthorOlBio.ContainsKey(sAuthor) && (!dAuthorWikiBio.ContainsKey(sAuthor) || normalizedTitle(dAuthorOlBio[sAuthor]) != normalizedTitle(dAuthorWikiBio[sAuthor])))
            {
                sbHtml.Append("<p>" + htmlText(dAuthorOlBio[sAuthor] + " (Open Library)") + "</p>\r\n");
                sbMd.Append(dAuthorOlBio[sAuthor] + " (Open Library)\r\n\r\n");
            }
            if (dAuthorWikiUrl.ContainsKey(sAuthor))
            {
                sbHtml.Append("<p>Wikipedia: <a href=\"" + htmlText(dAuthorWikiUrl[sAuthor]) + "\">" + htmlText(sAuthor) + "</a></p>\r\n");
                sbMd.Append("Wikipedia: [" + mdText(sAuthor) + "](" + dAuthorWikiUrl[sAuthor] + ")\r\n\r\n");
            }
            sbHtml.Append("<ul>\r\n");
            foreach (string[] aEntry in dByAuthor[sAuthor])
            {
                sbHtml.Append("<li><a href=\"#" + aEntry[1] + "\">" + htmlText(aEntry[0]) + "</a></li>\r\n");
                sbMd.Append("- [" + mdText(aEntry[0]) + "](#" + aEntry[1].ToLower() + ")\r\n");
            }
            sbHtml.Append("</ul>\r\n");
            sbMd.Append("\r\n");
        }
        Monitor.Exit(oAuthorLock);
    }

    // Builds label-and-count rows from an index.  With a fixed order, only
    // those labels appear in that order; otherwise rows sort by descending
    // count, then alphabetically.
    static List<string[]> countRows(Dictionary<string, List<string[]>> dIndex, string[] aFixedOrder)
    {
        List<string> lLabels;
        List<string[]> lRows;

        lRows = new List<string[]>();
        if (aFixedOrder != null)
        {
            foreach (string sLabel in aFixedOrder) { if (dIndex.ContainsKey(sLabel)) lRows.Add(new string[] { sLabel, dIndex[sLabel].Count.ToString() }); }
            return lRows;
        }
        lLabels = new List<string>(dIndex.Keys);
        lLabels.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string sLabel in lLabels) lRows.Add(new string[] { sLabel, dIndex[sLabel].Count.ToString() });
        lRows.Sort(compareCountRows);
        return lRows;
    }

    // Orders count rows by descending count, then label.
    static int compareCountRows(string[] aFirst, string[] aSecond)
    {
        int iFirst, iSecond;

        iFirst = int.Parse(aFirst[1]);
        iSecond = int.Parse(aSecond[1]);
        if (iFirst != iSecond) return iSecond.CompareTo(iFirst);
        return string.Compare(aFirst[0], aSecond[0], StringComparison.OrdinalIgnoreCase);
    }

    // Writes one tabulation to both outputs: an accessible HTML table with a
    // caption and column headers, and a Markdown pipe table.
    static void appendCountTable(StringBuilder sbHtml, StringBuilder sbMd, string sCaption, string sColumnName, List<string[]> lRows)
    {
        if (lRows.Count == 0) return;
        sbHtml.Append("<table>\r\n<caption>" + htmlText(sCaption) + "</caption>\r\n<tr><th scope=\"col\">" + htmlText(sColumnName) + "</th><th scope=\"col\">Titles</th></tr>\r\n");
        sbMd.Append(sCaption + "\r\n\r\n| " + sColumnName + " | Titles |\r\n| --- | --- |\r\n");
        foreach (string[] aRow in lRows)
        {
            sbHtml.Append("<tr><td>" + htmlText(aRow[0]) + "</td><td>" + aRow[1] + "</td></tr>\r\n");
            sbMd.Append("| " + mdText(aRow[0]) + " | " + aRow[1] + " |\r\n");
        }
        sbHtml.Append("</table>\r\n");
        sbMd.Append("\r\n");
    }

    // Writes one appendix to both outputs: categories in alphabetical order,
    // and under each an alphabetical list of titles linking to the title
    // headings.
    static void appendAppendix(StringBuilder sbHtml, StringBuilder sbMd, string sId, string sHeading, Dictionary<string, List<string[]>> dIndex)
    {
        appendAppendix(sbHtml, sbMd, sId, sHeading, dIndex, false);
    }

    static void appendAppendix(StringBuilder sbHtml, StringBuilder sbMd, string sId, string sHeading, Dictionary<string, List<string[]>> dIndex, bool bPersonNames)
    {
        List<string> lCategories;

        sbHtml.Append("<h2 id=\"" + sId + "\">" + htmlText(sHeading) + "</h2>\r\n");
        sbMd.Append("## " + sHeading + " {#" + sId + "}\r\n\r\n");
        lCategories = new List<string>(dIndex.Keys);
        if (bPersonNames) lCategories.Sort(comparePersonNames); else lCategories.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string sCategory in lCategories)
        {
            dIndex[sCategory].Sort(compareTitlePairs);
            sbHtml.Append("<h3>" + htmlText(sCategory) + "</h3>\r\n<ul>\r\n");
            sbMd.Append("### " + sCategory + "\r\n\r\n");
            foreach (string[] aPair in dIndex[sCategory])
            {
                sbHtml.Append("<li><a href=\"#" + aPair[1] + "\">" + htmlText(aPair[0]) + "</a></li>\r\n");
                sbMd.Append("- [" + mdText(aPair[0]) + "](#" + aPair[1].ToLower() + ")\r\n");
            }
            sbHtml.Append("</ul>\r\n");
            sbMd.Append("\r\n");
        }
    }

    // Orders title-and-asin pairs alphabetically by title.
    static int compareTitlePairs(string[] aFirst, string[] aSecond)
    {
        return string.Compare(titleSortKey(aFirst[0]), titleSortKey(aSecond[0]), StringComparison.OrdinalIgnoreCase);
    }

    // Writes one field line to both outputs, skipping empty values.
    // Returns true only for values worth showing.  Empty strings, lone
    // punctuation, and the dummy placeholders services sometimes return are
    // all rejected, so no field ever states that nothing is known.
    static bool meaningfulValue(string sValue)
    {
        string sTrimmed;

        if (sValue == null) return false;
        sTrimmed = sValue.Trim().Trim('.', '-', '_', '*').Trim().ToLower();
        if (sTrimmed.Length < 2) return false;
        if (sTrimmed == "n/a" || sTrimmed == "na" || sTrimmed == "none" || sTrimmed == "null" || sTrimmed == "unknown" || sTrimmed == "not available" || sTrimmed == "tbd" || sTrimmed == "undefined" || sTrimmed == "no summary available" || sTrimmed == "description coming soon" || sTrimmed == "coming soon" || sTrimmed == "0000-00-00") return false;
        return true;
    }

    static void appendField(StringBuilder sbHtml, StringBuilder sbMd, string sLabel, string sHtmlValue, string sMdValue)
    {
        if (!meaningfulValue(sMdValue)) return;
        sbHtml.Append("<p>" + htmlText(sLabel) + ": " + sHtmlValue + "</p>\r\n");
        sbMd.Append(sLabel + ": " + sMdValue + "\r\n\r\n");
    }

    // Builds Audible_Library.htm and Audible_Library.md in the download
    // folder from the harvested catalog, and returns the path of the htm
    // file, or an empty string when there was nothing to write.
    // One book's full entry, with every field its library offers, used by
    // both the library's own document and the consolidated catalog.
    static void appendAudibleEntry(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, object> dRow)
    {
        string sAsin, sProgress, sTitle, sUrl;
        List<string[]> lPairs;

        Monitor.Enter(dRow);
        sAsin = catalogValue(dRow, "asin");
        sTitle = catalogValue(dRow, "title");
        sUrl = cleanAudibleUrl(catalogValue(dRow, "url"));
        if (sTitle == "") { Monitor.Exit(dRow); return; }
        sbHtml.Append("<h2 id=\"" + sAsin + "\"><a href=\"" + htmlText(sUrl) + "\">" + htmlText(titleWithYear(dRow)) + "</a></h2>\r\n");
        sbMd.Append("## [" + mdText(titleWithYear(dRow)) + "](" + sUrl + ") {#" + sAsin.ToLower() + "}\r\n\r\n");
        lPairs = catalogLinks(dRow, "authors");
        appendField(sbHtml, sbMd, "By", linksHtml(lPairs), linksMd(lPairs));
        appendField(sbHtml, sbMd, "Edition", htmlText(editionText(dRow, "Audible")), editionText(dRow, "Audible"));
        appendField(sbHtml, sbMd, "First published", htmlText(catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)"), catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)");
        appendField(sbHtml, sbMd, "Genres", htmlText(catalogStrings(dRow, "genreLadders").Count == 0 ? "" : string.Join("; ", catalogStrings(dRow, "genreLadders").ToArray())), catalogStrings(dRow, "genreLadders").Count == 0 ? "" : string.Join("; ", catalogStrings(dRow, "genreLadders").ToArray()));
        appendField(sbHtml, sbMd, "Language", htmlText(catalogValue(dRow, "language")), catalogValue(dRow, "language"));
        appendField(sbHtml, sbMd, "Length", htmlText(catalogValue(dRow, "lengthText")), catalogValue(dRow, "lengthText"));
        lPairs = catalogLinks(dRow, "narrators");
        appendField(sbHtml, sbMd, "Narrated by", linksHtml(lPairs), linksMd(lPairs));
        sProgress = Convert.ToBoolean(dRow["finished"]) ? "Finished" : catalogValue(dRow, "timeLeft");
        appendField(sbHtml, sbMd, "Progress", htmlText(sProgress), sProgress);
        appendField(sbHtml, sbMd, "Publisher", htmlText(catalogValue(dRow, "publisher")), catalogValue(dRow, "publisher"));
        appendField(sbHtml, sbMd, "Rating", htmlText(catalogValue(dRow, "ratingText")), catalogValue(dRow, "ratingText"));
        appendField(sbHtml, sbMd, "Release date", htmlText(catalogValue(dRow, "releaseDate")), catalogValue(dRow, "releaseDate"));
        lPairs = catalogLinks(dRow, "series");
        appendField(sbHtml, sbMd, "Series", linksHtml(lPairs), linksMd(lPairs));
        appendField(sbHtml, sbMd, "Wikipedia", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "<a href=\"" + htmlText(catalogValue(dRow, "wikipediaUrl")) + "\">" + htmlText(catalogValue(dRow, "wikipediaTitle")) + "</a>", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "[" + mdText(catalogValue(dRow, "wikipediaTitle")) + "](" + catalogValue(dRow, "wikipediaUrl") + ")");
        appendField(sbHtml, sbMd, "Summary", htmlText(catalogValue(dRow, "summary")), catalogValue(dRow, "summary"));
        Monitor.Exit(dRow);
    }

    static string buildLibraryFiles(string sDownloadDir)
    {
        double nRatingSum;
        int iFinishedCount, iLongestMinutes, iMostRatings, iOldestYear, iRatedCount, iRowMinutes, iRowRatings, iRowYear, iTotalMinutes;
        string sLongestTitle, sMostRatedTitle, sOldestTitle;
        Dictionary<string, int> dAuthorMinutes;
        string sAsin, sHtmPath, sIntroHtml, sIntroMd, sMdPath, sProgress, sStats, sTitle, sUrl;
        Dictionary<string, List<string[]>> dByAuthor, dByGenre, dByNarrator, dByProgress, dByPublisher, dByRating, dBySeries;
        List<Dictionary<string, object>> lOrdered;
        List<string[]> lPairs;
        StringBuilder sbHtml, sbMd;

        if (lCatalog.Count == 0) { log("No catalog rows were harvested, so the library files were not written"); return ""; }
        dByAuthor = new Dictionary<string, List<string[]>>();
        dByGenre = new Dictionary<string, List<string[]>>();
        dByNarrator = new Dictionary<string, List<string[]>>();
        dByProgress = new Dictionary<string, List<string[]>>();
        dByPublisher = new Dictionary<string, List<string[]>>();
        dByRating = new Dictionary<string, List<string[]>>();
        dBySeries = new Dictionary<string, List<string[]>>();
        sbHtml = new StringBuilder();
        sbMd = new StringBuilder();
        dAuthorMinutes = new Dictionary<string, int>();
        iFinishedCount = 0;
        iLongestMinutes = 0;
        iMostRatings = 0;
        iOldestYear = 0;
        iRatedCount = 0;
        iTotalMinutes = 0;
        nRatingSum = 0;
        sLongestTitle = "";
        sMostRatedTitle = "";
        sOldestTitle = "";
        foreach (Dictionary<string, object> dRow in lCatalog)
        {
            Monitor.Enter(dRow);
            sAsin = catalogValue(dRow, "asin");
            sTitle = titleWithYear(dRow);
            if (Convert.ToBoolean(dRow["finished"])) iFinishedCount = iFinishedCount + 1;
            if (dRow.ContainsKey("ratingValue")) { iRatedCount = iRatedCount + 1; nRatingSum = nRatingSum + Convert.ToDouble(dRow["ratingValue"]); }
            iRowMinutes = dRow.ContainsKey("lengthMinutes") ? Convert.ToInt32(dRow["lengthMinutes"]) : 0;
            iTotalMinutes = iTotalMinutes + iRowMinutes;
            if (iRowMinutes > iLongestMinutes) { iLongestMinutes = iRowMinutes; sLongestTitle = sTitle; }
            iRowRatings = dRow.ContainsKey("ratingsCount") ? Convert.ToInt32(dRow["ratingsCount"]) : 0;
            if (iRowRatings > iMostRatings) { iMostRatings = iRowRatings; sMostRatedTitle = sTitle; }
            iRowYear = 0;
            int.TryParse(catalogValue(dRow, "firstPublished"), out iRowYear);
            if (iRowYear > 0 && (iOldestYear == 0 || iRowYear < iOldestYear)) { iOldestYear = iRowYear; sOldestTitle = sTitle; }
            foreach (string[] aPair in catalogLinks(dRow, "authors")) { if (!dAuthorMinutes.ContainsKey(aPair[0])) dAuthorMinutes[aPair[0]] = 0; dAuthorMinutes[aPair[0]] = dAuthorMinutes[aPair[0]] + iRowMinutes; }
            foreach (string[] aPair in catalogLinks(dRow, "authors")) addToIndex(dByAuthor, aPair[0], sTitle, sAsin);
            foreach (string[] aPair in catalogLinks(dRow, "narrators")) addToIndex(dByNarrator, aPair[0], sTitle, sAsin);
            foreach (string[] aPair in catalogLinks(dRow, "series")) addToIndex(dBySeries, aPair[0], sTitle, sAsin);
            addToIndex(dByProgress, Convert.ToBoolean(dRow["finished"]) ? "Finished" : "Not finished", sTitle, sAsin);
            foreach (string sGenre in catalogStrings(dRow, "genresTop")) addToIndex(dByGenre, sGenre, sTitle, sAsin);
            if (catalogValue(dRow, "publisher") != "") addToIndex(dByPublisher, catalogValue(dRow, "publisher"), sTitle, sAsin);
            if (dRow.ContainsKey("ratingValue")) addToIndex(dByRating, ratingBucket(Convert.ToDouble(dRow["ratingValue"])), sTitle, sAsin);
            Monitor.Exit(dRow);
        }
        sStats = "The library holds " + lCatalog.Count + " titles by " + dByAuthor.Count + " authors, read by " + dByNarrator.Count + " narrators, with " +
            dBySeries.Count + " series and " + dByGenre.Count + " Audible genres represented.  " + iFinishedCount + " titles are marked finished.  " +
            (iTotalMinutes > 0 ? "Altogether the library holds about " + (iTotalMinutes / 60).ToString("N0") + " hours of listening.  " : "") +
            (iRatedCount > 0 ? "Across " + iRatedCount + " rated titles, the average Audible rating is " + (nRatingSum / iRatedCount).ToString("0.0") + " out of 5 stars." : "");
        sIntroHtml = "This catalog was generated from the Audible library on " + DateTime.Now.ToString("MMMM d, yyyy") + " by bookFido.  " +
            "Every title appears as a level-two heading, in the order the library lists them with the most recent first, and each heading links to the title's detail page on the Audible website.  " +
            "Under each heading are the details Audible shows in the library, using Audible's own labels, with clickable links for authors, narrators, and series, " +
            "followed by details gathered from public web services: the rating, length, release date, publisher, language, and genres come from Audible's catalog service, the year of first publication comes from Open Library, which also lends the publisher of the print editions, clearly labeled, when Audible offers no publisher, and a Wikipedia field links to the book's own Wikipedia page when a confidently matching one exists; the closing appendix adds author biographies from Wikipedia and Open Library when reliable ones are found.  " +
            "The publisher's summary appears last under each title.  " +
            "After the titles come appendixes that index the library by author, by narrator, by series, by listening progress, by genre, by publisher, by rating, and by the authors themselves; every appendix entry is an internal link that jumps to the title's heading, and the table of contents links to every section.";
        sIntroMd = sIntroHtml;
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>Audible Library</title>\r\n</head>\r\n<body>\r\n");
        sbHtml.Append("<h1>Audible Library</h1>\r\n");
        sbHtml.Append("<nav aria-label=\"Table of contents\">\r\n<h2 id=\"contents\">Contents</h2>\r\n<ul>\r\n");
        lOrdered = new List<Dictionary<string, object>>(lCatalog);
        lOrdered.Sort(compareCatalogRowsByTitle);
        sbHtml.Append("<li><a href=\"#introduction\">Introduction</a></li>\r\n<li><a href=\"#titles\">Titles</a>\r\n<ul>\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbHtml.Append("<li><a href=\"#" + catalogValue(dRow, "asin") + "\">" + htmlText(titleWithYear(dRow)) + "</a></li>\r\n"); }
        }
        sbHtml.Append("</ul>\r\n</li>\r\n");
        sbHtml.Append("<li><a href=\"#appendix-a\">Appendix A: Titles by Author</a></li>\r\n<li><a href=\"#appendix-b\">Appendix B: Titles by Narrator</a></li>\r\n<li><a href=\"#appendix-c\">Appendix C: Titles by Series</a></li>\r\n<li><a href=\"#appendix-d\">Appendix D: Titles by Listening Progress</a></li>\r\n<li><a href=\"#appendix-e\">Appendix E: Titles by Genre</a></li>\r\n<li><a href=\"#appendix-f\">Appendix F: Titles by Publisher</a></li>\r\n<li><a href=\"#appendix-g\">Appendix G: Titles by Rating</a></li>\r\n<li><a href=\"#appendix-h\">Appendix H: About the Authors</a></li>\r\n");
        sbHtml.Append("</ul>\r\n</nav>\r\n");
        sbMd.Append("# Audible Library\r\n\r\n## Contents {#contents}\r\n\r\n");
        sbMd.Append("- [Introduction](#introduction)\r\n- [Titles](#titles)\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbMd.Append("    - [" + mdText(titleWithYear(dRow)) + "](#" + catalogValue(dRow, "asin").ToLower() + ")\r\n"); }
        }
        sbMd.Append("- [Appendix A: Titles by Author](#appendix-a)\r\n- [Appendix B: Titles by Narrator](#appendix-b)\r\n- [Appendix C: Titles by Series](#appendix-c)\r\n- [Appendix D: Titles by Listening Progress](#appendix-d)\r\n- [Appendix E: Titles by Genre](#appendix-e)\r\n- [Appendix F: Titles by Publisher](#appendix-f)\r\n- [Appendix G: Titles by Rating](#appendix-g)\r\n- [Appendix H: About the Authors](#appendix-h)\r\n\r\n");
        sbHtml.Append("<h2 id=\"introduction\">Introduction</h2>\r\n<p>" + htmlText(sIntroHtml) + "</p>\r\n<p>" + htmlText(sStats) + "</p>\r\n");
        sbMd.Append("## Introduction {#introduction}\r\n\r\n" + sIntroMd + "\r\n\r\n" + sStats + "\r\n\r\n");
        appendCountTable(sbHtml, sbMd, "Titles by Audible genre", "Genre", countRows(dByGenre, null));
        appendCountTable(sbHtml, sbMd, "Titles by rating", "Rating", countRows(dByRating, new string[] { "4.5 stars and up", "4.0 to 4.4 stars", "3.5 to 3.9 stars", "Below 3.5 stars" }));
        appendCountTable(sbHtml, sbMd, "Listening progress", "Progress", countRows(dByProgress, new string[] { "Finished", "Not finished" }));
        appendHighlights(sbHtml, sbMd, dByAuthor, dByNarrator, sLongestTitle, iLongestMinutes, sOldestTitle, iOldestYear, sMostRatedTitle, iMostRatings);
        sbHtml.Append("<h2 id=\"titles\">Titles</h2>\r\n");
        sbMd.Append("## Titles {#titles}\r\n\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered) appendAudibleEntry(sbHtml, sbMd, dRow);
        appendAppendix(sbHtml, sbMd, "appendix-a", "Appendix A: Titles by Author", dByAuthor, true);
        appendAppendix(sbHtml, sbMd, "appendix-b", "Appendix B: Titles by Narrator", dByNarrator);
        appendAppendix(sbHtml, sbMd, "appendix-c", "Appendix C: Titles by Series", dBySeries);
        appendAppendix(sbHtml, sbMd, "appendix-d", "Appendix D: Titles by Listening Progress", dByProgress);
        appendAppendix(sbHtml, sbMd, "appendix-e", "Appendix E: Titles by Genre", dByGenre);
        appendAppendix(sbHtml, sbMd, "appendix-f", "Appendix F: Titles by Publisher", dByPublisher);
        appendAppendix(sbHtml, sbMd, "appendix-g", "Appendix G: Titles by Rating", dByRating);
        appendAuthorsAppendix(sbHtml, sbMd, dByAuthor, dAuthorMinutes);
        sbHtml.Append("</body>\r\n</html>\r\n");
        sHtmPath = Path.Combine(sDownloadDir, "Audible_Library.htm");
        sMdPath = Path.Combine(sDownloadDir, "Audible_Library.md");
        File.WriteAllText(sHtmPath, sbHtml.ToString(), new UTF8Encoding(true));
        File.WriteAllText(sMdPath, sbMd.ToString(), new UTF8Encoding(true));
        log("Wrote the library catalog: " + sHtmPath);
        log("Wrote the library catalog: " + sMdPath);
        return sHtmPath;
    }

    // Opens a file in the user's default web browser, whichever browser that is.
    static void openInDefaultBrowser(string sPath)
    {
        ProcessStartInfo oStartInfo;

        try
        {
            oStartInfo = new ProcessStartInfo(sPath);
            oStartInfo.UseShellExecute = true;
            Process.Start(oStartInfo);
            log("Opened the catalog in the default web browser: " + sPath);
        }
        catch (Exception oException)
        {
            log("Could not open the catalog in the default browser: " + oException.Message);
        }
    }

    // Cleans a candidate title into a safe file-name root, porting renTitle's
    // rules: friendly punctuation replacements, illegal character removal,
    // bloat collapsing, junk-title blacklist, and rejection of numeric-only
    // or empty roots.  Returns an empty string when the candidate is unusable.
    static string cleanRoot(string sTitle)
    {
        bool bNumber;
        string sJunkList, sRoot;
        string[] aJunk;

        if (sTitle == null) return "";
        sRoot = sTitle.Trim();
        if (sRoot.StartsWith("Microsoft Word - ")) sRoot = sRoot.Substring(17).Trim();
        if (sRoot.StartsWith("Microsoft PowerPoint - ")) sRoot = sRoot.Substring(23).Trim();
        sRoot = sRoot.Replace(":", " - ").Replace("&", " and ").Replace(";", " - ").Replace("[", "(").Replace("<", "(").Replace("]", ")").Replace(">", ")");
        foreach (char chBad in Path.GetInvalidFileNameChars()) sRoot = sRoot.Replace(chBad, ' ');
        foreach (char chBad in "^%*|/?!" + "\"" + "\\") sRoot = sRoot.Replace(chBad, ' ');
        while (sRoot.Contains("--")) sRoot = sRoot.Replace("--", " - ");
        while (sRoot.Contains("  ")) sRoot = sRoot.Replace("  ", " ");
        sRoot = sRoot.Trim();
        while (sRoot.EndsWith(".") && sRoot.Length > 1) sRoot = sRoot.Substring(0, sRoot.Length - 1).Trim();
        while (sRoot.StartsWith(".") && sRoot.Length > 1) sRoot = sRoot.Substring(1).Trim();
        if (sRoot.Length > 2 && sRoot.StartsWith("(") && sRoot.EndsWith(")")) sRoot = sRoot.Substring(1, sRoot.Length - 2).Trim();
        if (sRoot.Length > 120) sRoot = sRoot.Substring(0, 120).Trim();
        if (sRoot.Length == 0) return "";
        bNumber = true;
        foreach (char chOne in sRoot) { if (!char.IsDigit(chOne)) { bNumber = false; break; } }
        if (bNumber) return "";
        sJunkList = "no-title|(no title)|title|untitled|untitled-1|untitled document|untitled attachment|unknown|none|null|loading|(unspecified)|document 1|powerpoint presentation|presentation1|presentation title|report title|title page|title document|sheet 1|slide 1|no slide title|page title|title goes here|click to edit presentation title|title of the report|main title|main title of the paper|paper title|pdf|companion pdf|audible|audible.com";
        aJunk = sJunkList.Split('|');
        foreach (string sJunk in aJunk) { if (sRoot.ToLower() == sJunk) return ""; }
        return sRoot;
    }

    // Returns a free path for a root name in the download folder, numbering
    // duplicates -001 through -999 the way renTitle does.
    static string uniquePdfPath(string sDir, string sRoot)
    {
        int iTry;
        string sPath;

        sPath = Path.Combine(sDir, sRoot + ".pdf");
        iTry = 1;
        while (File.Exists(sPath) && iTry < 1000)
        {
            sPath = Path.Combine(sDir, sRoot + "-" + iTry.ToString("000") + ".pdf");
            iTry = iTry + 1;
        }
        return sPath;
    }

    // Reads the Title metadata out of a PDF without any external tool: first
    // the last /Title entry in the file, which incremental updates make the
    // current one, then the XMP dc:title.  Literal strings with escapes, hex
    // strings, and UTF-16 big-endian values are all handled.
    static string pdfTitleFromFile(string sPath)
    {
        string sHay;
        FileInfo fInfo;

        try
        {
            fInfo = new FileInfo(sPath);
            if (fInfo.Length > 60000000) return "";
            sHay = Encoding.GetEncoding(28591).GetString(File.ReadAllBytes(sPath));
        }
        catch (Exception oException)
        {
            log("Could not read the PDF for its title: " + oException.Message);
            return "";
        }
        return pdfTitleFromHay(sHay);
    }

    // Reads the Title metadata out of already-loaded PDF bytes.
    static string pdfTitleFromHay(string sHay)
    {
        int iPos;
        string sCandidate, sXmpBlock;

        iPos = sHay.LastIndexOf("/Title", StringComparison.Ordinal);
        while (iPos >= 0)
        {
            sCandidate = pdfStringAt(sHay, iPos + 6);
            if (cleanRoot(sCandidate) != "") return sCandidate;
            iPos = iPos > 0 ? sHay.LastIndexOf("/Title", iPos - 1, StringComparison.Ordinal) : -1;
        }
        sXmpBlock = betweenMarkers(sHay, "<dc:title>", "</dc:title>");
        if (sXmpBlock != "")
        {
            sCandidate = betweenMarkers(sXmpBlock, "<rdf:li", "</rdf:li>");
            if (sCandidate.Contains(">")) sCandidate = sCandidate.Substring(sCandidate.IndexOf(">") + 1);
            sCandidate = xmlDecode(sCandidate).Trim();
            if (cleanRoot(sCandidate) != "") return sCandidate;
        }
        return "";
    }

    // Parses one PDF string that starts at the given offset: a literal string
    // in parentheses with backslash escapes and octal codes, or a hex string
    // in angle brackets.  A UTF-16 big-endian byte order mark selects that
    // decoding; otherwise the bytes are read as Latin-1, which covers the
    // PDFDocEncoding range well enough for titles.
    static string pdfStringAt(string sHay, int iStart)
    {
        bool bHex;
        int iCode, iDepth, iDigits, iPos;
        string sHexPair;
        List<byte> lBytes;

        iPos = iStart;
        while (iPos < sHay.Length && (sHay[iPos] == ' ' || sHay[iPos] == (char) 13 || sHay[iPos] == (char) 10 || sHay[iPos] == (char) 9)) iPos = iPos + 1;
        if (iPos >= sHay.Length) return "";
        bHex = sHay[iPos] == '<';
        if (!bHex && sHay[iPos] != '(') return "";
        iPos = iPos + 1;
        lBytes = new List<byte>();
        if (bHex)
        {
            sHexPair = "";
            while (iPos < sHay.Length && sHay[iPos] != '>')
            {
                if (Uri.IsHexDigit(sHay[iPos]))
                {
                    sHexPair = sHexPair + sHay[iPos];
                    if (sHexPair.Length == 2) { lBytes.Add(Convert.ToByte(sHexPair, 16)); sHexPair = ""; }
                }
                iPos = iPos + 1;
            }
        }
        else
        {
            iDepth = 1;
            while (iPos < sHay.Length)
            {
                if (sHay[iPos] == (char) 92)
                {
                    iPos = iPos + 1;
                    if (iPos >= sHay.Length) break;
                    if (sHay[iPos] == 'n') lBytes.Add(10);
                    else if (sHay[iPos] == 'r') lBytes.Add(13);
                    else if (sHay[iPos] == 't') lBytes.Add(9);
                    else if (sHay[iPos] == 'b') lBytes.Add(8);
                    else if (sHay[iPos] == 'f') lBytes.Add(12);
                    else if (char.IsDigit(sHay[iPos]))
                    {
                        iCode = 0;
                        iDigits = 0;
                        while (iPos < sHay.Length && char.IsDigit(sHay[iPos]) && iDigits < 3)
                        {
                            iCode = iCode * 8 + (sHay[iPos] - '0');
                            iDigits = iDigits + 1;
                            iPos = iPos + 1;
                        }
                        iPos = iPos - 1;
                        lBytes.Add((byte) (iCode % 256));
                    }
                    else if (sHay[iPos] != (char) 13 && sHay[iPos] != (char) 10) lBytes.Add((byte) sHay[iPos]);
                }
                else if (sHay[iPos] == '(') { iDepth = iDepth + 1; lBytes.Add((byte) '('); }
                else if (sHay[iPos] == ')')
                {
                    iDepth = iDepth - 1;
                    if (iDepth == 0) break;
                    lBytes.Add((byte) ')');
                }
                else lBytes.Add((byte) sHay[iPos]);
                iPos = iPos + 1;
            }
        }
        if (lBytes.Count >= 2 && lBytes[0] == 254 && lBytes[1] == 255) return Encoding.BigEndianUnicode.GetString(lBytes.ToArray(), 2, lBytes.Count - 2);
        return Encoding.GetEncoding(28591).GetString(lBytes.ToArray());
    }

    // Returns the text between two markers, or an empty string.
    static string betweenMarkers(string sHay, string sBegin, string sEnd)
    {
        int iBegin, iEnd;

        iBegin = sHay.IndexOf(sBegin, StringComparison.OrdinalIgnoreCase);
        if (iBegin < 0) return "";
        iBegin = iBegin + sBegin.Length;
        iEnd = sHay.IndexOf(sEnd, iBegin, StringComparison.OrdinalIgnoreCase);
        if (iEnd < 0) return "";
        return sHay.Substring(iBegin, iEnd - iBegin);
    }

    // Decodes the handful of XML entities that appear in XMP titles.
    static string xmlDecode(string sText)
    {
        return sText.Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&#39;", "'").Replace("&apos;", "'");
    }

    // Derives the file name from the Content-Disposition header, then from the
    // final response url, and finally from the asin in the companion-file url.
    static string fileNameFromResponse(HttpWebResponse httpResponse, string sUrl)
    {
        int iPos;
        string sDisposition, sName;

        sName = "";
        sDisposition = httpResponse.Headers["Content-Disposition"];
        if (sDisposition != null)
        {
            iPos = sDisposition.IndexOf("filename*=UTF-8''", StringComparison.OrdinalIgnoreCase);
            if (iPos >= 0) sName = sDisposition.Substring(iPos + 17);
            if (sName == "")
            {
                iPos = sDisposition.IndexOf("filename=", StringComparison.OrdinalIgnoreCase);
                if (iPos >= 0) sName = sDisposition.Substring(iPos + 9);
            }
            if (sName.Contains(";")) sName = sName.Substring(0, sName.IndexOf(";"));
            sName = Uri.UnescapeDataString(sName.Trim().Trim('"'));
        }
        if (sName == "")
        {
            sName = Uri.UnescapeDataString(Path.GetFileName(httpResponse.ResponseUri.AbsolutePath));
            if (!sName.ToLower().EndsWith(".pdf")) sName = "";
        }
        if (sName == "") sName = asinFromUrl(sUrl) + ".pdf";
        foreach (char chBad in Path.GetInvalidFileNameChars()) sName = sName.Replace(chBad.ToString(), "_");
        if (!sName.ToLower().EndsWith(".pdf")) sName = sName + ".pdf";
        return sName;
    }

    // Politely asks the browser to close, then makes sure the process tree we
    // launched is gone.  The program opened Edge, so it cleans up after
    // itself instead of leaving a library page on screen.
    static async Task closeEdgeAsync()
    {
        try
        {
            if (wsCdp != null && wsCdp.State == WebSocketState.Open)
            {
                log("Asking Edge to close itself");
                await cdpSendOnce("Browser.close", null);
                log("Edge accepted the close request");
                Thread.Sleep(2000);
            }
        }
        catch (Exception oException)
        {
            log("Browser.close was not accepted, so the process tree will be closed instead: " + oException.Message);
        }
        shutdownEdge();
    }

    // Ensures the Edge process tree launched by this run is terminated.  If
    // this run only reused an instance left over from before, there is no
    // tracked process, and any remaining window is reported instead.
    static void shutdownEdge()
    {
        int iPid;

        if (oEdgeProcess == null) return;
        iPid = -1;
        try { iPid = oEdgeProcess.Id; } catch (Exception) { }
        try { if (!oEdgeProcess.HasExited) oEdgeProcess.WaitForExit(2000); } catch (Exception) { }
        if (iPid > 0) killTree(iPid);
        oEdgeProcess = null;
    }

    // Terminates a whole process tree.  Edge is not one process: the one we
    // start spawns renderer, GPU, and utility children, and killing only the
    // parent can leave windows on screen.
    static void killTree(int iPid)
    {
        ProcessStartInfo oStartInfo;
        Process oKill;

        try
        {
            oStartInfo = new ProcessStartInfo("taskkill", "/PID " + iPid + " /T /F");
            oStartInfo.CreateNoWindow = true;
            oStartInfo.UseShellExecute = false;
            oKill = Process.Start(oStartInfo);
            if (oKill != null) oKill.WaitForExit(5000);
            log("Closed the Edge process tree " + iPid);
        }
        catch (Exception oException)
        {
            log("Could not close the Edge process tree " + iPid + ": " + oException.Message);
        }
    }

    // Loads assemblies embedded as manifest resources into this exe.  The
    // build script embeds the PdfPig assemblies with csc /resource, and the
    // CLR raises AssemblyResolve when it cannot find them on disk; this
    // handler reads the bytes from the embedded resource and loads them
    // in memory, the same technique 2htm uses for Markdig.
    static Assembly resolveEmbeddedAssembly(object oSender, ResolveEventArgs oArgs)
    {
        byte[] aBytes;
        int iJust, iRead;
        string sResource;

        try
        {
            sResource = new AssemblyName(oArgs.Name).Name + ".dll";
            using (Stream streamResource = Assembly.GetExecutingAssembly().GetManifestResourceStream(sResource))
            {
                if (streamResource == null) { log("No embedded resource found for assembly: " + sResource); return null; }
                aBytes = new byte[streamResource.Length];
                iRead = 0;
                while (iRead < aBytes.Length)
                {
                    iJust = streamResource.Read(aBytes, iRead, aBytes.Length - iRead);
                    if (iJust <= 0) break;
                    iRead = iRead + iJust;
                }
                log("Loaded embedded assembly: " + sResource);
                return Assembly.Load(aBytes);
            }
        }
        catch (Exception) { return null; }
    }

    // Ensures a same-root-name .htm sibling exists for a companion PDF, so
    // screen reader users can read the content in a browser instead of a PDF
    // viewer.  Failures are logged and never disturb the download bookkeeping.
    static int ensureHtmVersion(string sPdfPath)
    {
        string sHtmPath;

        sHtmPath = Path.ChangeExtension(sPdfPath, ".htm");
        if (File.Exists(sHtmPath)) { log("The HTML version already exists: " + sHtmPath); return iHtmStatusExisted; }
        log("Creating an HTML version for: " + Path.GetFileName(sPdfPath));
        try
        {
            convertPdfToHtm(sPdfPath, sHtmPath);
            iHtmCount = iHtmCount + 1;
            log("Created the HTML version: " + sHtmPath);
            return iHtmStatusCreated;
        }
        catch (Exception oException)
        {
            log("Could not create an HTML version of " + Path.GetFileName(sPdfPath) + ": " + oException.Message);
            return iHtmStatusFailed;
        }
    }

    // Builds the short spoken announcement for a companion file that was
    // already on disk or has just been renamed, mentioning the PDF's base
    // name and, when one exists or was just made, its HTML companion.
    static string foundFileAnnouncement(string sLead, string sPdfName, int iHtmStatus)
    {
        if (iHtmStatus == iHtmStatusExisted) return sLead + " " + sPdfName + " and .htm";
        if (iHtmStatus == iHtmStatusCreated) return sLead + " " + sPdfName + "; made .htm";
        return sLead + " " + sPdfName;
    }

    // Converts one PDF to structured HTML with the embedded PdfPig library.
    // Pass one gathers every text block and a histogram of letter point
    // sizes; the most frequent size is the body text, and larger sizes form
    // heading tiers mapped to h2 through h4.  Pass two renders each block in
    // top-to-bottom, left-to-right order as a heading, a bullet or numbered
    // list, or a paragraph with hyphenated line breaks rejoined.  Images are
    // dropped, and any /Alt descriptions found in the raw PDF are listed in
    // an Image descriptions section at the end.
    static void convertPdfToHtm(string sPdfPath, string sHtmPath)
    {
        double nBlockSize, nBodySize, nCount, nSize;
        int iLastLevel, iLevel;
        string sBlockText, sCarry, sDocTitle, sHay, sLang, sNormalized;
        Dictionary<double, double> dSizeCounts;
        Dictionary<double, int> dTierLevels;
        Dictionary<string, int> dRepeatCounts;
        HashSet<string> setOnThisPage;
        List<double> lHeadingSizes;
        List<string> lAltTexts;
        List<TextBlock> lPageBlocks;
        List<List<TextBlock>> lAllPages;
        StringBuilder sbHtml;

        dSizeCounts = new Dictionary<double, double>();
        lAllPages = new List<List<TextBlock>>();
        using (PdfDocument pdfDocument = PdfDocument.Open(sPdfPath))
        {
            foreach (Page pdfPage in pdfDocument.GetPages())
            {
                foreach (Letter pdfLetter in pdfPage.Letters)
                {
                    nSize = Math.Round(pdfLetter.PointSize * 2.0) / 2.0;
                    if (!dSizeCounts.ContainsKey(nSize)) dSizeCounts[nSize] = 0;
                    dSizeCounts[nSize] = dSizeCounts[nSize] + 1;
                }
                lPageBlocks = new List<TextBlock>(DocstrumBoundingBoxes.Instance.GetBlocks(NearestNeighbourWordExtractor.Instance.GetWords(pdfPage.Letters)));
                lPageBlocks.Sort(compareTextBlocks);
                lAllPages.Add(lPageBlocks);
            }
        }
        nBodySize = 0;
        nCount = 0;
        foreach (double nOne in dSizeCounts.Keys) { if (dSizeCounts[nOne] > nCount) { nCount = dSizeCounts[nOne]; nBodySize = nOne; } }
        lHeadingSizes = new List<double>();
        foreach (double nOne in dSizeCounts.Keys) { if (nOne > nBodySize + 1.0) lHeadingSizes.Add(nOne); }
        lHeadingSizes.Sort();
        lHeadingSizes.Reverse();
        dTierLevels = new Dictionary<double, int>();
        foreach (double nOne in lHeadingSizes) dTierLevels[nOne] = Math.Min(4, 2 + dTierLevels.Count);
        // Count how many pages contain each short normalized block text, so
        // running headers and footers can be dropped as pagination artifacts.
        dRepeatCounts = new Dictionary<string, int>();
        foreach (List<TextBlock> lBlocks in lAllPages)
        {
            setOnThisPage = new HashSet<string>();
            foreach (TextBlock textBlock in lBlocks)
            {
                sNormalized = normalizedFurniture(joinBlockLines(textBlock));
                if (sNormalized == "" || setOnThisPage.Contains(sNormalized)) continue;
                setOnThisPage.Add(sNormalized);
                if (!dRepeatCounts.ContainsKey(sNormalized)) dRepeatCounts[sNormalized] = 0;
                dRepeatCounts[sNormalized] = dRepeatCounts[sNormalized] + 1;
            }
        }
        sHay = "";
        try { sHay = Encoding.GetEncoding(28591).GetString(File.ReadAllBytes(sPdfPath)); }
        catch (Exception) { }
        sLang = langFromHay(sHay);
        if (sLang == "") sLang = "en";
        sDocTitle = pdfTitleFromHay(sHay);
        if (cleanRoot(sDocTitle) == "") sDocTitle = Path.GetFileNameWithoutExtension(sPdfPath);
        sbHtml = new StringBuilder();
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"" + htmlText(sLang) + "\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>" + htmlText(sDocTitle) + "</title>\r\n</head>\r\n<body>\r\n<main>\r\n");
        sbHtml.Append("<h1>" + htmlText(sDocTitle) + "</h1>\r\n");
        iLastLevel = 1;
        sCarry = "";
        foreach (List<TextBlock> lBlocks in lAllPages)
        {
            foreach (TextBlock textBlock in lBlocks)
            {
                sBlockText = joinBlockLines(textBlock);
                if (sBlockText == "") continue;
                if (sBlockText.Length <= 2 && !isNumericOnly(sBlockText) && !isRomanPageMarker(sBlockText))
                {
                    // A one- or two-character block is a drop cap or stray
                    // symbol; carry it into the next block so the sentence
                    // it opens stays whole.
                    sCarry = sCarry + sBlockText;
                    continue;
                }
                if (sCarry != "")
                {
                    sBlockText = (sCarry.Length == 1 && char.IsLetter(sCarry[0]) && sBlockText.Length > 0 && char.IsLower(sBlockText[0])) ? sCarry + sBlockText : sCarry + " " + sBlockText;
                    sCarry = "";
                }
                sBlockText = collapseDotLeaders(sBlockText);
                sNormalized = normalizedFurniture(sBlockText);
                if (isNumericOnly(sBlockText)) continue;
                if (isRomanPageMarker(sBlockText)) continue;
                if (sNormalized != "" && dRepeatCounts.ContainsKey(sNormalized) && dRepeatCounts[sNormalized] >= 3) continue;
                nBlockSize = blockPointSize(textBlock);
                iLevel = 0;
                if (dTierLevels.ContainsKey(nBlockSize)) iLevel = dTierLevels[nBlockSize];
                if (iLevel > 0 && sBlockText.Length <= 200)
                {
                    if (iLevel > iLastLevel + 1) iLevel = iLastLevel + 1;
                    sbHtml.Append("<h" + iLevel + ">" + htmlText(sBlockText) + "</h" + iLevel + ">\r\n");
                    iLastLevel = iLevel;
                    continue;
                }
                if (appendListBlock(sbHtml, textBlock)) continue;
                sbHtml.Append("<p>" + htmlText(sBlockText) + "</p>\r\n");
            }
        }
        lAltTexts = altTextsFromHay(sHay);
        if (lAltTexts.Count > 0)
        {
            sbHtml.Append("<h2>Image descriptions</h2>\r\n<p>The PDF contains images with the following alternative text descriptions.</p>\r\n<ul>\r\n");
            foreach (string sAlt in lAltTexts) sbHtml.Append("<li>" + htmlText(sAlt) + "</li>\r\n");
            sbHtml.Append("</ul>\r\n");
            iLastLevel = 2;
        }
        sbHtml.Append("</main>\r\n</body>\r\n</html>\r\n");
        File.WriteAllText(sHtmPath, sbHtml.ToString(), new UTF8Encoding(true));
    }

    // Returns true when a block is only digits and whitespace, the shape of a
    // bare page number, which PDF/UA treats as a pagination artifact.
    static bool isNumericOnly(string sText)
    {
        bool bSawDigit;

        bSawDigit = false;
        foreach (char chOne in sText)
        {
            if (char.IsDigit(chOne)) { bSawDigit = true; continue; }
            if (!char.IsWhiteSpace(chOne)) return false;
        }
        return bSawDigit;
    }

    // Returns true for a standalone roman-numeral page marker such as the
    // v or xii of a front-matter page, another pagination artifact.
    static bool isRomanPageMarker(string sText)
    {
        string sTrimmed;

        sTrimmed = sText.Trim();
        if (sTrimmed.Length == 0 || sTrimmed.Length > 6) return false;
        foreach (char chOne in sTrimmed) { if (!char.IsLower(chOne)) return false; }
        return Regex.IsMatch(sTrimmed, "^l?x{0,3}(ix|iv|v?i{0,3})$");
    }

    // Collapses a run of table-of-contents dot leaders to one ellipsis, so
    // a screen reader hears the entry and its page number without a march
    // of periods between them.
    static string collapseDotLeaders(string sText)
    {
        return Regex.Replace(sText, "(?:\\s*\\.){4,}", " \u2026");
    }

    // Normalizes a short block for the repeated-furniture check: digits become
    // a placeholder so "Page 3" and "Page 17" match, whitespace collapses, and
    // long blocks return empty because real content is never page furniture.
    static string normalizedFurniture(string sText)
    {
        StringBuilder sbOut;

        if (sText.Length == 0 || sText.Length > 120) return "";
        sbOut = new StringBuilder();
        foreach (char chOne in sText)
        {
            if (char.IsDigit(chOne)) { if (sbOut.Length == 0 || sbOut[sbOut.Length - 1] != '#') sbOut.Append('#'); }
            else if (char.IsWhiteSpace(chOne)) { if (sbOut.Length > 0 && sbOut[sbOut.Length - 1] != ' ') sbOut.Append(' '); }
            else sbOut.Append(char.ToLower(chOne));
        }
        return sbOut.ToString().Trim();
    }

    // Reads the PDF's declared natural language from its /Lang entry, per the
    // Matterhorn Protocol's declared-language checkpoint.  Returns an empty
    // string when no plausible declaration is present.
    static string langFromHay(string sHay)
    {
        int iPos;
        string sValue;

        iPos = sHay.IndexOf("/Lang", StringComparison.Ordinal);
        while (iPos >= 0)
        {
            sValue = pdfStringAt(sHay, iPos + 5).Trim();
            if (looksLikeLanguageTag(sValue)) return sValue;
            iPos = sHay.IndexOf("/Lang", iPos + 5, StringComparison.Ordinal);
        }
        return "";
    }

    // Returns true for strings shaped like a language tag such as en or en-US.
    static bool looksLikeLanguageTag(string sValue)
    {
        if (sValue.Length < 2 || sValue.Length > 12) return false;
        foreach (char chOne in sValue) { if (!char.IsLetterOrDigit(chOne) && chOne != '-') return false; }
        return char.IsLetter(sValue[0]);
    }

    // Orders text blocks top to bottom, then left to right.  PDF y
    // coordinates grow upward, so a larger top comes first.
    static int compareTextBlocks(TextBlock oFirst, TextBlock oSecond)
    {
        if (oFirst.BoundingBox.Top > oSecond.BoundingBox.Top + 1.0) return -1;
        if (oSecond.BoundingBox.Top > oFirst.BoundingBox.Top + 1.0) return 1;
        return oFirst.BoundingBox.Left.CompareTo(oSecond.BoundingBox.Left);
    }

    // Returns the rounded average letter point size of a block.
    static double blockPointSize(TextBlock textBlock)
    {
        double nSum;
        int iLetters;

        nSum = 0;
        iLetters = 0;
        foreach (TextLine textLine in textBlock.TextLines)
        {
            foreach (Word pdfWord in textLine.Words)
            {
                foreach (Letter pdfLetter in pdfWord.Letters) { nSum = nSum + pdfLetter.PointSize; iLetters = iLetters + 1; }
            }
        }
        if (iLetters == 0) return 0;
        return Math.Round(nSum / iLetters * 2.0) / 2.0;
    }

    // Joins a block's lines into one paragraph, rejoining hyphenated line
    // breaks and collapsing whitespace.
    static string joinBlockLines(TextBlock textBlock)
    {
        string sLine;
        StringBuilder sbOut;

        sbOut = new StringBuilder();
        foreach (TextLine textLine in textBlock.TextLines)
        {
            sLine = textLine.Text == null ? "" : textLine.Text.Trim();
            if (sLine == "") continue;
            if (sbOut.Length > 0)
            {
                if (sbOut[sbOut.Length - 1] == '-' && sLine.Length > 0 && char.IsLower(sLine[0])) sbOut.Length = sbOut.Length - 1;
                else sbOut.Append(" ");
            }
            sbOut.Append(sLine);
        }
        return sbOut.ToString().Trim();
    }

    // Detects a block whose lines are bullet or numbered items and appends
    // it as a ul or ol; returns false when the block is not a list.
    static bool appendListBlock(StringBuilder sbHtml, TextBlock textBlock)
    {
        int iBulleted, iLines, iNumbered;
        string sItem, sLine;
        List<string> lItems;

        iBulleted = 0;
        iLines = 0;
        iNumbered = 0;
        lItems = new List<string>();
        foreach (TextLine textLine in textBlock.TextLines)
        {
            sLine = textLine.Text == null ? "" : textLine.Text.Trim();
            if (sLine == "") continue;
            iLines = iLines + 1;
            sItem = stripListMarker(sLine);
            if (sItem != sLine && sLine.Length > 0 && char.IsDigit(sLine[0])) iNumbered = iNumbered + 1;
            else if (sItem != sLine) iBulleted = iBulleted + 1;
            lItems.Add(sItem);
        }
        if (iLines < 2) return false;
        if (iBulleted + iNumbered < iLines) return false;
        sbHtml.Append(iNumbered > iBulleted ? "<ol>\r\n" : "<ul>\r\n");
        foreach (string sOne in lItems) sbHtml.Append("<li>" + htmlText(sOne) + "</li>\r\n");
        sbHtml.Append(iNumbered > iBulleted ? "</ol>\r\n" : "</ul>\r\n");
        return true;
    }

    // Removes a leading bullet character or a leading number-and-period
    // marker from a line; returns the line unchanged when no marker is
    // present.
    static string stripListMarker(string sLine)
    {
        int iPos;
        string sBullets;

        sBullets = "\u2022\u25E6\u25AA\u2013\u2014\u00B7*";
        if (sLine.Length > 1 && sBullets.IndexOf(sLine[0]) >= 0) return sLine.Substring(1).Trim();
        iPos = 0;
        while (iPos < sLine.Length && char.IsDigit(sLine[iPos])) iPos = iPos + 1;
        if (iPos > 0 && iPos < sLine.Length && sLine[iPos] == '.' && iPos <= 3) return sLine.Substring(iPos + 1).Trim();
        return sLine;
    }

    // Collects distinct /Alt image descriptions from the raw PDF bytes,
    // reusing the same string parser as the Title metadata reader.
    static List<string> altTextsFromHay(string sHay)
    {
        int iPos;
        string sValue;
        List<string> lAltTexts;

        lAltTexts = new List<string>();
        iPos = sHay.IndexOf("/Alt", StringComparison.Ordinal);
        while (iPos >= 0)
        {
            sValue = pdfStringAt(sHay, iPos + 4).Trim();
            if (sValue.Length > 2 && !lAltTexts.Contains(sValue)) lAltTexts.Add(sValue);
            iPos = sHay.IndexOf("/Alt", iPos + 4, StringComparison.Ordinal);
        }
        return lAltTexts;
    }

    // Shows a message box for about iMessageBoxMs milliseconds with the base
    // file name as its title, which JAWS speaks automatically.  Uses the
    // long-stable but undocumented MessageBoxTimeoutW function in user32.
    static void showTimedMessageBox(string sTitle)
    {
        log("Announcing: " + sTitle + (sProgressText != "" ? " (" + sProgressText + ")" : ""));
        showActivatedTimedBox(sTitle);
    }

    // Shows the timed message box from a helper thread and, from this
    // thread, actively fights for its focus, because a screen reader only
    // speaks a window that is truly activated.  Three documented stages are
    // tried in turn until the box is the foreground window: a plain
    // SetForegroundWindow; then AttachThreadInput to the current foreground
    // window's thread, which satisfies the last-input-event criterion, and
    // and activation and focus applied while attached.
    static void showActivatedTimedBox(string sCaption)
    {
        int iTry;
        IntPtr hBox;
        Thread threadBox;

        threadBox = new Thread(delegate() { MessageBoxTimeoutW(IntPtr.Zero, sProgressText, sCaption, iMbOk | iMbSetForeground | iMbTopmost, 0, (uint) iMessageBoxMs); });
        threadBox.IsBackground = true;
        threadBox.Start();
        hBox = IntPtr.Zero;
        for (iTry = 0; iTry < 30 && hBox == IntPtr.Zero; iTry = iTry + 1)
        {
            Thread.Sleep(20);
            hBox = FindWindowW("#32770", sCaption);
        }
        if (hBox == IntPtr.Zero) log("The announcement window was not found in time: " + sCaption);
        else if (!forceForeground(hBox)) log("The announcement window could not take focus: " + sCaption);
        threadBox.Join();
    }

    // Activates the announcement window using the classic attach recipe.
    // AttachThreadInput is documented to require a message queue on the
    // calling thread, and a windowless main thread has none, so the first
    // step is a PeekMessage call, which causes the system to create the
    // queue.  The thread then attaches to the input of both the current
    // foreground window's thread and the box's own thread, which satisfies
    // the last-input-event criterion, activates and focuses the box, and
    // detaches.  A synthesized ALT tap is deliberately not used: the tap
    // lands as real input in the foreground browser, whose menu mode then
    // trips the no-active-menus rule that SetForegroundWindow enforces.



    // How many rows of a catalog carry a given key.
    static int countRowsWithKey(List<Dictionary<string, object>> lRows, string sKey)
    {
        int iCount;

        iCount = 0;
        foreach (Dictionary<string, object> dRow in lRows) { if (dRow.ContainsKey(sKey)) iCount = iCount + 1; }
        return iCount;
    }

    // Reports whether the current page is showing a sign-in form in place:
    // some sites present one at the library's own address instead of
    // redirecting, which a url check alone cannot see.
    static async Task<bool> loginFormShowing()
    {
        string sProbe;

        sProbe = await evaluate("(function () { return document.querySelector(\"input[type='password']\") != null ? \"yes\" : \"no\"; })()");
        return sProbe == "yes";
    }

    // ---- The Bookshare library ------------------------------------------
    // The My History page at bookshare.org is server-rendered: each entry
    // is a resultsBook block with the title link carrying the book id,
    // author links in natural order, and metadata fields as label and
    // value pairs, walked by offset.  The same book downloaded more than
    // once collapses into one catalog row that counts its history entries.
    // A failure anywhere in this phase is announced and the run continues.
    // The My History page's filter defaults to a rolling one-month date
    // window, which hides everything older; the walk clears it by sending
    // an explicit wide range in the url, since the filter form submits its
    // fields by GET to the same address.
    static string bookshareHistoryUrl(int iOffset)
    {
        return "https://www.bookshare.org/bookHistory?moduleName=public&titleDownloadPackagingStatus=&startDate=" + Uri.EscapeDataString("01/01/2000") + "&endDate=" + Uri.EscapeDataString(DateTime.Now.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture)) + "&offset=" + iOffset;
    }

    static async Task harvestBookshareAsync()
    {
        int iOffset, iPage;
        string sCurrent, sFirstId, sJson, sPreviousFirstId;
        Dictionary<string, object> dItem, dReply, dRow, dSavedRow;
        Dictionary<string, Dictionary<string, object>> dByKey, dSavedById, dSeenById;
        object[] aRows;

        try
        {
            sProgressText = "";
            log("Bookshare history harvest starting, with the date filter cleared from 01/01/2000 through today");
            await navigate(bookshareHistoryUrl(0));
            await Task.Delay(3000);
            sCurrent = await evaluate("location.href");
            while (sCurrent.ToLower().Contains("/login") || sCurrent.ToLower().Contains("signin"))
            {
                focusWhenShown("bookFido: log in to Bookshare");
                if (MessageBox.Show("You are not logged in to Bookshare yet.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Bookshare history this run.", "bookFido: log in to Bookshare", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Bookshare history was skipped at the login prompt"); return; }
                await navigate(bookshareHistoryUrl(0));
                await Task.Delay(3000);
                sCurrent = await evaluate("location.href");
            }
            dSavedById = new Dictionary<string, Dictionary<string, object>>();
            if (aSavedBookshareRows != null)
            {
                foreach (object oRow in aSavedBookshareRows)
                {
                    dSavedRow = oRow as Dictionary<string, object>;
                    if (dSavedRow != null && dSavedRow.ContainsKey("bsId")) dSavedById[Convert.ToString(dSavedRow["bsId"])] = dSavedRow;
                }
            }
            dByKey = crossLibraryMap();
            dSeenById = new Dictionary<string, Dictionary<string, object>>();
            iOffset = 0;
            iPage = 0;
            sPreviousFirstId = "";
            while (true)
            {
                iPage = iPage + 1;
                if (iPage > 500) { log("The Bookshare harvest stopped at the safety cap of 500 pages"); break; }
                if (iPage > 1)
                {
                    await navigate(bookshareHistoryUrl(iOffset));
                    await Task.Delay(700);
                }
                sJson = await evaluate(bookshareScanScript());
                if (sJson == "") { log("The Bookshare page scan returned no result, so the harvest stopped at offset " + iOffset); break; }
                dReply = (Dictionary<string, object>) jsonCodec.DeserializeObject(sJson);
                aRows = dReply.ContainsKey("rows") ? dReply["rows"] as object[] : null;
                if (aRows == null || aRows.Length == 0)
                {
                    if (iPage == 1 && await loginFormShowing())
                    {
                        focusWhenShown("bookFido: log in to Bookshare");
                        if (MessageBox.Show("Bookshare is showing a sign-in form.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Bookshare history this run.", "bookFido: log in to Bookshare", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Bookshare history was skipped at the login prompt"); return; }
                        await navigate(bookshareHistoryUrl(0));
                        await Task.Delay(3000);
                        iPage = 0;
                        iOffset = 0;
                        continue;
                    }
                    if (iPage == 1) log("The Bookshare history page held no entries; the page title was: " + await evaluate("document.title") + "; the address was: " + await evaluate("location.href"));
                    break;
                }
                sFirstId = "";
                foreach (object oItem in aRows)
                {
                    dItem = oItem as Dictionary<string, object>;
                    if (dItem == null || !dItem.ContainsKey("id")) continue;
                    if (sFirstId == "") sFirstId = Convert.ToString(dItem["id"]);
                    if (dSeenById.ContainsKey(Convert.ToString(dItem["id"])))
                    {
                        dRow = dSeenById[Convert.ToString(dItem["id"])];
                        dRow["bsCount"] = (dRow.ContainsKey("bsCount") ? Convert.ToInt32(dRow["bsCount"]) : 1) + 1;
                        continue;
                    }
                    dRow = bookshareRowFromItem(dItem, dSavedById);
                    if (dRow == null) continue;
                    dSeenById[Convert.ToString(dItem["id"])] = dRow;
                    lBookshareCatalog.Add(dRow);
                    enqueueSharedEnrichment(dRow, dByKey);
                }
                if (sFirstId == "" || sFirstId == sPreviousFirstId) break;
                sPreviousFirstId = sFirstId;
                iOffset = iOffset + aRows.Length;
                if (iPage % 3 == 0) showTimedMessageBox("Bookshare history: " + lBookshareCatalog.Count + " books so far");
            }
            log("Bookshare history harvest complete: " + lBookshareCatalog.Count + " books, of which " + countRowsWithKey(lBookshareCatalog, "twinKey") + " share details with another library");
            showTimedMessageBox("Bookshare history: " + lBookshareCatalog.Count + " books found");
            savePeriodically(true);
        }
        catch (Exception oException)
        {
            log("The Bookshare history could not be harvested, so the run continues without it: " + oException.Message);
            showTimedMessageBox("The Bookshare history could not be gathered this run");
        }
    }

    // The in-page scan of one My History page: each resultsBook block gives
    // the title link with the book id, the author links, and every
    // metadata field as its own label and value, with the inline help
    // popovers stripped before the value is read.
    static string bookshareScanScript()
    {
        return "(function () {" +
            " var lOut = [];" +
            " var lBlocks = document.querySelectorAll(\"div.resultsBook\");" +
            " for (var i = 0; i < lBlocks.length; i++) {" +
            "  var oBlock = lBlocks[i];" +
            "  var o = { fields: {}, authors: [] };" +
            "  var oTitle = oBlock.querySelector(\"h2.bookTitle a\");" +
            "  if (!oTitle) continue;" +
            "  o.title = oTitle.textContent.trim();" +
            "  var sHref = oTitle.getAttribute(\"href\") || \"\";" +
            "  var oIdMatch = sHref.match(/\\/browse\\/book\\/(\\d+)/);" +
            "  o.id = oIdMatch ? oIdMatch[1] : \"\";" +
            "  var lAuthors = oBlock.querySelectorAll(\"span.bookAuthor a\");" +
            "  for (var j = 0; j < lAuthors.length; j++) o.authors.push(lAuthors[j].textContent.trim());" +
            "  var lFields = oBlock.querySelectorAll(\"p.metadataField\");" +
            "  for (var k = 0; k < lFields.length; k++) {" +
            "   var oClone = lFields[k].cloneNode(true);" +
            "   var oStrong = oClone.querySelector(\"strong\");" +
            "   var sLabel = oStrong ? oStrong.textContent.replace(\":\", \"\").trim() : \"\";" +
            "   if (oStrong) oStrong.parentNode.removeChild(oStrong);" +
            "   var lStrip = oClone.querySelectorAll(\"a.inlineHelp, div[id], select, script, style\");" +
            "   for (var m = 0; m < lStrip.length; m++) if (lStrip[m].parentNode) lStrip[m].parentNode.removeChild(lStrip[m]);" +
            "   var sValue = oClone.textContent.replace(/\\s+/g, \" \").trim();" +
            "   if (sLabel != \"\" && sValue != \"\") o.fields[sLabel] = sValue;" +
            "  }" +
            "  lOut.push(o);" +
            " }" +
            " return JSON.stringify({ rows: lOut });" +
            "})()";
    }

    // Shapes one scanned Bookshare entry into a catalog row: the book id
    // and page url, the authors, and every metadata field kept generically
    // under Bookshare's own label.  Enrichment already saved for this book
    // id in the state snapshot is carried over.
    static Dictionary<string, object> bookshareRowFromItem(Dictionary<string, object> dItem, Dictionary<string, Dictionary<string, object>> dSavedById)
    {
        string sId, sName;
        Dictionary<string, object> dPair, dRow, dSaved;
        List<object> lAuthors;

        if (!dItem.ContainsKey("title") || Convert.ToString(dItem["title"]).Trim() == "") return null;
        sId = dItem.ContainsKey("id") ? Convert.ToString(dItem["id"]) : "";
        if (sId == "") return null;
        dRow = new Dictionary<string, object>();
        dRow["bsId"] = sId;
        dRow["title"] = Convert.ToString(dItem["title"]).Trim();
        dRow["bsUrl"] = "https://www.bookshare.org/browse/book/" + sId;
        dRow["bsCount"] = 1;
        lAuthors = new List<object>();
        if (dItem.ContainsKey("authors") && dItem["authors"] != null)
        {
            foreach (object oName in (IEnumerable) dItem["authors"])
            {
                sName = Convert.ToString(oName).Trim();
                if (sName == "") continue;
                dPair = new Dictionary<string, object>();
                dPair["name"] = sName;
                dPair["url"] = "";
                lAuthors.Add(dPair);
            }
        }
        dRow["authors"] = lAuthors;
        if (dItem.ContainsKey("fields")) dRow["bsFields"] = dItem["fields"];
        if (dSavedById.ContainsKey(sId))
        {
            dSaved = dSavedById[sId];
            foreach (string sField in new string[] { "firstPublished", "publisher", "wikipediaUrl", "wikipediaTitle", "openLibraryChecked", "wikipediaChecked" })
            {
                if (dSaved.ContainsKey(sField) && !dRow.ContainsKey(sField)) dRow[sField] = dSaved[sField];
            }
        }
        return dRow;
    }

    // A named Bookshare metadata field's value, or an empty string.
    static string bsField(Dictionary<string, object> dRow, string sLabel)
    {
        Dictionary<string, object> dFields;

        if (!dRow.ContainsKey("bsFields")) return "";
        dFields = dRow["bsFields"] as Dictionary<string, object>;
        if (dFields == null || !dFields.ContainsKey(sLabel)) return "";
        return Convert.ToString(dFields[sLabel]).Trim();
    }

    // After the lanes drain, every Bookshare row tied to a twin copies the
    // Builds Bookshare_Library.htm and Bookshare_Library.md in the download
    // folder, in the family shape: a table of contents, an introduction
    // with counts, every book's fields under Bookshare's own labels plus
    // the gathered details, and appendixes ending with About the Authors.
    // One book's full entry, with every field its library offers, used by
    // both the library's own document and the consolidated catalog.
    static void appendBookshareEntry(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, object> dRow)
    {
        string sId, sTitle;
        List<string> lExtraLabels;
        List<string[]> lPairs;

        Monitor.Enter(dRow);
        sId = "bs" + Convert.ToString(dRow["bsId"]);
        sTitle = catalogValue(dRow, "title");
        if (sTitle == "") { Monitor.Exit(dRow); return; }
        sbHtml.Append("<h2 id=\"" + sId + "\"><a href=\"" + htmlText(catalogValue(dRow, "bsUrl")) + "\">" + htmlText(titleWithYear(dRow)) + "</a></h2>\r\n");
        sbMd.Append("## [" + mdText(titleWithYear(dRow)) + "](" + catalogValue(dRow, "bsUrl") + ") {#" + sId + "}\r\n\r\n");
        lPairs = catalogLinks(dRow, "authors");
        appendField(sbHtml, sbMd, "By", linksHtml(lPairs), linksMd(lPairs));
        appendField(sbHtml, sbMd, "Edition", htmlText(editionText(dRow, "Bookshare")), editionText(dRow, "Bookshare"));
        appendField(sbHtml, sbMd, "First published", htmlText(catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)"), catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)");
        if (dRow.ContainsKey("bsCount") && Convert.ToInt32(dRow["bsCount"]) > 1) appendField(sbHtml, sbMd, "History entries", htmlText(Convert.ToString(dRow["bsCount"])), Convert.ToString(dRow["bsCount"]));
        appendField(sbHtml, sbMd, "Publisher", htmlText(catalogValue(dRow, "publisher")), catalogValue(dRow, "publisher"));
        appendField(sbHtml, sbMd, "Wikipedia", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "<a href=\"" + htmlText(catalogValue(dRow, "wikipediaUrl")) + "\">" + htmlText(catalogValue(dRow, "wikipediaTitle")) + "</a>", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "[" + mdText(catalogValue(dRow, "wikipediaTitle")) + "](" + catalogValue(dRow, "wikipediaUrl") + ")");
        lExtraLabels = new List<string>();
        if (dRow.ContainsKey("bsFields") && dRow["bsFields"] is Dictionary<string, object>)
        {
            foreach (KeyValuePair<string, object> oEntry in (Dictionary<string, object>) dRow["bsFields"]) lExtraLabels.Add(oEntry.Key);
        }
        lExtraLabels.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string sLabel in lExtraLabels)
        {
            appendField(sbHtml, sbMd, friendlyKindleLabel(sLabel), htmlText(bsField(dRow, sLabel)), bsField(dRow, sLabel));
        }
        Monitor.Exit(dRow);
    }

    static void buildBookshareFiles(string sDownloadDir)
    {
        string sId, sIntroText, sStats, sTitle;
        Dictionary<string, List<string[]>> dByAuthor, dByFormat, dByPublisher, dByStatus;
        List<Dictionary<string, object>> lOrdered;
        List<string> lExtraLabels;
        List<string[]> lPairs;
        StringBuilder sbHtml, sbMd;

        dByAuthor = new Dictionary<string, List<string[]>>();
        dByFormat = new Dictionary<string, List<string[]>>();
        dByPublisher = new Dictionary<string, List<string[]>>();
        dByStatus = new Dictionary<string, List<string[]>>();
        sbHtml = new StringBuilder();
        sbMd = new StringBuilder();
        lOrdered = new List<Dictionary<string, object>>(lBookshareCatalog);
        lOrdered.Sort(compareCatalogRowsByTitle);
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow)
            {
                sId = "bs" + Convert.ToString(dRow["bsId"]);
                sTitle = catalogValue(dRow, "title");
                foreach (string[] aPair in catalogLinks(dRow, "authors")) addToIndex(dByAuthor, aPair[0], sTitle, sId);
                if (bsField(dRow, "Format") != "") addToIndex(dByFormat, bsField(dRow, "Format"), sTitle, sId);
                if (bsField(dRow, "Status") != "") addToIndex(dByStatus, bsField(dRow, "Status"), sTitle, sId);
                if (catalogValue(dRow, "publisher") != "") addToIndex(dByPublisher, catalogValue(dRow, "publisher"), sTitle, sId);
            }
        }
        sIntroText = "This catalog lists every book in the Bookshare My History list, with every field the history page offers under its own label, enriched from Open Library and Wikipedia.  A book that also exists in the Audible, Kindle, or Goodreads library shares the details already gathered there, and a book downloaded more than once appears one time with its history entries counted.";
        sStats = "The history holds " + lBookshareCatalog.Count + (lBookshareCatalog.Count == 1 ? " book" : " books") + " by " + dByAuthor.Count + (dByAuthor.Count == 1 ? " author" : " authors") + ".";
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>Bookshare Library</title>\r\n</head>\r\n<body>\r\n");
        sbHtml.Append("<h1>Bookshare Library</h1>\r\n");
        sbHtml.Append("<nav aria-label=\"Table of contents\">\r\n<h2 id=\"contents\">Contents</h2>\r\n<ul>\r\n");
        sbHtml.Append("<li><a href=\"#introduction\">Introduction</a></li>\r\n<li><a href=\"#books\">Books</a>\r\n<ul>\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbHtml.Append("<li><a href=\"#bs" + Convert.ToString(dRow["bsId"]) + "\">" + htmlText(titleWithYear(dRow)) + "</a></li>\r\n"); }
        }
        sbHtml.Append("</ul>\r\n</li>\r\n");
        sbHtml.Append("<li><a href=\"#appendix-a\">Appendix A: Books by Author</a></li>\r\n<li><a href=\"#appendix-b\">Appendix B: Books by Format</a></li>\r\n<li><a href=\"#appendix-c\">Appendix C: Books by Status</a></li>\r\n<li><a href=\"#appendix-d\">Appendix D: Books by Publisher</a></li>\r\n<li><a href=\"#appendix-e\">Appendix E: About the Authors</a></li>\r\n");
        sbHtml.Append("</ul>\r\n</nav>\r\n");
        sbMd.Append("# Bookshare Library\r\n\r\n## Contents {#contents}\r\n\r\n");
        sbMd.Append("- [Introduction](#introduction)\r\n- [Books](#books)\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbMd.Append("    - [" + mdText(titleWithYear(dRow)) + "](#bs" + Convert.ToString(dRow["bsId"]) + ")\r\n"); }
        }
        sbMd.Append("- [Appendix A: Books by Author](#appendix-a)\r\n- [Appendix B: Books by Format](#appendix-b)\r\n- [Appendix C: Books by Status](#appendix-c)\r\n- [Appendix D: Books by Publisher](#appendix-d)\r\n- [Appendix E: About the Authors](#appendix-e)\r\n\r\n");
        sbHtml.Append("<h2 id=\"introduction\">Introduction</h2>\r\n<p>" + htmlText(sIntroText) + "</p>\r\n<p>" + htmlText(sStats) + "</p>\r\n");
        sbMd.Append("## Introduction {#introduction}\r\n\r\n" + sIntroText + "\r\n\r\n" + sStats + "\r\n\r\n");
        sbHtml.Append("<h2 id=\"books\">Books</h2>\r\n");
        sbMd.Append("## Books {#books}\r\n\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered) appendBookshareEntry(sbHtml, sbMd, dRow);
        appendAppendix(sbHtml, sbMd, "appendix-a", "Appendix A: Books by Author", dByAuthor, true);
        appendAppendix(sbHtml, sbMd, "appendix-b", "Appendix B: Books by Format", dByFormat);
        appendAppendix(sbHtml, sbMd, "appendix-c", "Appendix C: Books by Status", dByStatus);
        appendAppendix(sbHtml, sbMd, "appendix-d", "Appendix D: Books by Publisher", dByPublisher);
        appendAuthorsAppendix(sbHtml, sbMd, dByAuthor, new Dictionary<string, int>(), "appendix-e", "Appendix E: About the Authors", "Every author in the history, with the number of books they account for, and a biography when a reliable one was found.");
        sbHtml.Append("</body>\r\n</html>\r\n");
        File.WriteAllText(Path.Combine(sDownloadDir, "Bookshare_Library.htm"), sbHtml.ToString(), new UTF8Encoding(true));
        File.WriteAllText(Path.Combine(sDownloadDir, "Bookshare_Library.md"), sbMd.ToString(), new UTF8Encoding(true));
        log("The Bookshare catalog was saved as Bookshare_Library.htm and Bookshare_Library.md in " + sDownloadDir);
    }

    // Writes Bookshare_Library.xlsx with the family's screen-reader
    // conventions.
    static void fillBookshareSheet(ExcelPackage oPackage)
    {
        int iCol, iRow, iWidth;
        int[] aWidths;
        string[] aHeaders;
        List<Dictionary<string, object>> lSorted;
        List<string[]> lPairs;

        aHeaders = new string[] { "Title", "Book id", "Bookshare link", "By", "Date", "Format", "Status", "History entries", "First published", "Publisher", "Wikipedia" };
        lSorted = new List<Dictionary<string, object>>(lBookshareCatalog);
        lSorted.Sort(compareCatalogRowsByTitle);
        ExcelWorksheet oSheet = oPackage.Workbook.Worksheets.Add("Bookshare");
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1) oSheet.Cells[1, iCol].Value = aHeaders[iCol - 1];
        oSheet.Cells[1, 1, 1, aHeaders.Length].Style.Font.Bold = true;
        iRow = 1;
        foreach (Dictionary<string, object> dRow in lSorted)
        {
            iRow = iRow + 1;
            Monitor.Enter(dRow);
            oSheet.Cells[iRow, 1].Value = catalogValue(dRow, "title");
            oSheet.Cells[iRow, 2].Value = catalogValue(dRow, "bsId");
            oSheet.Cells[iRow, 3].Value = catalogValue(dRow, "bsUrl");
            lPairs = catalogLinks(dRow, "authors");
            oSheet.Cells[iRow, 4].Value = joinPairNames(lPairs);
            oSheet.Cells[iRow, 5].Value = bsField(dRow, "Date");
            oSheet.Cells[iRow, 6].Value = bsField(dRow, "Format");
            oSheet.Cells[iRow, 7].Value = bsField(dRow, "Status");
            if (dRow.ContainsKey("bsCount")) oSheet.Cells[iRow, 8].Value = Convert.ToInt32(dRow["bsCount"]);
            oSheet.Cells[iRow, 9].Value = catalogValue(dRow, "firstPublished");
            oSheet.Cells[iRow, 10].Value = catalogValue(dRow, "publisher");
            oSheet.Cells[iRow, 11].Value = catalogValue(dRow, "wikipediaUrl");
            Monitor.Exit(dRow);
        }
        aWidths = new int[aHeaders.Length];
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1)
        {
            aWidths[iCol - 1] = aHeaders[iCol - 1].Length;
            for (iRow = 2; iRow <= lSorted.Count + 1; iRow = iRow + 1)
            {
                iWidth = oSheet.Cells[iRow, iCol].Value == null ? 0 : Convert.ToString(oSheet.Cells[iRow, iCol].Value).Length;
                if (iWidth > aWidths[iCol - 1]) aWidths[iCol - 1] = iWidth;
            }
            if (aWidths[iCol - 1] > 40)
            {
                aWidths[iCol - 1] = 40;
                oSheet.Column(iCol).Style.WrapText = true;
            }
            oSheet.Column(iCol).Width = aWidths[iCol - 1] + 2;
        }
        oSheet.View.FreezePanes(2, 1);
        oSheet.Names.Add("ColumnTitle01", oSheet.Cells[1, 1]);
    }

    // ---- The Goodreads library ------------------------------------------
    // The My Books table at goodreads.com is server-rendered, so the walk
    // scans its rows straight from the DOM through the logged-in Edge
    // session, one hundred books per page.  Every column the table offers
    // is captured under Goodreads' own label for it, including columns the
    // user has hidden, which still arrive in the page with their values.
    // A failure anywhere in this phase is announced and the run continues.
    static async Task harvestGoodreadsAsync()
    {
        int iPage, iTotal;
        string sCurrent, sJson, sListUrl, sUserId;
        Dictionary<string, object> dItem, dReply, dRow, dSavedRow;
        Dictionary<string, Dictionary<string, object>> dByKey, dSavedById;
        object[] aRows;

        try
        {
            sProgressText = "";
            log("Goodreads library harvest starting");
            await navigate("https://www.goodreads.com/review/list?shelf=all&view=table&per_page=100");
            await Task.Delay(3000);
            sCurrent = await evaluate("location.href");
            while (sCurrent.Contains("/user/sign_in") || sCurrent.ToLower().Contains("signin"))
            {
                focusWhenShown("bookFido: log in to Goodreads");
                if (MessageBox.Show("You are not logged in to Goodreads yet.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Goodreads library this run.", "bookFido: log in to Goodreads", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Goodreads library was skipped at the login prompt"); return; }
                await navigate("https://www.goodreads.com/review/list?shelf=all&view=table&per_page=100");
                await Task.Delay(3000);
                sCurrent = await evaluate("location.href");
            }
            sUserId = "";
            foreach (System.Text.RegularExpressions.Match oMatch in System.Text.RegularExpressions.Regex.Matches(sCurrent, "/review/list/(\\d+)")) sUserId = oMatch.Groups[1].Value;
            if (sUserId == "")
            {
                sJson = await evaluate("(function () { var oLink = document.querySelector(\"link[rel='canonical']\"); return oLink ? oLink.href : \"\"; })()");
                foreach (System.Text.RegularExpressions.Match oMatch in System.Text.RegularExpressions.Regex.Matches(sJson, "/review/list/(\\d+)")) sUserId = oMatch.Groups[1].Value;
            }
            if (sUserId == "")
            {
                log("The Goodreads list address did not reveal a user id, so the harvest stopped: " + sCurrent);
                showTimedMessageBox("The Goodreads library could not be gathered this run");
                return;
            }
            sListUrl = "https://www.goodreads.com/review/list/" + sUserId + "?shelf=all&view=table&per_page=100&page=";
            dSavedById = new Dictionary<string, Dictionary<string, object>>();
            if (aSavedGoodreadsRows != null)
            {
                foreach (object oRow in aSavedGoodreadsRows)
                {
                    dSavedRow = oRow as Dictionary<string, object>;
                    if (dSavedRow != null && dSavedRow.ContainsKey("grId")) dSavedById[Convert.ToString(dSavedRow["grId"])] = dSavedRow;
                }
            }
            dByKey = crossLibraryMap();
            iTotal = 0;
            iPage = 0;
            while (true)
            {
                iPage = iPage + 1;
                if (iPage > 300) { log("The Goodreads harvest stopped at the safety cap of 300 pages"); break; }
                if (iPage > 1) await navigate(sListUrl + iPage);
                await Task.Delay(700);
                sJson = await evaluate(goodreadsScanScript());
                if (sJson == "") { log("The Goodreads page scan returned no result, so the harvest stopped at page " + iPage); break; }
                dReply = (Dictionary<string, object>) jsonCodec.DeserializeObject(sJson);
                if (iTotal == 0 && dReply.ContainsKey("total") && Convert.ToString(dReply["total"]) != "") iTotal = Convert.ToInt32(Convert.ToString(dReply["total"]).Replace(",", ""));
                aRows = dReply.ContainsKey("rows") ? dReply["rows"] as object[] : null;
                if (aRows == null || aRows.Length == 0)
                {
                    if (iPage == 1 && await loginFormShowing())
                    {
                        focusWhenShown("bookFido: log in to Goodreads");
                        if (MessageBox.Show("Goodreads is showing a sign-in form.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Goodreads library this run.", "bookFido: log in to Goodreads", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Goodreads library was skipped at the login prompt"); return; }
                        await navigate(sListUrl + "1");
                        await Task.Delay(3000);
                        iPage = 0;
                        continue;
                    }
                    break;
                }
                foreach (object oItem in aRows)
                {
                    dItem = oItem as Dictionary<string, object>;
                    if (dItem == null) continue;
                    dRow = goodreadsRowFromItem(dItem, dSavedById);
                    if (dRow == null) continue;
                    lGoodreadsCatalog.Add(dRow);
                    enqueueSharedEnrichment(dRow, dByKey);
                }
                if (iTotal > 0) sProgressText = (lGoodreadsCatalog.Count * 100 / iTotal) + "%";
                if (iPage % 3 == 0) showTimedMessageBox("Goodreads library: " + lGoodreadsCatalog.Count + " books so far");
                if (iTotal > 0 && lGoodreadsCatalog.Count >= iTotal) break;
            }
            log("Goodreads library harvest complete: " + lGoodreadsCatalog.Count + " books, of which " + countRowsWithKey(lGoodreadsCatalog, "twinKey") + " share details with another library");
            showTimedMessageBox("Goodreads library: " + lGoodreadsCatalog.Count + " books found");
            savePeriodically(true);
        }
        catch (Exception oException)
        {
            log("The Goodreads library could not be harvested, so the run continues without it: " + oException.Message);
            showTimedMessageBox("The Goodreads library could not be gathered this run");
        }
    }

    // The in-page scan of one My Books table page: every row's cells are
    // read generically, the field name from the cell's class, the display
    // label from its label element, and the cleaned text from its value
    // element; title, author, rating, and shelves get exact handling.
    static string goodreadsScanScript()
    {
        return "(function () {" +
            " var lOut = [];" +
            " var lRows = document.querySelectorAll(\"tr[id^='review_']\");" +
            " for (var i = 0; i < lRows.length; i++) {" +
            "  var oRow = lRows[i];" +
            "  var o = { fields: {} };" +
            "  var lCells = oRow.querySelectorAll(\"td\");" +
            "  for (var j = 0; j < lCells.length; j++) {" +
            "   var oCell = lCells[j];" +
            "   var sClass = (oCell.className || \"\").replace(\"field\", \"\").trim();" +
            "   if (!sClass || sClass == \"checkbox\" || sClass == \"actions\" || sClass == \"cover\" || sClass == \"position\") continue;" +
            "   var oValue = oCell.querySelector(\"div.value\");" +
            "   if (!oValue) continue;" +
            "   if (sClass == \"title\") { var oLink = oValue.querySelector(\"a\"); if (oLink) { o.title = (oLink.getAttribute(\"title\") || oLink.textContent).trim(); o.url = oLink.getAttribute(\"href\") || \"\"; } continue; }" +
            "   if (sClass == \"author\") { var oA = oValue.querySelector(\"a\"); o.author = (oA ? oA.textContent : oValue.textContent).trim(); o.authorUrl = oA ? (oA.getAttribute(\"href\") || \"\") : \"\"; continue; }" +
            "   if (sClass == \"rating\") { var oStars = oValue.querySelector(\"div.stars\"); o.myRating = oStars ? (oStars.getAttribute(\"data-rating\") || \"0\") : \"0\"; continue; }" +
            "   if (sClass == \"shelves\") { var lLinks = oValue.querySelectorAll(\"a.shelfLink\"); var lNames = []; for (var k = 0; k < lLinks.length; k++) lNames.push(lLinks[k].textContent.trim()); o.shelves = lNames; continue; }" +
            "   var oClone = oValue.cloneNode(true);" +
            "   var lStrip = oClone.querySelectorAll(\"a[href='#'], a.smallText, a.floatingBoxLink, script, style\");" +
            "   for (var m = 0; m < lStrip.length; m++) lStrip[m].parentNode.removeChild(lStrip[m]);" +
            "   var sText = oClone.textContent.replace(/\\s+/g, \" \").trim();" +
            "   if (sText == \"Write a review\" || sText == \"[edit]\" || sText == \"view (with text)\") sText = \"\";" +
            "   var oLabel = oCell.querySelector(\"label\");" +
            "   if (sText != \"\") o.fields[(oLabel ? oLabel.textContent.trim() : sClass)] = sText;" +
            "  }" +
            "  lOut.push(o);" +
            " }" +
            " var oTitleMatch = document.title.match(/\\(([0-9,]+) books?\\)/);" +
            " return JSON.stringify({ total: oTitleMatch ? oTitleMatch[1] : \"\", rows: lOut });" +
            "})()";
    }

    // Shapes one scanned Goodreads row into a catalog row: the book id and
    // absolute url, the author restored to natural order, the numeric my
    // rating, the shelf list, and every other column kept generically under
    // Goodreads' own label.  Enrichment already saved for this book id in
    // the state snapshot is carried over.
    static Dictionary<string, object> goodreadsRowFromItem(Dictionary<string, object> dItem, Dictionary<string, Dictionary<string, object>> dSavedById)
    {
        string sAuthorUrl, sId, sName, sUrl;
        Dictionary<string, object> dPair, dRow, dSaved;
        List<object> lAuthors;

        if (!dItem.ContainsKey("title") || Convert.ToString(dItem["title"]).Trim() == "") return null;
        sUrl = dItem.ContainsKey("url") ? Convert.ToString(dItem["url"]) : "";
        if (sUrl.StartsWith("/")) sUrl = "https://www.goodreads.com" + sUrl;
        sId = "";
        foreach (System.Text.RegularExpressions.Match oMatch in System.Text.RegularExpressions.Regex.Matches(sUrl, "/show/(\\d+)")) sId = oMatch.Groups[1].Value;
        if (sId == "") return null;
        dRow = new Dictionary<string, object>();
        dRow["grId"] = sId;
        dRow["title"] = Convert.ToString(dItem["title"]).Trim();
        dRow["grUrl"] = sUrl;
        sName = dItem.ContainsKey("author") ? kindleAuthorName(Convert.ToString(dItem["author"])) : "";
        sAuthorUrl = dItem.ContainsKey("authorUrl") ? Convert.ToString(dItem["authorUrl"]) : "";
        if (sAuthorUrl.StartsWith("/")) sAuthorUrl = "https://www.goodreads.com" + sAuthorUrl;
        lAuthors = new List<object>();
        if (sName != "")
        {
            dPair = new Dictionary<string, object>();
            dPair["name"] = sName;
            dPair["url"] = sAuthorUrl;
            lAuthors.Add(dPair);
        }
        dRow["authors"] = lAuthors;
        if (dItem.ContainsKey("myRating")) dRow["grMyRating"] = Convert.ToInt32(Convert.ToString(dItem["myRating"]));
        if (dItem.ContainsKey("shelves")) dRow["grShelves"] = dItem["shelves"];
        if (dItem.ContainsKey("fields")) dRow["grFields"] = dItem["fields"];
        if (dSavedById.ContainsKey(sId))
        {
            dSaved = dSavedById[sId];
            foreach (string sField in new string[] { "firstPublished", "publisher", "wikipediaUrl", "wikipediaTitle", "openLibraryChecked", "wikipediaChecked" })
            {
                if (dSaved.ContainsKey(sField) && !dRow.ContainsKey(sField)) dRow[sField] = dSaved[sField];
            }
        }
        return dRow;
    }

    // Fills a catalog list straight from the state snapshot's saved rows,
    // for a library that was not selected this run: its books still feed
    // cross-library matching and the documents, without any searching.
    static void materializeSavedRows(object[] aSaved, List<Dictionary<string, object>> lTarget)
    {
        Dictionary<string, object> dRow;

        if (aSaved == null || lTarget.Count > 0) return;
        foreach (object oRow in aSaved) { dRow = oRow as Dictionary<string, object>; if (dRow != null) lTarget.Add(dRow); }
    }

    // One map of every Audible and Kindle row by the cross-library key, the
    // Audible row preferred when both hold the same work.
    static Dictionary<string, Dictionary<string, object>> crossLibraryMap()
    {
        Dictionary<string, Dictionary<string, object>> dByKey;

        dByKey = new Dictionary<string, Dictionary<string, object>>();
        addRowsToKeyMap(dByKey, lCatalog, aSavedRows);
        addRowsToKeyMap(dByKey, lKindleCatalog, aSavedKindleRows);
        addRowsToKeyMap(dByKey, lGoodreadsCatalog, aSavedGoodreadsRows);
        addRowsToKeyMap(dByKey, lBookshareCatalog, aSavedBookshareRows);
        return dByKey;
    }

    // Adds one library's rows to the twin map: the live catalog when it has
    // been filled, otherwise the saved rows from the state snapshot, so a
    // library processed later, or not at all this run, still lends its
    // gathered details to the others.
    static void addRowsToKeyMap(Dictionary<string, Dictionary<string, object>> dByKey, List<Dictionary<string, object>> lLive, object[] aSaved)
    {
        string sKey;
        Dictionary<string, object> dRow;
        List<string[]> lPairs;

        if (lLive.Count > 0)
        {
            foreach (Dictionary<string, object> dLiveRow in lLive)
            {
                lock (dLiveRow)
                {
                    lPairs = catalogLinks(dLiveRow, "authors");
                    sKey = crossLibraryKey(catalogValue(dLiveRow, "title"), lPairs.Count > 0 ? lPairs[0][0] : "");
                    if (!dByKey.ContainsKey(sKey)) dByKey[sKey] = dLiveRow;
                }
            }
            return;
        }
        if (aSaved == null) return;
        foreach (object oRow in aSaved)
        {
            dRow = oRow as Dictionary<string, object>;
            if (dRow == null) continue;
            lPairs = catalogLinks(dRow, "authors");
            sKey = crossLibraryKey(catalogValue(dRow, "title"), lPairs.Count > 0 ? lPairs[0][0] : "");
            if (!dByKey.ContainsKey(sKey)) dByKey[sKey] = dRow;
        }
    }

    // After the lanes drain, every row tied to a twin, or matching one by
    // key, copies the twin's gathered details, whichever library the twin
    // lives in and whatever order the libraries were processed.
    static void resolveTwins()
    {
        Dictionary<string, Dictionary<string, object>> dByKey;

        dByKey = crossLibraryMap();
        resolveTwinList(lKindleCatalog, dByKey);
        resolveTwinList(lGoodreadsCatalog, dByKey);
        resolveTwinList(lBookshareCatalog, dByKey);
    }

    static void resolveTwinList(List<Dictionary<string, object>> lRows, Dictionary<string, Dictionary<string, object>> dByKey)
    {
        string sKey;
        Dictionary<string, object> dTwin;
        List<string[]> lPairs;

        foreach (Dictionary<string, object> dRow in lRows)
        {
            lock (dRow)
            {
                lPairs = catalogLinks(dRow, "authors");
                sKey = dRow.ContainsKey("twinKey") ? Convert.ToString(dRow["twinKey"]) : crossLibraryKey(catalogValue(dRow, "title"), lPairs.Count > 0 ? lPairs[0][0] : "");
            }
            if (!dByKey.ContainsKey(sKey)) continue;
            dTwin = dByKey[sKey];
            if (object.ReferenceEquals(dTwin, dRow)) continue;
            lock (dTwin)
            {
                foreach (string sField in new string[] { "firstPublished", "publisher", "wikipediaUrl", "wikipediaTitle" })
                {
                    if (dTwin.ContainsKey(sField) && !dRow.ContainsKey(sField)) dRow[sField] = dTwin[sField];
                }
            }
        }
    }

    // Hands a Goodreads row to the Open Library and Wikipedia lanes, unless
    // the same work already sits in the Audible or Kindle catalog, in which
    // case the row is tied by key and every gathered detail is copied after
    // the lanes drain instead of requested again.  Authors dedupe through
    // the shared sets, so no biography is ever fetched twice.
    static void enqueueSharedEnrichment(Dictionary<string, object> dRow, Dictionary<string, Dictionary<string, object>> dByKey)
    {
        string sKey;
        List<string[]> lAuthorPairs;

        lAuthorPairs = catalogLinks(dRow, "authors");
        sKey = crossLibraryKey(catalogValue(dRow, "title"), lAuthorPairs.Count > 0 ? lAuthorPairs[0][0] : "");
        if (dByKey.ContainsKey(sKey))
        {
            dRow["twinKey"] = sKey;
        }
        else
        {
            if (!dRow.ContainsKey("openLibraryChecked")) { lock (queueOpenLibrary) { queueOpenLibrary.Enqueue(dRow); } iWorkTotal = iWorkTotal + 1; }
            if (!dRow.ContainsKey("wikipediaChecked")) { lock (queueWikipedia) { queueWikipedia.Enqueue(dRow); } iWorkTotal = iWorkTotal + 1; }
        }
        foreach (string[] aPair in lAuthorPairs)
        {
            if (aPair[0] == "" || !setAuthorsQueued.Add(aPair[0])) continue;
            if (!setAuthorsCheckedWiki.Contains(aPair[0])) { lock (queueAuthorWikipedia) { queueAuthorWikipedia.Enqueue(aPair[0]); } iWorkTotal = iWorkTotal + 1; }
            if (!setAuthorsCheckedOl.Contains(aPair[0])) { lock (queueAuthorOpenLibrary) { queueAuthorOpenLibrary.Enqueue(aPair[0]); } iWorkTotal = iWorkTotal + 1; }
        }
    }

    // After the lanes drain, every Goodreads row tied to a twin copies the
    // A named Goodreads column's value, or an empty string.
    static string grField(Dictionary<string, object> dRow, string sLabel)
    {
        Dictionary<string, object> dFields;

        if (!dRow.ContainsKey("grFields")) return "";
        dFields = dRow["grFields"] as Dictionary<string, object>;
        if (dFields == null || !dFields.ContainsKey(sLabel)) return "";
        return Convert.ToString(dFields[sLabel]).Trim();
    }

    static string grMyRatingText(Dictionary<string, object> dRow)
    {
        int iStars;

        if (!dRow.ContainsKey("grMyRating")) return "";
        iStars = Convert.ToInt32(dRow["grMyRating"]);
        if (iStars <= 0) return "Unrated";
        return iStars + " of 5 stars";
    }

    // Builds Goodreads_Library.htm and Goodreads_Library.md in the download
    // folder: a table of contents listing every book, an introduction with
    // counts, every book's columns under Goodreads' own labels plus the
    // gathered details, and appendixes ending with About the Authors.
    // One book's full entry, with every field its library offers, used by
    // both the library's own document and the consolidated catalog.
    static void appendGoodreadsEntry(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, object> dRow)
    {
        string sId, sTitle;
        List<string> lExtraLabels;
        List<string[]> lPairs;

        Monitor.Enter(dRow);
        sId = "gr" + Convert.ToString(dRow["grId"]);
        sTitle = catalogValue(dRow, "title");
        if (sTitle == "") { Monitor.Exit(dRow); return; }
        sbHtml.Append("<h2 id=\"" + sId + "\"><a href=\"" + htmlText(catalogValue(dRow, "grUrl")) + "\">" + htmlText(titleWithYear(dRow)) + "</a></h2>\r\n");
        sbMd.Append("## [" + mdText(titleWithYear(dRow)) + "](" + catalogValue(dRow, "grUrl") + ") {#" + sId + "}\r\n\r\n");
        lPairs = catalogLinks(dRow, "authors");
        appendField(sbHtml, sbMd, "By", linksHtml(lPairs), linksMd(lPairs));
        appendField(sbHtml, sbMd, "Edition", htmlText(editionText(dRow, "Goodreads")), editionText(dRow, "Goodreads"));
        appendField(sbHtml, sbMd, "First published", htmlText(catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)"), catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)");
        appendField(sbHtml, sbMd, "My rating", htmlText(grMyRatingText(dRow)), grMyRatingText(dRow));
        appendField(sbHtml, sbMd, "Publisher", htmlText(catalogValue(dRow, "publisher")), catalogValue(dRow, "publisher"));
        appendField(sbHtml, sbMd, "Shelves", htmlText(catalogStrings(dRow, "grShelves").Count == 0 ? "" : string.Join(", ", catalogStrings(dRow, "grShelves").ToArray())), catalogStrings(dRow, "grShelves").Count == 0 ? "" : string.Join(", ", catalogStrings(dRow, "grShelves").ToArray()));
        appendField(sbHtml, sbMd, "Wikipedia", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "<a href=\"" + htmlText(catalogValue(dRow, "wikipediaUrl")) + "\">" + htmlText(catalogValue(dRow, "wikipediaTitle")) + "</a>", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "[" + mdText(catalogValue(dRow, "wikipediaTitle")) + "](" + catalogValue(dRow, "wikipediaUrl") + ")");
        lExtraLabels = new List<string>();
        if (dRow.ContainsKey("grFields") && dRow["grFields"] is Dictionary<string, object>)
        {
            foreach (KeyValuePair<string, object> oEntry in (Dictionary<string, object>) dRow["grFields"]) lExtraLabels.Add(oEntry.Key);
        }
        lExtraLabels.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string sLabel in lExtraLabels)
        {
            appendField(sbHtml, sbMd, friendlyKindleLabel(sLabel), htmlText(grField(dRow, sLabel)), grField(dRow, sLabel));
        }
        Monitor.Exit(dRow);
    }

    static void buildGoodreadsFiles(string sDownloadDir)
    {
        int iRatedCount;
        string sId, sIntroText, sRatingText, sStats, sTitle;
        Dictionary<string, List<string[]>> dByAuthor, dByPublisher, dByRating, dByShelf;
        List<Dictionary<string, object>> lOrdered;
        List<string> lExtraLabels;
        List<string[]> lPairs;
        StringBuilder sbHtml, sbMd;

        dByAuthor = new Dictionary<string, List<string[]>>();
        dByPublisher = new Dictionary<string, List<string[]>>();
        dByRating = new Dictionary<string, List<string[]>>();
        dByShelf = new Dictionary<string, List<string[]>>();
        sbHtml = new StringBuilder();
        sbMd = new StringBuilder();
        iRatedCount = 0;
        lOrdered = new List<Dictionary<string, object>>(lGoodreadsCatalog);
        lOrdered.Sort(compareCatalogRowsByTitle);
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow)
            {
                sId = "gr" + Convert.ToString(dRow["grId"]);
                sTitle = catalogValue(dRow, "title");
                foreach (string[] aPair in catalogLinks(dRow, "authors")) addToIndex(dByAuthor, aPair[0], sTitle, sId);
                foreach (string sShelf in catalogStrings(dRow, "grShelves")) addToIndex(dByShelf, sShelf, sTitle, sId);
                sRatingText = grMyRatingText(dRow);
                if (sRatingText != "") addToIndex(dByRating, sRatingText, sTitle, sId);
                if (sRatingText != "" && sRatingText != "Unrated") iRatedCount = iRatedCount + 1;
                if (catalogValue(dRow, "publisher") != "") addToIndex(dByPublisher, catalogValue(dRow, "publisher"), sTitle, sId);
            }
        }
        sIntroText = "This catalog lists every book on the Goodreads My Books shelves, with every column the Goodreads table offers under its own label, enriched from Open Library and Wikipedia.  A book that also exists in the Audible or Kindle library shares the details already gathered there.";
        sStats = "The shelves hold " + lGoodreadsCatalog.Count + (lGoodreadsCatalog.Count == 1 ? " book" : " books") + " by " + dByAuthor.Count + (dByAuthor.Count == 1 ? " author" : " authors") + ", across " + dByShelf.Count + (dByShelf.Count == 1 ? " shelf" : " shelves") + ", with " + iRatedCount + " rated.";
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>Goodreads Library</title>\r\n</head>\r\n<body>\r\n");
        sbHtml.Append("<h1>Goodreads Library</h1>\r\n");
        sbHtml.Append("<nav aria-label=\"Table of contents\">\r\n<h2 id=\"contents\">Contents</h2>\r\n<ul>\r\n");
        sbHtml.Append("<li><a href=\"#introduction\">Introduction</a></li>\r\n<li><a href=\"#books\">Books</a>\r\n<ul>\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbHtml.Append("<li><a href=\"#gr" + Convert.ToString(dRow["grId"]) + "\">" + htmlText(titleWithYear(dRow)) + "</a></li>\r\n"); }
        }
        sbHtml.Append("</ul>\r\n</li>\r\n");
        sbHtml.Append("<li><a href=\"#appendix-a\">Appendix A: Books by Author</a></li>\r\n<li><a href=\"#appendix-b\">Appendix B: Books by Shelf</a></li>\r\n<li><a href=\"#appendix-c\">Appendix C: Books by My Rating</a></li>\r\n<li><a href=\"#appendix-d\">Appendix D: Books by Publisher</a></li>\r\n<li><a href=\"#appendix-e\">Appendix E: About the Authors</a></li>\r\n");
        sbHtml.Append("</ul>\r\n</nav>\r\n");
        sbMd.Append("# Goodreads Library\r\n\r\n## Contents {#contents}\r\n\r\n");
        sbMd.Append("- [Introduction](#introduction)\r\n- [Books](#books)\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbMd.Append("    - [" + mdText(titleWithYear(dRow)) + "](#gr" + Convert.ToString(dRow["grId"]) + ")\r\n"); }
        }
        sbMd.Append("- [Appendix A: Books by Author](#appendix-a)\r\n- [Appendix B: Books by Shelf](#appendix-b)\r\n- [Appendix C: Books by My Rating](#appendix-c)\r\n- [Appendix D: Books by Publisher](#appendix-d)\r\n- [Appendix E: About the Authors](#appendix-e)\r\n\r\n");
        sbHtml.Append("<h2 id=\"introduction\">Introduction</h2>\r\n<p>" + htmlText(sIntroText) + "</p>\r\n<p>" + htmlText(sStats) + "</p>\r\n");
        sbMd.Append("## Introduction {#introduction}\r\n\r\n" + sIntroText + "\r\n\r\n" + sStats + "\r\n\r\n");
        sbHtml.Append("<h2 id=\"books\">Books</h2>\r\n");
        sbMd.Append("## Books {#books}\r\n\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered) appendGoodreadsEntry(sbHtml, sbMd, dRow);
        appendAppendix(sbHtml, sbMd, "appendix-a", "Appendix A: Books by Author", dByAuthor, true);
        appendAppendix(sbHtml, sbMd, "appendix-b", "Appendix B: Books by Shelf", dByShelf);
        appendAppendix(sbHtml, sbMd, "appendix-c", "Appendix C: Books by My Rating", dByRating);
        appendAppendix(sbHtml, sbMd, "appendix-d", "Appendix D: Books by Publisher", dByPublisher);
        appendAuthorsAppendix(sbHtml, sbMd, dByAuthor, new Dictionary<string, int>(), "appendix-e", "Appendix E: About the Authors", "Every author on the shelves, with the number of books they account for, and a biography when a reliable one was found.");
        sbHtml.Append("</body>\r\n</html>\r\n");
        File.WriteAllText(Path.Combine(sDownloadDir, "Goodreads_Library.htm"), sbHtml.ToString(), new UTF8Encoding(true));
        File.WriteAllText(Path.Combine(sDownloadDir, "Goodreads_Library.md"), sbMd.ToString(), new UTF8Encoding(true));
        log("The Goodreads catalog was saved as Goodreads_Library.htm and Goodreads_Library.md in " + sDownloadDir);
    }

    // Writes Goodreads_Library.xlsx with the same screen-reader conventions
    // as the other spreadsheets.
    static void fillGoodreadsSheet(ExcelPackage oPackage)
    {
        int iCol, iRow, iWidth;
        string sPages;
        int[] aWidths;
        string[] aHeaders;
        List<Dictionary<string, object>> lSorted;
        List<string[]> lPairs;

        aHeaders = new string[] { "Title", "Goodreads link", "By", "Avg rating", "My rating", "Shelves", "Date read", "Date added", "Date published", "Format", "ISBN", "ISBN13", "Pages", "First published", "Publisher", "Wikipedia" };
        lSorted = new List<Dictionary<string, object>>(lGoodreadsCatalog);
        lSorted.Sort(compareCatalogRowsByTitle);
        ExcelWorksheet oSheet = oPackage.Workbook.Worksheets.Add("Goodreads");
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1) oSheet.Cells[1, iCol].Value = aHeaders[iCol - 1];
        oSheet.Cells[1, 1, 1, aHeaders.Length].Style.Font.Bold = true;
        iRow = 1;
        foreach (Dictionary<string, object> dRow in lSorted)
        {
            iRow = iRow + 1;
            Monitor.Enter(dRow);
            oSheet.Cells[iRow, 1].Value = catalogValue(dRow, "title");
            oSheet.Cells[iRow, 2].Value = catalogValue(dRow, "grUrl");
            lPairs = catalogLinks(dRow, "authors");
            oSheet.Cells[iRow, 3].Value = joinPairNames(lPairs);
            if (grField(dRow, "avg rating") != "") { try { oSheet.Cells[iRow, 4].Value = Convert.ToDouble(grField(dRow, "avg rating")); } catch (Exception) { oSheet.Cells[iRow, 4].Value = grField(dRow, "avg rating"); } }
            if (dRow.ContainsKey("grMyRating") && Convert.ToInt32(dRow["grMyRating"]) > 0) oSheet.Cells[iRow, 5].Value = Convert.ToInt32(dRow["grMyRating"]);
            oSheet.Cells[iRow, 6].Value = string.Join(", ", catalogStrings(dRow, "grShelves").ToArray());
            oSheet.Cells[iRow, 7].Value = grField(dRow, "date read");
            oSheet.Cells[iRow, 8].Value = grField(dRow, "date added");
            oSheet.Cells[iRow, 9].Value = grField(dRow, "date pub");
            oSheet.Cells[iRow, 10].Value = grField(dRow, "format");
            oSheet.Cells[iRow, 11].Value = grField(dRow, "isbn");
            oSheet.Cells[iRow, 12].Value = grField(dRow, "isbn13");
            sPages = grField(dRow, "num pages").Replace("pp", "").Trim();
            if (sPages != "") { try { oSheet.Cells[iRow, 13].Value = Convert.ToInt32(sPages); } catch (Exception) { oSheet.Cells[iRow, 13].Value = sPages; } }
            oSheet.Cells[iRow, 14].Value = catalogValue(dRow, "firstPublished");
            oSheet.Cells[iRow, 15].Value = catalogValue(dRow, "publisher");
            oSheet.Cells[iRow, 16].Value = catalogValue(dRow, "wikipediaUrl");
            Monitor.Exit(dRow);
        }
        aWidths = new int[aHeaders.Length];
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1)
        {
            aWidths[iCol - 1] = aHeaders[iCol - 1].Length;
            for (iRow = 2; iRow <= lSorted.Count + 1; iRow = iRow + 1)
            {
                iWidth = oSheet.Cells[iRow, iCol].Value == null ? 0 : Convert.ToString(oSheet.Cells[iRow, iCol].Value).Length;
                if (iWidth > aWidths[iCol - 1]) aWidths[iCol - 1] = iWidth;
            }
            if (aWidths[iCol - 1] > 40)
            {
                aWidths[iCol - 1] = 40;
                oSheet.Column(iCol).Style.WrapText = true;
            }
            oSheet.Column(iCol).Width = aWidths[iCol - 1] + 2;
        }
        oSheet.View.FreezePanes(2, 1);
        oSheet.Names.Add("ColumnTitle01", oSheet.Cells[1, 1]);
    }

    // ---- The Kindle library ---------------------------------------------
    // The Kindle catalog comes from the JSON search endpoint behind the
    // read.amazon.com library page, fetched inside the logged-in Edge
    // session, so every field the Kindle user interface offers arrives
    // structured rather than scraped.  A failure anywhere in this phase is
    // announced and the run continues with the Audible catalog alone.
    static async Task harvestKindleAsync()
    {
        int iBatch;
        string sCurrent, sJson, sScript, sToken;
        Dictionary<string, object> dItem, dReply, dRow, dSavedRow;
        Dictionary<string, Dictionary<string, object>> dByKey, dSavedByAsin;

        try
        {
            sProgressText = "";
            log("Kindle library harvest starting");
            await navigate(sKindleLibraryUrl);
            await Task.Delay(4000);
            sCurrent = await evaluate("location.href");
            while (sCurrent.Contains("/ap/") || sCurrent.ToLower().Contains("signin"))
            {
                focusWhenShown("bookFido: log in to Amazon");
                if (MessageBox.Show("You are not logged in to Amazon yet.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Kindle library this run.", "bookFido: log in to Amazon", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Kindle library was skipped at the login prompt"); return; }
                await navigate(sKindleLibraryUrl);
                await Task.Delay(4000);
                sCurrent = await evaluate("location.href");
            }
            dSavedByAsin = new Dictionary<string, Dictionary<string, object>>();
            if (aSavedKindleRows != null)
            {
                foreach (object oRow in aSavedKindleRows)
                {
                    dSavedRow = oRow as Dictionary<string, object>;
                    if (dSavedRow != null && dSavedRow.ContainsKey("asin")) dSavedByAsin[Convert.ToString(dSavedRow["asin"])] = dSavedRow;
                }
            }
            dByKey = crossLibraryMap();
            sToken = "";
            iBatch = 0;
            while (true)
            {
                iBatch = iBatch + 1;
                if (iBatch > 400) { log("The Kindle harvest stopped at the safety cap of 400 batches"); break; }
                sScript = "(async function () { try { var oResponse = await fetch(\"" + sKindleSearchUrl + (sToken == "" ? "" : "&paginationToken=" + sToken) + "\", { credentials: \"include\" }); if (!oResponse.ok) return \"HTTPERROR \" + oResponse.status; return await oResponse.text(); } catch (oError) { return \"FETCHERROR \" + oError; } })()";
                sJson = await evaluate(sScript);
                if (sJson.StartsWith("HTTPERROR 401") || sJson.StartsWith("HTTPERROR 403") || (sJson == "" && iBatch == 1 && await loginFormShowing()))
                {
                    // Amazon sometimes serves the library shell without a
                    // redirect and only the data fetch reveals the missing
                    // login, so an authorization failure asks for one.
                    focusWhenShown("bookFido: log in to Amazon");
                    if (MessageBox.Show("Amazon has not accepted the session as logged in.  Log in within the Edge window that is open, then choose OK to continue.  Or choose Cancel to skip the Kindle library this run.", "bookFido: log in to Amazon", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.Cancel) { log("The Kindle library was skipped at the login prompt"); return; }
                    await navigate(sKindleLibraryUrl);
                    await Task.Delay(4000);
                    iBatch = iBatch - 1;
                    continue;
                }
                if (sJson == "" || sJson.StartsWith("HTTPERROR") || sJson.StartsWith("FETCHERROR"))
                {
                    log("The Kindle library endpoint did not answer as expected: " + (sJson == "" ? "empty reply" : sJson));
                    showTimedMessageBox("The Kindle library could not be gathered this run");
                    return;
                }
                dReply = (Dictionary<string, object>) jsonCodec.DeserializeObject(sJson);
                if (!dReply.ContainsKey("itemsList") || dReply["itemsList"] == null)
                {
                    log("The Kindle library reply held no itemsList, so the harvest stopped: " + (sJson.Length > 300 ? sJson.Substring(0, 300) : sJson));
                    showTimedMessageBox("The Kindle library could not be gathered this run");
                    return;
                }
                foreach (object oItem in (IEnumerable) dReply["itemsList"])
                {
                    dItem = oItem as Dictionary<string, object>;
                    if (dItem == null) continue;
                    dRow = kindleRowFromItem(dItem, dSavedByAsin);
                    lKindleCatalog.Add(dRow);
                    enqueueSharedEnrichment(dRow, dByKey);
                }
                if (iBatch % 5 == 0) showTimedMessageBox("Kindle library: " + lKindleCatalog.Count + " books so far");
                sToken = dReply.ContainsKey("paginationToken") && dReply["paginationToken"] != null ? Convert.ToString(dReply["paginationToken"]) : "";
                if (sToken == "") break;
                await Task.Delay(400);
            }
            bKindleHarvested = true;
            log("Kindle library harvest complete: " + lKindleCatalog.Count + " books, of which " + countRowsWithKey(lKindleCatalog, "twinKey") + " share details with the Audible catalog");
            showTimedMessageBox("Kindle library: " + lKindleCatalog.Count + " books found");
            savePeriodically(true);
        }
        catch (Exception oException)
        {
            log("The Kindle library could not be harvested, so the run continues with Audible only: " + oException.Message);
            showTimedMessageBox("The Kindle library could not be gathered this run");
        }
    }

    // Shapes one Kindle library item into a catalog row.  Every scalar field
    // the endpoint offers is kept under its own name, so nothing the Kindle
    // user interface knows is dropped; authors are normalized into the same
    // name-and-url pairs the Audible rows use, and enrichment already saved
    // for this asin in the state snapshot is carried over.
    static Dictionary<string, object> kindleRowFromItem(Dictionary<string, object> dItem, Dictionary<string, Dictionary<string, object>> dSavedByAsin)
    {
        string sAsin, sName;
        Dictionary<string, object> dPair, dRow, dSaved;
        List<object> lAuthors;

        dRow = new Dictionary<string, object>();
        foreach (KeyValuePair<string, object> oEntry in dItem)
        {
            if (oEntry.Value == null || oEntry.Key == "authors") continue;
            if (oEntry.Value is object[] || oEntry.Value is Dictionary<string, object>) continue;
            dRow[oEntry.Key] = oEntry.Value;
        }
        lAuthors = new List<object>();
        if (dItem.ContainsKey("authors") && dItem["authors"] != null)
        {
            // The endpoint joins all of a book's authors into one string
            // with colons, sometimes repeating each name, and occasionally
            // leaking non-name text; each segment is restored to natural
            // order, duplicates within the book are dropped, and segments
            // that are plainly not names are discarded.
            HashSet<string> setSeen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object oName in (IEnumerable) dItem["authors"])
            {
                foreach (string sPart in Convert.ToString(oName).Split(':'))
                {
                    sName = kindleAuthorName(sPart);
                    if (sName == "" || sName.Length > 50 || sName.Contains("?")) continue;
                    if (!setSeen.Add(sName)) continue;
                    dPair = new Dictionary<string, object>();
                    dPair["name"] = sName;
                    dPair["url"] = "";
                    lAuthors.Add(dPair);
                }
            }
        }
        dRow["authors"] = lAuthors;
        sAsin = dRow.ContainsKey("asin") ? Convert.ToString(dRow["asin"]) : "";
        if (sAsin != "" && dSavedByAsin.ContainsKey(sAsin))
        {
            dSaved = dSavedByAsin[sAsin];
            foreach (string sField in new string[] { "firstPublished", "publisher", "wikipediaUrl", "wikipediaTitle", "openLibraryChecked", "wikipediaChecked" })
            {
                if (dSaved.ContainsKey(sField) && !dRow.ContainsKey(sField)) dRow[sField] = dSaved[sField];
            }
        }
        return dRow;
    }

    // The Kindle endpoint reports an author as "Last, First:"; this strips
    // the trailing colon and restores natural order, leaving names with no
    // comma or several commas untouched.
    static string kindleAuthorName(string sRaw)
    {
        int iComma;
        string sName;

        sName = sRaw == null ? "" : sRaw.Trim();
        while (sName.EndsWith(":")) sName = sName.Substring(0, sName.Length - 1).Trim();
        iComma = sName.IndexOf(", ");
        if (iComma > 0 && sName.IndexOf(", ", iComma + 2) < 0) sName = (sName.Substring(iComma + 2) + " " + sName.Substring(0, iComma)).Trim();
        return sName;
    }

    // The key that recognizes the same work across the two libraries: the
    // title with any leading article ignored, joined with the first
    // author's name, case folded.
    static string crossLibraryKey(string sTitle, string sAuthor)
    {
        int iColon;
        string sMain;

        sMain = sTitle == null ? "" : sTitle;
        iColon = sMain.IndexOf(":");
        if (iColon > 0) sMain = sMain.Substring(0, iColon);
        return (titleSortKey(sMain.Trim()) + "|" + sAuthor).ToLowerInvariant();
    }

    // Hands a Kindle row to the Open Library and Wikipedia lanes, unless the
    // same work already sits in the Audible catalog, in which case the row
    // is tied to its Audible twin and every gathered detail is copied later
    // instead of requested again.  Authors dedupe through the same shared
    // sets the Audible rows use, so a shared author's biography is fetched
    // After the lanes drain, every Kindle row tied to an Audible twin copies
    // the twin's gathered details, so shared works read identically in both
    // Friendly wording for the endpoint's format and origin codes; an
    // unrecognized code passes through readably rather than vanishing.
    static string kindleFormatText(string sCode)
    {
        if (sCode == "EBOOK") return "Book";
        if (sCode == "EBOOK_SAMPLE") return "Sample";
        return friendlyKindleLabel(sCode);
    }

    static string kindleOriginText(string sCode)
    {
        if (sCode == "PURCHASE") return "Purchased";
        if (sCode == "KINDLE_UNLIMITED" || sCode == "KU") return "Kindle Unlimited";
        if (sCode == "PRIME") return "Prime Reading";
        if (sCode == "COMIXOLOGY_UNLIMITED") return "Comixology Unlimited";
        if (sCode == "FREE") return "Free";
        return friendlyKindleLabel(sCode);
    }

    static string kindleProgressText(Dictionary<string, object> dRow)
    {
        int iPercent;

        if (!dRow.ContainsKey("percentageRead")) return "";
        iPercent = Convert.ToInt32(Convert.ToDouble(dRow["percentageRead"]));
        if (iPercent <= 0) return "Unread";
        if (iPercent >= 100) return "Finished";
        return iPercent + " percent read";
    }

    // Turns an endpoint field name such as originType or RESOURCE_TYPE into
    // readable words for a label or a value.
    static string friendlyKindleLabel(string sName)
    {
        int iAt;
        string sOut;
        StringBuilder sbOut;

        sOut = sName == null ? "" : sName.Trim();
        if (sOut == "") return "";
        if (sOut.Contains("_")) sOut = sOut.Replace("_", " ").ToLowerInvariant();
        sbOut = new StringBuilder();
        for (iAt = 0; iAt < sOut.Length; iAt = iAt + 1)
        {
            if (iAt > 0 && char.IsUpper(sOut[iAt]) && char.IsLower(sOut[iAt - 1])) sbOut.Append(' ');
            sbOut.Append(iAt == 0 ? char.ToUpper(sOut[iAt]) : char.ToLower(sOut[iAt]));
        }
        return sbOut.ToString();
    }

    // Builds Kindle_Library.htm and Kindle_Library.md in the download
    // folder, in the same shape as the Audible catalog: a table of contents
    // listing every book, an introduction with counts, every book's fields,
    // and appendixes, ending with the shared About the Authors.
    // One book's full entry, with every field its library offers, used by
    // both the library's own document and the consolidated catalog.
    static void appendKindleEntry(StringBuilder sbHtml, StringBuilder sbMd, Dictionary<string, object> dRow)
    {
        string sAmazonUrl, sAsin, sProgress, sTitle;
        List<string> lExtraKeys;
        List<string[]> lPairs;

        Monitor.Enter(dRow);
        sAsin = catalogValue(dRow, "asin");
        sTitle = catalogValue(dRow, "title");
        if (sTitle == "") { Monitor.Exit(dRow); return; }
        sAmazonUrl = "https://www.amazon.com/dp/" + sAsin;
        sbHtml.Append("<h2 id=\"" + sAsin + "\"><a href=\"" + htmlText(sAmazonUrl) + "\">" + htmlText(titleWithYear(dRow)) + "</a></h2>\r\n");
        sbMd.Append("## [" + mdText(titleWithYear(dRow)) + "](" + sAmazonUrl + ") {#" + sAsin.ToLower() + "}\r\n\r\n");
        lPairs = catalogLinks(dRow, "authors");
        appendField(sbHtml, sbMd, "By", linksHtml(lPairs), linksMd(lPairs));
        appendField(sbHtml, sbMd, "Edition", htmlText(editionText(dRow, "Kindle")), editionText(dRow, "Kindle"));
        appendField(sbHtml, sbMd, "First published", htmlText(catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)"), catalogValue(dRow, "firstPublished") == "" ? "" : catalogValue(dRow, "firstPublished") + " (Open Library)");
        appendField(sbHtml, sbMd, "Format", htmlText(kindleFormatText(catalogValue(dRow, "resourceType"))), kindleFormatText(catalogValue(dRow, "resourceType")));
        appendField(sbHtml, sbMd, "Origin", htmlText(kindleOriginText(catalogValue(dRow, "originType"))), kindleOriginText(catalogValue(dRow, "originType")));
        sProgress = kindleProgressText(dRow);
        appendField(sbHtml, sbMd, "Progress", htmlText(sProgress), sProgress);
        appendField(sbHtml, sbMd, "Publisher", htmlText(catalogValue(dRow, "publisher")), catalogValue(dRow, "publisher"));
        appendField(sbHtml, sbMd, "Read online", catalogValue(dRow, "webReaderUrl") == "" ? "" : "<a href=\"" + htmlText(catalogValue(dRow, "webReaderUrl")) + "\">Kindle Cloud Reader</a>", catalogValue(dRow, "webReaderUrl") == "" ? "" : "[Kindle Cloud Reader](" + catalogValue(dRow, "webReaderUrl") + ")");
        appendField(sbHtml, sbMd, "Wikipedia", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "<a href=\"" + htmlText(catalogValue(dRow, "wikipediaUrl")) + "\">" + htmlText(catalogValue(dRow, "wikipediaTitle")) + "</a>", catalogValue(dRow, "wikipediaUrl") == "" ? "" : "[" + mdText(catalogValue(dRow, "wikipediaTitle")) + "](" + catalogValue(dRow, "wikipediaUrl") + ")");
        lExtraKeys = new List<string>();
        foreach (KeyValuePair<string, object> oEntry in dRow)
        {
            if (oEntry.Value == null || oEntry.Value is List<object> || oEntry.Value is object[] || oEntry.Value is Dictionary<string, object>) continue;
            if ("|asin|title|authors|percentageRead|resourceType|originType|productUrl|webReaderUrl|firstPublished|publisher|wikipediaUrl|wikipediaTitle|openLibraryChecked|wikipediaChecked|twinAsin|".Contains("|" + oEntry.Key + "|")) continue;
            lExtraKeys.Add(oEntry.Key);
        }
        lExtraKeys.Sort(StringComparer.OrdinalIgnoreCase);
        foreach (string sExtraKey in lExtraKeys)
        {
            appendField(sbHtml, sbMd, friendlyKindleLabel(sExtraKey), htmlText(Convert.ToString(dRow[sExtraKey])), Convert.ToString(dRow[sExtraKey]));
        }
        Monitor.Exit(dRow);
    }

    static void buildKindleFiles(string sDownloadDir)
    {
        int iFinishedCount, iSampleCount;
        string sAmazonUrl, sAsin, sIntroText, sProgress, sStats, sTitle;
        Dictionary<string, List<string[]>> dByAuthor, dByOrigin, dByProgress, dByPublisher;
        List<Dictionary<string, object>> lOrdered;
        List<string> lExtraKeys;
        List<string[]> lPairs;
        StringBuilder sbHtml, sbMd;

        dByAuthor = new Dictionary<string, List<string[]>>();
        dByOrigin = new Dictionary<string, List<string[]>>();
        dByProgress = new Dictionary<string, List<string[]>>();
        dByPublisher = new Dictionary<string, List<string[]>>();
        sbHtml = new StringBuilder();
        sbMd = new StringBuilder();
        iFinishedCount = 0;
        iSampleCount = 0;
        lOrdered = new List<Dictionary<string, object>>(lKindleCatalog);
        lOrdered.Sort(compareCatalogRowsByTitle);
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow)
            {
                sAsin = catalogValue(dRow, "asin");
                sTitle = catalogValue(dRow, "title");
                foreach (string[] aPair in catalogLinks(dRow, "authors")) addToIndex(dByAuthor, aPair[0], sTitle, sAsin);
                addToIndex(dByOrigin, kindleOriginText(catalogValue(dRow, "originType")), sTitle, sAsin);
                sProgress = kindleProgressText(dRow);
                addToIndex(dByProgress, sProgress == "" ? "Unknown" : (sProgress == "Unread" || sProgress == "Finished" ? sProgress : "In progress"), sTitle, sAsin);
                if (catalogValue(dRow, "publisher") != "") addToIndex(dByPublisher, catalogValue(dRow, "publisher"), sTitle, sAsin);
                if (sProgress == "Finished") iFinishedCount = iFinishedCount + 1;
                if (catalogValue(dRow, "resourceType") == "EBOOK_SAMPLE") iSampleCount = iSampleCount + 1;
            }
        }
        sIntroText = "This catalog lists every Kindle book on the Amazon account, gathered from the Kindle library service with every field its own user interface offers, and enriched from Open Library and Wikipedia.  A book that also exists in the Audible library shares the details already gathered there.";
        sStats = "The library holds " + lKindleCatalog.Count + (lKindleCatalog.Count == 1 ? " book" : " books") + (iSampleCount > 0 ? ", of which " + iSampleCount + (iSampleCount == 1 ? " is a sample" : " are samples") : "") + ", by " + dByAuthor.Count + (dByAuthor.Count == 1 ? " author" : " authors") + ", with " + iFinishedCount + " finished.";
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>Kindle Library</title>\r\n</head>\r\n<body>\r\n");
        sbHtml.Append("<h1>Kindle Library</h1>\r\n");
        sbHtml.Append("<nav aria-label=\"Table of contents\">\r\n<h2 id=\"contents\">Contents</h2>\r\n<ul>\r\n");
        sbHtml.Append("<li><a href=\"#introduction\">Introduction</a></li>\r\n<li><a href=\"#books\">Books</a>\r\n<ul>\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbHtml.Append("<li><a href=\"#" + catalogValue(dRow, "asin") + "\">" + htmlText(titleWithYear(dRow)) + "</a></li>\r\n"); }
        }
        sbHtml.Append("</ul>\r\n</li>\r\n");
        sbHtml.Append("<li><a href=\"#appendix-a\">Appendix A: Books by Author</a></li>\r\n<li><a href=\"#appendix-b\">Appendix B: Books by Origin</a></li>\r\n<li><a href=\"#appendix-c\">Appendix C: Books by Reading Progress</a></li>\r\n<li><a href=\"#appendix-d\">Appendix D: Books by Publisher</a></li>\r\n<li><a href=\"#appendix-e\">Appendix E: About the Authors</a></li>\r\n");
        sbHtml.Append("</ul>\r\n</nav>\r\n");
        sbMd.Append("# Kindle Library\r\n\r\n## Contents {#contents}\r\n\r\n");
        sbMd.Append("- [Introduction](#introduction)\r\n- [Books](#books)\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            lock (dRow) { sbMd.Append("    - [" + mdText(titleWithYear(dRow)) + "](#" + catalogValue(dRow, "asin").ToLower() + ")\r\n"); }
        }
        sbMd.Append("- [Appendix A: Books by Author](#appendix-a)\r\n- [Appendix B: Books by Origin](#appendix-b)\r\n- [Appendix C: Books by Reading Progress](#appendix-c)\r\n- [Appendix D: Books by Publisher](#appendix-d)\r\n- [Appendix E: About the Authors](#appendix-e)\r\n\r\n");
        sbHtml.Append("<h2 id=\"introduction\">Introduction</h2>\r\n<p>" + htmlText(sIntroText) + "</p>\r\n<p>" + htmlText(sStats) + "</p>\r\n");
        sbMd.Append("## Introduction {#introduction}\r\n\r\n" + sIntroText + "\r\n\r\n" + sStats + "\r\n\r\n");
        sbHtml.Append("<h2 id=\"books\">Books</h2>\r\n");
        sbMd.Append("## Books {#books}\r\n\r\n");
        foreach (Dictionary<string, object> dRow in lOrdered) appendKindleEntry(sbHtml, sbMd, dRow);
        appendAppendix(sbHtml, sbMd, "appendix-a", "Appendix A: Books by Author", dByAuthor, true);
        appendAppendix(sbHtml, sbMd, "appendix-b", "Appendix B: Books by Origin", dByOrigin);
        appendAppendix(sbHtml, sbMd, "appendix-c", "Appendix C: Books by Reading Progress", dByProgress);
        appendAppendix(sbHtml, sbMd, "appendix-d", "Appendix D: Books by Publisher", dByPublisher);
        appendAuthorsAppendix(sbHtml, sbMd, dByAuthor, new Dictionary<string, int>(), "appendix-e", "Appendix E: About the Authors", "Every author in the Kindle library, with the number of books they account for, and a biography when a reliable one was found.");
        sbHtml.Append("</body>\r\n</html>\r\n");
        File.WriteAllText(Path.Combine(sDownloadDir, "Kindle_Library.htm"), sbHtml.ToString(), new UTF8Encoding(true));
        File.WriteAllText(Path.Combine(sDownloadDir, "Kindle_Library.md"), sbMd.ToString(), new UTF8Encoding(true));
        log("The Kindle catalog was saved as Kindle_Library.htm and Kindle_Library.md in " + sDownloadDir);
    }

    // Writes Kindle_Library.xlsx with the same screen-reader conventions as
    // the Audible spreadsheet: one region at A1, bold unique headers, the
    // ColumnTitle01 workbook name, capped widths, a frozen top row, and
    // truly empty cells for unknowns.
    static void fillKindleSheet(ExcelPackage oPackage)
    {
        int iCol, iRow, iWidth;
        int[] aWidths;
        string[] aHeaders;
        List<Dictionary<string, object>> lSorted;
        List<string[]> lPairs;

        aHeaders = new string[] { "Title", "ASIN", "Amazon link", "By", "First published", "Format", "Origin", "Percent read", "Progress", "Publisher", "Read online", "Wikipedia" };
        lSorted = new List<Dictionary<string, object>>(lKindleCatalog);
        lSorted.Sort(compareCatalogRowsByTitle);
        ExcelWorksheet oSheet = oPackage.Workbook.Worksheets.Add("Kindle");
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1) oSheet.Cells[1, iCol].Value = aHeaders[iCol - 1];
        oSheet.Cells[1, 1, 1, aHeaders.Length].Style.Font.Bold = true;
        iRow = 1;
        foreach (Dictionary<string, object> dRow in lSorted)
        {
            iRow = iRow + 1;
            Monitor.Enter(dRow);
            oSheet.Cells[iRow, 1].Value = catalogValue(dRow, "title");
            oSheet.Cells[iRow, 2].Value = catalogValue(dRow, "asin");
            oSheet.Cells[iRow, 3].Value = "https://www.amazon.com/dp/" + catalogValue(dRow, "asin");
            lPairs = catalogLinks(dRow, "authors");
            oSheet.Cells[iRow, 4].Value = joinPairNames(lPairs);
            oSheet.Cells[iRow, 5].Value = catalogValue(dRow, "firstPublished");
            oSheet.Cells[iRow, 6].Value = kindleFormatText(catalogValue(dRow, "resourceType"));
            oSheet.Cells[iRow, 7].Value = kindleOriginText(catalogValue(dRow, "originType"));
            if (dRow.ContainsKey("percentageRead")) oSheet.Cells[iRow, 8].Value = Convert.ToInt32(Convert.ToDouble(dRow["percentageRead"]));
            oSheet.Cells[iRow, 9].Value = kindleProgressText(dRow);
            oSheet.Cells[iRow, 10].Value = catalogValue(dRow, "publisher");
            oSheet.Cells[iRow, 11].Value = catalogValue(dRow, "webReaderUrl");
            oSheet.Cells[iRow, 12].Value = catalogValue(dRow, "wikipediaUrl");
            Monitor.Exit(dRow);
        }
        aWidths = new int[aHeaders.Length];
        for (iCol = 1; iCol <= aHeaders.Length; iCol = iCol + 1)
        {
            aWidths[iCol - 1] = aHeaders[iCol - 1].Length;
            for (iRow = 2; iRow <= lSorted.Count + 1; iRow = iRow + 1)
            {
                iWidth = oSheet.Cells[iRow, iCol].Value == null ? 0 : Convert.ToString(oSheet.Cells[iRow, iCol].Value).Length;
                if (iWidth > aWidths[iCol - 1]) aWidths[iCol - 1] = iWidth;
            }
            if (aWidths[iCol - 1] > 40)
            {
                aWidths[iCol - 1] = 40;
                oSheet.Column(iCol).Style.WrapText = true;
            }
            oSheet.Column(iCol).Width = aWidths[iCol - 1] + 2;
        }
        oSheet.View.FreezePanes(2, 1);
        oSheet.Names.Add("ColumnTitle01", oSheet.Cells[1, 1]);
    }

    // Builds bookFido.htm and bookFido.md: one consolidated catalog of every
    // library together, with an introduction, a table of contents naming
    // each entry's library, and a body where every book's entry lists the
    // fields its own library offers, including the Edition field.  There
    // are no appendixes, by design, because of the catalog's size.
    static string buildConsolidatedFiles(string sDownloadDir)
    {
        string sAnchor, sIntroText, sLibrary, sStats;
        Dictionary<string, object> dInner;
        List<Dictionary<string, object>> lEntries;
        StringBuilder sbHtml, sbMd;

        lEntries = new List<Dictionary<string, object>>();
        foreach (Dictionary<string, object> dRow in lCatalog) { dInner = new Dictionary<string, object>(); dInner["row"] = dRow; dInner["lib"] = "Audible"; lEntries.Add(dInner); }
        foreach (Dictionary<string, object> dRow in lKindleCatalog) { dInner = new Dictionary<string, object>(); dInner["row"] = dRow; dInner["lib"] = "Kindle"; lEntries.Add(dInner); }
        foreach (Dictionary<string, object> dRow in lGoodreadsCatalog) { dInner = new Dictionary<string, object>(); dInner["row"] = dRow; dInner["lib"] = "Goodreads"; lEntries.Add(dInner); }
        foreach (Dictionary<string, object> dRow in lBookshareCatalog) { dInner = new Dictionary<string, object>(); dInner["row"] = dRow; dInner["lib"] = "Bookshare"; lEntries.Add(dInner); }
        if (lEntries.Count == 0) return "";
        lEntries.Sort(delegate(Dictionary<string, object> dA, Dictionary<string, object> dB) { return compareCatalogRowsByTitle((Dictionary<string, object>) dA["row"], (Dictionary<string, object>) dB["row"]); });
        sIntroText = "This is the consolidated bookFido catalog: every book from every library together, in order by title.  Each entry lists the fields its own library offers, and an Edition field tells whether the book is Kindle, Audible, Print, EPUB, or Generic.  A book owned in more than one library appears once for each, so the same title may follow itself in a different edition.  There are no appendixes here; each library's own catalog document carries those.";
        sStats = "The libraries hold " + lEntries.Count + " entries together: Audible " + lCatalog.Count + ", Kindle " + lKindleCatalog.Count + ", Goodreads " + lGoodreadsCatalog.Count + ", and Bookshare " + lBookshareCatalog.Count + ".";
        sbHtml = new StringBuilder();
        sbMd = new StringBuilder();
        sbHtml.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n<meta charset=\"utf-8\">\r\n<title>bookFido Catalog</title>\r\n</head>\r\n<body>\r\n");
        sbHtml.Append("<h1>bookFido Catalog</h1>\r\n");
        sbHtml.Append("<nav aria-label=\"Table of contents\">\r\n<h2 id=\"contents\">Contents</h2>\r\n<ul>\r\n");
        sbHtml.Append("<li><a href=\"#introduction\">Introduction</a></li>\r\n<li><a href=\"#books\">Books</a>\r\n<ul>\r\n");
        foreach (Dictionary<string, object> dEntry in lEntries)
        {
            sLibrary = Convert.ToString(dEntry["lib"]);
            sAnchor = entryAnchor((Dictionary<string, object>) dEntry["row"], sLibrary);
            lock ((Dictionary<string, object>) dEntry["row"]) { sbHtml.Append("<li><a href=\"#" + sAnchor + "\">" + htmlText(titleWithYear((Dictionary<string, object>) dEntry["row"])) + ", " + sLibrary + "</a></li>\r\n"); }
        }
        sbHtml.Append("</ul>\r\n</li>\r\n</ul>\r\n</nav>\r\n");
        sbMd.Append("# bookFido Catalog\r\n\r\n## Contents {#contents}\r\n\r\n");
        sbMd.Append("- [Introduction](#introduction)\r\n- [Books](#books)\r\n");
        foreach (Dictionary<string, object> dEntry in lEntries)
        {
            sLibrary = Convert.ToString(dEntry["lib"]);
            sAnchor = entryAnchor((Dictionary<string, object>) dEntry["row"], sLibrary);
            lock ((Dictionary<string, object>) dEntry["row"]) { sbMd.Append("    - [" + mdText(titleWithYear((Dictionary<string, object>) dEntry["row"])) + ", " + sLibrary + "](#" + sAnchor.ToLower() + ")\r\n"); }
        }
        sbMd.Append("\r\n");
        sbHtml.Append("<h2 id=\"introduction\">Introduction</h2>\r\n<p>" + htmlText(sIntroText) + "</p>\r\n<p>" + htmlText(sStats) + "</p>\r\n");
        sbMd.Append("## Introduction {#introduction}\r\n\r\n" + sIntroText + "\r\n\r\n" + sStats + "\r\n\r\n");
        sbHtml.Append("<h2 id=\"books\">Books</h2>\r\n");
        sbMd.Append("## Books {#books}\r\n\r\n");
        foreach (Dictionary<string, object> dEntry in lEntries)
        {
            sLibrary = Convert.ToString(dEntry["lib"]);
            if (sLibrary == "Audible") appendAudibleEntry(sbHtml, sbMd, (Dictionary<string, object>) dEntry["row"]);
            if (sLibrary == "Kindle") appendKindleEntry(sbHtml, sbMd, (Dictionary<string, object>) dEntry["row"]);
            if (sLibrary == "Goodreads") appendGoodreadsEntry(sbHtml, sbMd, (Dictionary<string, object>) dEntry["row"]);
            if (sLibrary == "Bookshare") appendBookshareEntry(sbHtml, sbMd, (Dictionary<string, object>) dEntry["row"]);
        }
        sbHtml.Append("</body>\r\n</html>\r\n");
        File.WriteAllText(Path.Combine(sDownloadDir, "bookFido.htm"), sbHtml.ToString(), new UTF8Encoding(true));
        File.WriteAllText(Path.Combine(sDownloadDir, "bookFido.md"), sbMd.ToString(), new UTF8Encoding(true));
        log("The consolidated catalog was saved as bookFido.htm and bookFido.md in " + sDownloadDir + ": " + lEntries.Count + " entries");
        return Path.Combine(sDownloadDir, "bookFido.htm");
    }

    // The anchor an entry's own renderer gives it, per library.
    static string entryAnchor(Dictionary<string, object> dRow, string sLibrary)
    {
        if (sLibrary == "Goodreads") return "gr" + Convert.ToString(dRow["grId"]);
        if (sLibrary == "Bookshare") return "bs" + Convert.ToString(dRow["bsId"]);
        return catalogValue(dRow, "asin");
    }

    // The opening dialog, built with the Homer Lbc primitives: the
    // explanation in a read-only memo, one checkbox per library in
    // alphabetical order, all checked to begin with, and OK and Cancel.
    static bool showOpeningDialog(string sIntroText)
    {
        bool bOk;
        CheckBox checkAudible, checkBookshare, checkGoodreads, checkKindle;
        TextBox textIntro;

        using (LbcDialog dialogOpen = new LbcDialog("Welcome to bookFido", null))
        {
            textIntro = dialogOpen.addMemo(sIntroText, "What bookFido does, and what to expect");
            textIntro.ReadOnly = true;
            checkAudible = dialogOpen.addCheckBox("Search &Audible", true, "The Audible library walk, companion PDF files, and catalog");
            checkBookshare = dialogOpen.addCheckBox("Search &Bookshare", true, "The Bookshare My History list");
            checkGoodreads = dialogOpen.addCheckBox("Search &Goodreads", true, "The Goodreads My Books shelves");
            checkKindle = dialogOpen.addCheckBox("Search &Kindle", true, "The Kindle library on the Amazon account");
            bOk = dialogOpen.runOkCancel();
            if (!bOk) return false;
            bSearchAudible = checkAudible.Checked;
            bSearchBookshare = checkBookshare.Checked;
            bSearchGoodreads = checkGoodreads.Checked;
            bSearchKindle = checkKindle.Checked;
        }
        if (!bSearchAudible && !bSearchBookshare && !bSearchGoodreads && !bSearchKindle)
        {
            log("No library was checked, so there is nothing to search");
            MessageBox.Show("No library was checked, so there is nothing to search.", "bookFido", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }
        return true;
    }

    // Gives a plain message box the same focus treatment as the timed
    // announcements: a helper thread polls for the box by its caption and
    // forces it to the foreground, so the screen reader announces it even
    // when the program has been in the background for a long stretch.
    static void focusWhenShown(string sCaption)
    {
        Thread threadFocus;

        threadFocus = new Thread(delegate()
        {
            int iWaited;
            IntPtr hBox;

            hBox = IntPtr.Zero;
            iWaited = 0;
            while (hBox == IntPtr.Zero && iWaited < 5000) { Thread.Sleep(100); iWaited = iWaited + 100; hBox = FindWindowW("#32770", sCaption); }
            if (hBox == IntPtr.Zero) { log("The dialog was not found in time for focus: " + sCaption); return; }
            if (!forceForeground(hBox)) log("The dialog could not take focus: " + sCaption);
        });
        threadFocus.IsBackground = true;
        threadFocus.Start();
    }

    static bool forceForeground(IntPtr hWindow)
    {
        uint iBoxThread, iForeThread, iOurThread;
        IntPtr hForeground;
        MSG oMessage;

        if (SetForegroundWindow(hWindow) && GetForegroundWindow() == hWindow) return true;
        PeekMessageW(out oMessage, IntPtr.Zero, 0, 0, 0);
        hForeground = GetForegroundWindow();
        iOurThread = GetCurrentThreadId();
        iForeThread = hForeground == IntPtr.Zero ? 0 : GetWindowThreadProcessId(hForeground, IntPtr.Zero);
        iBoxThread = GetWindowThreadProcessId(hWindow, IntPtr.Zero);
        if (iForeThread != 0 && iForeThread != iOurThread) AttachThreadInput(iOurThread, iForeThread, true);
        if (iBoxThread != 0 && iBoxThread != iOurThread) AttachThreadInput(iOurThread, iBoxThread, true);
        SetForegroundWindow(hWindow);
        BringWindowToTop(hWindow);
        SetFocus(hWindow);
        if (iForeThread != 0 && iForeThread != iOurThread) AttachThreadInput(iOurThread, iForeThread, false);
        if (iBoxThread != 0 && iBoxThread != iOurThread) AttachThreadInput(iOurThread, iBoxThread, false);
        return GetForegroundWindow() == hWindow;
    }

}
