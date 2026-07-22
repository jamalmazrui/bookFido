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
    const string sApiUserAgent = "bookFido/1.0 (https://github.com/JamalMazrui/bookFido; personal library catalog)", sAudibleApiUrl = "https://api.audible.com/1.0/catalog/products/", sEdgePathPrimary = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe", sEdgePathSecondary = "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe", sLibraryUrl = "https://www.audible.com/library/titles", sOpenLibraryUrl = "https://openlibrary.org/search.json", sOpenLibraryAuthorSearchUrl = "https://openlibrary.org/search/authors.json?q=", sOpenLibraryAuthorUrl = "https://openlibrary.org/authors/", sWikipediaApiUrl = "https://en.wikipedia.org/w/api.php?action=query&list=search&format=json&srlimit=1&srsearch=", sWikipediaPageUrl = "https://en.wikipedia.org/wiki/", sWikipediaSummaryUrl = "https://en.wikipedia.org/api/rest_v1/page/summary/", sVersionText = "34", sStartUrl = "https://www.audible.com/", sUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0";
    const uint iMbOk = 0x00000000, iMbSetForeground = 0x00010000, iMbTopmost = 0x00040000;

    // Static variable definitions
    static bool bSummaryWritten = false;
    static ClientWebSocket wsCdp = null;
    static HashSet<string> setSeenUrls = new HashSet<string>(), setVisitedPages = new HashSet<string>();
    static int iNextId = 1;
    static Dictionary<string, string> dFailedTitles = new Dictionary<string, string>(), dPdfNames = new Dictionary<string, string>();
    static HashSet<string> setCatalogAsins = new HashSet<string>();
    static List<Dictionary<string, object>> lCatalog = new List<Dictionary<string, object>>();
    static Dictionary<string, string> dAuthorOlBio = new Dictionary<string, string>(), dAuthorWikiBio = new Dictionary<string, string>(), dAuthorWikiUrl = new Dictionary<string, string>();
    static HashSet<string> setAuthorsQueued = new HashSet<string>();
    static int iAudibleDone = 0, iAuthorsDone = 0, iHtmCount = 0, iOpenLibraryDone = 0, iWikipediaDone = 0, iWorkTotal = 0;
    static object oAuthorLock = new object();
    [ThreadStatic] static bool bLastFetchRateLimited;
    static bool[] aLaneSkipping = new bool[3];
    static volatile string sProgressText = "";
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
            "with a heading for every title, its details enriched from Audible's catalog service, Open Library, and Wikipedia (this gathering step takes a few minutes for a large library), and appendixes indexed by author, narrator, series, publisher, rating, and more, saved as Audible_Library.htm, Audible_Library.md, and a sortable Audible_Library.xlsx spreadsheet in your Downloads folder.  " +
            "A note on announcements: this program works to keep each announcement window focused so your screen reader speaks it, but Windows can occasionally withhold focus from a background program; the complete play-by-play is always in bookFido.log.  " +
            "It opens Microsoft Edge at audible.com and uses your existing Audible login when possible.  " +
            "Progress is spoken through brief message boxes, and a full record is written to bookFido.log beside the program.  " +
            "When it finishes, it reports totals, opens the catalog in your web browser, and closes the Edge window it opened.  " +
            "Choose OK to begin, or Cancel to exit.";
        dialogResultAnswer = MessageBox.Show(sIntro, "Welcome to bookFido", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
        if (dialogResultAnswer == DialogResult.Cancel) { log("The user chose Cancel at the introduction, so exiting"); return 0; }
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
        bSignedIn = false;
        for (iTry = 1; iTry <= iLoginTryMax; iTry++)
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
        while (iPage <= iPageMax)
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
        bWalkComplete = true;
        sweepCompanions(sCookieHeader, sDownloadDir);
        waitForEnrichment();
        saveState();
        sLibraryHtmPath = buildLibraryFiles(sDownloadDir);
        try { writeLibraryXlsx(sDownloadDir); }
        catch (Exception oException) { log("The spreadsheet catalog could not be written: " + oException.Message); }
        log("HTML versions of PDFs created this run: " + iHtmCount);
        sSummary = "Downloaded " + lDownloaded.Count + ", skipped " + lSkipped.Count + ", failed " + lFailed.Count + (setDeadCompanions.Count > lFailed.Count ? " (" + setDeadCompanions.Count + " companions are known to be permanently unavailable from Audible)" : "") + ", with " + iHtmCount + " HTML versions of PDFs created and " + lCatalog.Count + " titles cataloged.  " +
            (lFailed.Count > 0 ? "The failed companions are remembered as unavailable and will not be retried on future runs.  " : "") +
            (skippedLaneNames() != "" ? "Because of rate limiting, remaining lookups from " + skippedLaneNames() + " were deferred and will be gathered on a future run.  " : "") +
            "The catalog was saved as Audible_Library.htm, Audible_Library.md, and Audible_Library.xlsx in Downloads, and will open in your web browser when you choose OK.  See bookFido.log for details.";
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
                writeLibraryXlsx(sDownloadDirShared);
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
                log("Details gathered so far: Audible " + iAudibleDone + ", Open Library " + iOpenLibraryDone + ", Wikipedia " + iWikipediaDone + " of " + lCatalog.Count + "; author pages checked: " + iAuthorsDone + " of " + setAuthorsQueued.Count + "; overall " + iPercent + " percent" + (bEtaReady && iEtaMinutes > 0 ? ", about " + iEtaMinutes + sMinutesWord + " remaining" : ""));
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
        log("Detail gathering finished: Audible " + iAudibleDone + ", Open Library " + iOpenLibraryDone + ", Wikipedia " + iWikipediaDone + " of " + lCatalog.Count + "; author pages checked: " + iAuthorsDone + " of " + setAuthorsQueued.Count);
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
    static string writeLibraryXlsx(string sDownloadDir)
    {
        int iCol, iRow, iWidth;
        string sPath;
        int[] aWidths;
        string[] aHeaders;
        List<Dictionary<string, object>> lSorted;
        List<string[]> lPairs;

        sPath = Path.Combine(sDownloadDir, "Audible_Library.xlsx");
        aHeaders = new string[] { "Title", "ASIN", "Audible link", "By", "First published", "Genres", "Language", "Length", "Length in minutes", "Narrated by", "Progress", "Publisher", "Rating", "Rating average", "Ratings count", "Release date", "Series", "Summary", "Wikipedia" };
        lSorted = new List<Dictionary<string, object>>(lCatalog);
        lSorted.Sort(compareCatalogRowsByTitle);
        using (ExcelPackage oPackage = new ExcelPackage())
        {
            ExcelWorksheet oSheet = oPackage.Workbook.Worksheets.Add("Audible Library");
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
            oPackage.Workbook.Names.Add("ColumnTitle01", oSheet.Cells[1, 1]);
            if (File.Exists(sPath)) File.Delete(sPath);
            File.WriteAllBytes(sPath, oPackage.GetAsByteArray());
        }
        log("The spreadsheet catalog was saved as " + sPath);
        return sPath;
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
        if (sYear.Length != 4) { sYear = catalogValue(dRow, "releaseDate"); sYear = sYear.Length >= 4 ? sYear.Substring(0, 4) : ""; }
        foreach (char chOne in sYear) { if (!char.IsDigit(chOne)) return sTitle; }
        return sYear.Length == 4 ? sTitle + " (" + sYear + ")" : sTitle;
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
        int iMinutes;
        string sLead;
        List<string> lAuthors;

        sbHtml.Append("<h2 id=\"appendix-h\">Appendix H: About the Authors</h2>\r\n");
        sbMd.Append("## Appendix H: About the Authors {#appendix-h}\r\n\r\n");
        sLead = "Every author in the library, with the number of titles and the listening time they account for, and a biography when a reliable one was found.";
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
        foreach (Dictionary<string, object> dRow in lOrdered)
        {
            Monitor.Enter(dRow);
            sAsin = catalogValue(dRow, "asin");
            sTitle = catalogValue(dRow, "title");
            sUrl = cleanAudibleUrl(catalogValue(dRow, "url"));
            if (sTitle == "") { Monitor.Exit(dRow); continue; }
            sbHtml.Append("<h2 id=\"" + sAsin + "\"><a href=\"" + htmlText(sUrl) + "\">" + htmlText(titleWithYear(dRow)) + "</a></h2>\r\n");
            sbMd.Append("## [" + mdText(titleWithYear(dRow)) + "](" + sUrl + ") {#" + sAsin.ToLower() + "}\r\n\r\n");
            lPairs = catalogLinks(dRow, "authors");
            appendField(sbHtml, sbMd, "By", linksHtml(lPairs), linksMd(lPairs));
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
