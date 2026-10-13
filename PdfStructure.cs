// PdfStructure.cs -- a tagged PDF's own structure, as HTML: its headings, lists, tables and figure descriptions.
//
// WHY (11 October 2026). bookFido made each companion PDF's .htm from the look of the page alone: type sizes for
// headings, block positions for order. A tagged PDF already says what each piece is -- H1 to H6, P, L and LI,
// Table, TR, TH and TD, Figure with its Alt -- and saying so is better than guessing. This reads that structure
// tree and writes it as HTML, the way HomerScribe reads one before making a Word document. Its helpers --
// joining tags to the words they mark, resolving references, mapping custom tags to standard ones -- are
// HomerScribe's, from PdfRead.cs, where they are tested on real files.
//
// USED ONLY WHEN IT IS WHOLE. Many PDFs are tagged badly or in part. The structure is used only when its words
// cover most of the words on the pages; otherwise bodyHtml returns nothing, and bookFido reads the pages by
// layout as before. Any error also returns nothing, never an exception, so this can only add.

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

static class PdfStructure
{
    // At least this share of the pages' letters must be reached through the tags for the tags to be trusted.
    const double c_nCoverageNeeded = 0.6;
    const int c_iDepthLimit = 60;

    class Node
    {
        public string sKind = "";
        public string sAlt = "";
        public StringBuilder sbText = new StringBuilder();
        public List<Node> lChildren = new List<Node>();
    }

    static Dictionary<string, string> dRoleMap = new Dictionary<string, string>();
    static int iPages = 0;
    static int iCurrentPage = 1;

    // The body of the HTML document, from h2 down, or "" when the PDF is not tagged well enough to use. sHow says
    // which, for the log.
    public static string bodyHtml(string sPdfPath, out string sHow)
    {
        sHow = "";
        try
        {
            using (PdfDocument oDoc = PdfDocument.Open(sPdfPath))
            {
                Dictionary<string, string> dText = textByMark(oDoc);
                if (dText.Count == 0) { sHow = "not tagged: no marked content"; return ""; }
                UglyToad.PdfPig.Tokens.DictionaryToken oCatalog = oDoc.Structure.Catalog.CatalogDictionary;
                UglyToad.PdfPig.Tokens.IToken oRootToken;
                if (oCatalog == null || !oCatalog.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("StructTreeRoot"), out oRootToken)) { sHow = "not tagged: no structure tree"; return ""; }
                UglyToad.PdfPig.Tokens.DictionaryToken oRoot = resolved(oDoc, oRootToken) as UglyToad.PdfPig.Tokens.DictionaryToken;
                if (oRoot == null) { sHow = "not tagged: the structure tree is unreadable"; return ""; }
                readRoleMap(oDoc, oRoot);
                iPages = oDoc.NumberOfPages;
                iCurrentPage = 1;
                Node oTop = new Node();
                oTop.sKind = "Document";
                UglyToad.PdfPig.Tokens.IToken oKids;
                if (oRoot.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("K"), out oKids)) walk(oDoc, oKids, oTop, dText, 0);
                int iPageLetters = 0;
                foreach (Page oPage in oDoc.GetPages()) foreach (Letter oLetter in oPage.Letters) if (!string.IsNullOrWhiteSpace(oLetter.Value)) iPageLetters = iPageLetters + 1;
                int iTagLetters = letterCount(oTop);
                if (iPageLetters == 0) { sHow = "no text on the pages"; return ""; }
                double nCoverage = (double)iTagLetters / iPageLetters;
                if (nCoverage < c_nCoverageNeeded) { sHow = string.Format("tagged in part: the tags reach {0:0}% of the text, so the layout is read instead", nCoverage * 100); return ""; }
                StringBuilder sbHtml = new StringBuilder();
                int iLastLevel = 1;
                render(oTop, sbHtml, ref iLastLevel);
                if (sbHtml.Length == 0) { sHow = "tagged, but the tags gave nothing to write"; return ""; }
                sHow = string.Format("tagged structure, reaching {0:0}% of the text", Math.Min(100, nCoverage * 100));
                return sbHtml.ToString();
            }
        }
        catch (Exception oError)
        {
            sHow = "the structure could not be read: " + oError.Message;
            return "";
        }
    }

    // ---- reading the tree --------------------------------------------------------------------------------------

    static void walk(PdfDocument oDoc, UglyToad.PdfPig.Tokens.IToken oToken, Node oParent, Dictionary<string, string> dText, int iDepth)
    {
        if (iDepth > c_iDepthLimit) return;
        oToken = resolved(oDoc, oToken);
        UglyToad.PdfPig.Tokens.ArrayToken oArray = oToken as UglyToad.PdfPig.Tokens.ArrayToken;
        if (oArray != null)
        {
            foreach (UglyToad.PdfPig.Tokens.IToken oOne in oArray.Data) walk(oDoc, oOne, oParent, dText, iDepth + 1);
            return;
        }
        UglyToad.PdfPig.Tokens.NumericToken oNumber = oToken as UglyToad.PdfPig.Tokens.NumericToken;
        if (oNumber != null) { addMarked(oParent, dText, oNumber.Int); return; }
        UglyToad.PdfPig.Tokens.DictionaryToken oNode = oToken as UglyToad.PdfPig.Tokens.DictionaryToken;
        if (oNode == null) return;
        UglyToad.PdfPig.Tokens.IToken oValue;
        // A marked-content reference: { /Type /MCR /MCID n }.
        if (!oNode.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("S"), out oValue))
        {
            if (oNode.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("MCID"), out oValue))
            {
                UglyToad.PdfPig.Tokens.NumericToken oId = resolved(oDoc, oValue) as UglyToad.PdfPig.Tokens.NumericToken;
                if (oId != null) addMarked(oParent, dText, oId.Int);
            }
            return;
        }
        UglyToad.PdfPig.Tokens.NameToken oName = resolved(oDoc, oValue) as UglyToad.PdfPig.Tokens.NameToken;
        Node oChild = new Node();
        oChild.sKind = oName == null ? "" : standardTag(oName.Data);
        oChild.sAlt = stringFrom(oDoc, oNode, "Alt");
        if (oChild.sAlt == "") oChild.sAlt = stringFrom(oDoc, oNode, "ActualText");
        oParent.lChildren.Add(oChild);
        if (oNode.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("K"), out oValue)) walk(oDoc, oValue, oChild, dText, iDepth + 1);
    }

    // The words one marked-content id holds, from the page reached last, then onward, then from the start: a
    // structure tree is mostly in page order, and this is HomerScribe's wordsForMark, which relies on that.
    static void addMarked(Node oNode, Dictionary<string, string> dText, int iId)
    {
        string sSaid;
        for (int iPass = 0; iPass < 2; iPass = iPass + 1)
        {
            int iFrom = iPass == 0 ? iCurrentPage : 1, iTo = iPass == 0 ? iPages : iCurrentPage - 1;
            for (int iAt = iFrom; iAt <= iTo; iAt = iAt + 1)
            {
                if (!dText.TryGetValue(iAt.ToString() + ":" + iId.ToString(), out sSaid)) continue;
                iCurrentPage = iAt;
                Node oText = new Node();
                oText.sKind = "#text";
                oText.sbText.Append(sSaid);
                oNode.lChildren.Add(oText);
                return;
            }
        }
    }

    static Dictionary<string, string> textByMark(PdfDocument oDoc)
    {
        Dictionary<string, string> dText = new Dictionary<string, string>();
        int iPage = 0;
        foreach (Page oPage in oDoc.GetPages())
        {
            iPage = iPage + 1;
            try { foreach (MarkedContentElement oMark in oPage.GetMarkedContents()) gatherMark(oMark, iPage, dText); }
            catch (Exception) { }
        }
        return dText;
    }

    static void gatherMark(MarkedContentElement oMark, int iPage, Dictionary<string, string> dText)
    {
        if (oMark == null) return;
        try
        {
            if (oMark.MarkedContentIdentifier >= 0)
            {
                StringBuilder sbSaid = new StringBuilder();
                foreach (Letter oLetter in oMark.Letters) sbSaid.Append(oLetter.Value);
                string sSaid = sbSaid.ToString().Trim();
                if (sSaid != "")
                {
                    string sKey = iPage.ToString() + ":" + oMark.MarkedContentIdentifier.ToString();
                    dText[sKey] = dText.ContainsKey(sKey) ? dText[sKey] + " " + sSaid : sSaid;
                }
            }
            foreach (MarkedContentElement oChild in oMark.Children) gatherMark(oChild, iPage, dText);
        }
        catch (Exception) { }
    }

    static void readRoleMap(PdfDocument oDoc, UglyToad.PdfPig.Tokens.DictionaryToken oRoot)
    {
        dRoleMap.Clear();
        try
        {
            UglyToad.PdfPig.Tokens.IToken oToken;
            if (!oRoot.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("RoleMap"), out oToken)) return;
            UglyToad.PdfPig.Tokens.DictionaryToken oMap = resolved(oDoc, oToken) as UglyToad.PdfPig.Tokens.DictionaryToken;
            if (oMap == null) return;
            foreach (KeyValuePair<string, UglyToad.PdfPig.Tokens.IToken> oPair in oMap.Data)
            {
                UglyToad.PdfPig.Tokens.NameToken oTo = resolved(oDoc, oPair.Value) as UglyToad.PdfPig.Tokens.NameToken;
                if (oTo != null) dRoleMap[oPair.Key] = oTo.Data;
            }
        }
        catch (Exception) { }
    }

    static string standardTag(string sName)
    {
        string sAt = sName ?? "", sNext;
        for (int iHop = 0; iHop < 8; iHop = iHop + 1)
        {
            if (!dRoleMap.TryGetValue(sAt, out sNext) || sNext == null || sNext == sAt) break;
            sAt = sNext;
        }
        return sAt;
    }

    static UglyToad.PdfPig.Tokens.IToken resolved(PdfDocument oDoc, UglyToad.PdfPig.Tokens.IToken oToken)
    {
        try
        {
            for (int iHop = 0; iHop < 8; iHop = iHop + 1)
            {
                UglyToad.PdfPig.Tokens.IndirectReferenceToken oRef = oToken as UglyToad.PdfPig.Tokens.IndirectReferenceToken;
                if (oRef == null) return oToken;
                // PdfPig 0.1.14's public way to follow a reference (11 October 2026: TokenScanner, which HomerScribe used, is not
                // public in this version, and the first build stopped on it).
                oToken = oDoc.Structure.GetObject(oRef.Data).Data;
            }
        }
        catch (Exception) { }
        return oToken;
    }

    static string stringFrom(PdfDocument oDoc, UglyToad.PdfPig.Tokens.DictionaryToken oNode, string sKey)
    {
        try
        {
            UglyToad.PdfPig.Tokens.IToken oToken;
            if (!oNode.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create(sKey), out oToken)) return "";
            oToken = resolved(oDoc, oToken);
            UglyToad.PdfPig.Tokens.StringToken oText = oToken as UglyToad.PdfPig.Tokens.StringToken;
            if (oText != null) return (oText.Data ?? "").Trim();
            UglyToad.PdfPig.Tokens.HexToken oHex = oToken as UglyToad.PdfPig.Tokens.HexToken;
            if (oHex != null) return (oHex.Data ?? "").Trim();
        }
        catch (Exception) { }
        return "";
    }

    // ---- writing it ------------------------------------------------------------------------------------------

    static int letterCount(Node oNode)
    {
        int iCount = 0;
        foreach (char c in oNode.sbText.ToString()) if (!char.IsWhiteSpace(c)) iCount = iCount + 1;
        foreach (Node oChild in oNode.lChildren) iCount = iCount + letterCount(oChild);
        return iCount;
    }

    // Every word under a node, as one run of text, with a hyphen at a line end rejoined.
    static string textOf(Node oNode)
    {
        StringBuilder sbAll = new StringBuilder();
        gatherText(oNode, sbAll);
        string sText = Regex.Replace(sbAll.ToString(), @"(\w)- (\p{Ll})", "$1$2");
        return Regex.Replace(sText, @"\s+", " ").Trim();
    }

    static void gatherText(Node oNode, StringBuilder sbAll)
    {
        if (oNode.sKind == "Artifact") return;
        if (oNode.sbText.Length > 0) sbAll.Append(oNode.sbText).Append(' ');
        foreach (Node oChild in oNode.lChildren) gatherText(oChild, sbAll);
    }

    static string html(string sText)
    {
        return (sText ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }

    static bool isBlock(string sKind)
    {
        return Regex.IsMatch(sKind, @"^(H[1-6]?|P|L|LI|Table|TR|TH|TD|THead|TBody|TFoot|Figure|BlockQuote|Caption|TOC|TOCI|Note|Sect|Art|Part|Div|Document|Index|Code|Formula)$");
    }

    // Headings keep their own levels, below the page's one h1, and never jump more than one level deeper; lists,
    // tables and figures are written as what they are; anything else is a paragraph or a container of blocks.
    static void render(Node oNode, StringBuilder sbHtml, ref int iLastLevel)
    {
        string sKind = oNode.sKind, sText;
        Match oHead = Regex.Match(sKind, @"^H([1-6])?$");
        if (oHead.Success)
        {
            sText = textOf(oNode);
            if (sText == "") return;
            int iLevel = oHead.Groups[1].Success ? int.Parse(oHead.Groups[1].Value) + 1 : 2;
            if (iLevel > 6) iLevel = 6;
            if (iLevel > iLastLevel + 1) iLevel = iLastLevel + 1;
            iLastLevel = iLevel;
            sbHtml.Append("<h" + iLevel + ">" + html(sText) + "</h" + iLevel + ">\r\n");
            return;
        }
        switch (sKind)
        {
            case "Artifact": return;
            case "Figure":
                sText = oNode.sAlt != "" ? oNode.sAlt : textOf(oNode);
                if (sText != "") sbHtml.Append("<p><strong>Image:</strong> " + html(sText) + "</p>\r\n");
                return;
            case "L":
            case "TOC":
                sbHtml.Append(looksNumbered(oNode) ? "<ol>\r\n" : "<ul>\r\n");
                foreach (Node oChild in oNode.lChildren) renderItem(oChild, sbHtml, ref iLastLevel);
                sbHtml.Append(looksNumbered(oNode) ? "</ol>\r\n" : "</ul>\r\n");
                return;
            case "Table":
                sbHtml.Append("<table>\r\n");
                renderRows(oNode, sbHtml);
                sbHtml.Append("</table>\r\n");
                return;
            case "BlockQuote":
                sText = textOf(oNode);
                if (sText != "") sbHtml.Append("<blockquote><p>" + html(sText) + "</p></blockquote>\r\n");
                return;
        }
        bool bHasBlocks = false;
        foreach (Node oChild in oNode.lChildren) if (isBlock(oChild.sKind)) { bHasBlocks = true; break; }
        if (!bHasBlocks && sKind != "Document" && sKind != "Part" && sKind != "Sect" && sKind != "Art" && sKind != "Div")
        {
            sText = textOf(oNode);
            if (sText != "") sbHtml.Append("<p>" + html(sText) + "</p>\r\n");
            return;
        }
        // A container: its loose words, between its blocks, become paragraphs of their own.
        StringBuilder sbLoose = new StringBuilder();
        foreach (Node oChild in oNode.lChildren)
        {
            if (isBlock(oChild.sKind))
            {
                flushLoose(sbLoose, sbHtml);
                render(oChild, sbHtml, ref iLastLevel);
            }
            else
            {
                gatherText(oChild, sbLoose);
            }
        }
        flushLoose(sbLoose, sbHtml);
    }

    static void flushLoose(StringBuilder sbLoose, StringBuilder sbHtml)
    {
        string sText = Regex.Replace(sbLoose.ToString(), @"\s+", " ").Trim();
        if (sText != "") sbHtml.Append("<p>" + html(sText) + "</p>\r\n");
        sbLoose.Length = 0;
    }

    // One list item: its label dropped, since ol or ul says it, and a list nested in its body kept as a list.
    static void renderItem(Node oItem, StringBuilder sbHtml, ref int iLastLevel)
    {
        if (oItem.sKind == "Artifact") return;
        StringBuilder sbItem = new StringBuilder(), sbNested = new StringBuilder();
        collectItem(oItem, sbItem, sbNested, ref iLastLevel);
        string sText = Regex.Replace(sbItem.ToString(), @"\s+", " ").Trim();
        if (sText == "" && sbNested.Length == 0) return;
        sbHtml.Append("<li>" + html(sText) + (sbNested.Length > 0 ? "\r\n" + sbNested : "") + "</li>\r\n");
    }

    static void collectItem(Node oNode, StringBuilder sbItem, StringBuilder sbNested, ref int iLastLevel)
    {
        foreach (Node oChild in oNode.lChildren)
        {
            if (oChild.sKind == "Lbl" || oChild.sKind == "Artifact") continue;
            if (oChild.sKind == "L") { render(oChild, sbNested, ref iLastLevel); continue; }
            if (oChild.sbText.Length > 0) sbItem.Append(oChild.sbText).Append(' ');
            collectItem(oChild, sbItem, sbNested, ref iLastLevel);
        }
    }

    // A list is numbered when its labels are: 1. 2. a) iv.
    static bool looksNumbered(Node oList)
    {
        int iLabels = 0, iNumbered = 0;
        foreach (Node oItem in oList.lChildren)
            foreach (Node oPart in oItem.lChildren)
                if (oPart.sKind == "Lbl")
                {
                    iLabels = iLabels + 1;
                    if (Regex.IsMatch(textOf(oPart), @"^\(?(\d+|[a-zA-Z]|[ivxlcIVXLC]+)[.)]$")) iNumbered = iNumbered + 1;
                }
        return iLabels > 0 && iNumbered * 2 > iLabels;
    }

    // Rows wherever they sit -- directly, or in THead, TBody and TFoot -- and their cells, a header cell as th.
    static void renderRows(Node oNode, StringBuilder sbHtml)
    {
        foreach (Node oChild in oNode.lChildren)
        {
            if (oChild.sKind == "TR")
            {
                sbHtml.Append("<tr>");
                foreach (Node oCell in oChild.lChildren)
                {
                    if (oCell.sKind != "TH" && oCell.sKind != "TD") continue;
                    string sTag = oCell.sKind == "TH" ? "th" : "td";
                    sbHtml.Append("<" + sTag + ">" + html(textOf(oCell)) + "</" + sTag + ">");
                }
                sbHtml.Append("</tr>\r\n");
            }
            else if (oChild.sKind == "THead" || oChild.sKind == "TBody" || oChild.sKind == "TFoot")
            {
                renderRows(oChild, sbHtml);
            }
        }
    }
}
