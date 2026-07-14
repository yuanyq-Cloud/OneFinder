using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace OneFinder
{
    /// <summary>
    /// 代表一个匹配到的 OneNote 页面
    /// </summary>
    public class PageResult
    {
        public string NotebookName { get; set; } = string.Empty;
        public string SectionName  { get; set; } = string.Empty;
        public string PageName     { get; set; } = string.Empty;
        public string PageId       { get; set; } = string.Empty;

        /// <summary>
        /// 页面最后修改时间（来自 OneNote Hierarchy XML 的 lastModifiedTime 属性）
        /// </summary>
        public DateTime LastModifiedTime { get; set; }

        /// <summary>
        /// 命中片段列表（包含前后文）
        /// </summary>
        public List<string> Snippets { get; set; } = new();

        /// <summary>
        /// 命中对象的 ID 列表（用于页内导航）
        /// </summary>
        public List<string> HitObjectIds { get; set; } = new();

        public override string ToString() =>
            $"{NotebookName}  ▸ {SectionName}  ▸ {PageName}";

        public string GetDisplayText()
        {
            string basePath = ToString();
            if (Snippets.Count == 0) return basePath;

            string firstSnippet = Snippets[0];
            if (Snippets.Count > 1)
                return $"{basePath}\n    {firstSnippet} … ({Loc.Fmt("MatchCountSuffix", Snippets.Count - 1)})";
            else
                return $"{basePath}\n    {firstSnippet}";
        }
    }

    /// <summary>
    /// 代表单个匹配项（用于列表显示）
    /// </summary>
    public class MatchResult
    {
        public string NotebookName { get; set; } = string.Empty;
        public string SectionName  { get; set; } = string.Empty;
        public string PageName     { get; set; } = string.Empty;
        public string PageId       { get; set; } = string.Empty;

        public string Snippet { get; set; } = string.Empty;
        public string? ObjectId { get; set; }
        public int MatchIndex { get; set; }
        public int TotalMatches { get; set; }

        /// <summary>
        /// 页面最后修改时间（用于"最近修改"列表显示；搜索匹配为 DateTime.MinValue）
        /// </summary>
        public DateTime LastModifiedTime { get; set; }

        public string GetPagePath() =>
            $"{NotebookName}  ▸ {SectionName}  ▸ {PageName}";

        public string GetMatchInfo() =>
            TotalMatches > 1 ? $"[{MatchIndex}/{TotalMatches}]" : "";

        /// <summary>
        /// 获取用于列表显示的辅助信息：搜索结果显示匹配序号，最近修改列表显示相对时间
        /// </summary>
        public string GetSecondaryInfo()
        {
            if (LastModifiedTime != DateTime.MinValue)
                return FormatRelativeTime(LastModifiedTime);
            return GetMatchInfo();
        }

        public static string FormatRelativeTime(DateTime time)
        {
            DateTime localTime = time.Kind == DateTimeKind.Utc ? time.ToLocalTime() : time;
            TimeSpan diff = DateTime.Now - localTime;

            if (diff.TotalSeconds < 0)
                return Loc.Get("TimeJustNow");
            if (diff.TotalSeconds < 60)
                return Loc.Get("TimeJustNow");
            if (diff.TotalMinutes < 60)
                return Loc.Fmt("TimeMinutesAgo", (int)diff.TotalMinutes);
            if (diff.TotalHours < 24)
                return Loc.Fmt("TimeHoursAgo", (int)diff.TotalHours);
            if (diff.TotalDays < 2 && localTime.Date == DateTime.Now.Date.AddDays(-1))
                return Loc.Fmt("TimeYesterday", localTime.ToString("HH:mm"));
            if (diff.TotalDays < 7)
                return Loc.Fmt("TimeDaysAgo", (int)diff.TotalDays);
            if (localTime.Year == DateTime.Now.Year)
                return localTime.ToString("MM-dd HH:mm");
            return localTime.ToString("yyyy-MM-dd");
        }
    }

    /// <summary>
    /// 封装对 OneNote COM API 的访问，通过 XML 实现全文搜索
    /// </summary>
    public class OneNoteService : IDisposable
    {
        private Application? _app;
        private bool _disposed;

        private static readonly XNamespace NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        private static readonly Regex BlockBreakTagRegex = new(@"<(?:br|hr)\s*/?>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex BlockClosingTagRegex = new(@"</(?:p|div|li|tr|td|th|h[1-6])\s*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex HtmlTagRegex = new(@"<[^>]+>", RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex HeadingElementNameRegex = new(@"^h[1-6]$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex LiteralCDataRegex = new(@"<!\[CDATA\[(.*?)\]\]>", RegexOptions.Singleline | RegexOptions.Compiled);
        private const StringComparison SearchComparison = StringComparison.OrdinalIgnoreCase;

        internal static void Log(string msg)
        {
            try
            {
                var path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "OneFinder.log");
                System.IO.File.AppendAllText(path,
                    $"[{DateTime.Now:HH:mm:ss.fff}][tid={System.Threading.Thread.CurrentThread.ManagedThreadId} apt={System.Threading.Thread.CurrentThread.GetApartmentState()}] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        /// <summary>
        /// 解码 OneNote Hierarchy XML 中 name 属性的值。
        /// OneNote 内部会对文件名非法字符做 ^X 转义（如 + → ^M），
        /// 同时 name 中可能含有 HTML 实体或控制字符，需一并处理。
        /// </summary>
        internal static string DecodeOneNoteName(string name)
        {
            if (string.IsNullOrEmpty(name)) return name;

            // 处理 ^X 转义序列（OneNote 文件系统层编码）
            name = name.Replace("^M", "+")       // CR → +
                      .Replace("^/", "/")
                      .Replace("^\\", "\\")
                      .Replace("^:", ":")
                      .Replace("^*", "*")
                      .Replace("^?", "?")
                      .Replace("^\"", "\"")
                      .Replace("^<", "<")
                      .Replace("^>", ">")
                      .Replace("^|", "|");

            // HTML-decode 处理可能的实体编码（&amp;、&#43; 等）
            name = WebUtility.HtmlDecode(name);

            // 替换残留的控制字符为空格
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                sb.Append(char.IsControl(c) ? ' ' : c);
            }

            return sb.ToString().Trim();
        }

        public OneNoteService()
        {
            Log("OneNoteService.ctor: creating COM Application...");
            _app = new Application();
            Log("OneNoteService.ctor: COM Application created OK");
        }

        /// <summary>
        /// 获取当前打开的笔记本 ID
        /// </summary>
        public string? GetCurrentNotebookId()
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            try
            {
                string? currentPageId = GetCurrentPageId();
                if (string.IsNullOrEmpty(currentPageId)) return null;

                _app.GetHierarchy(null, HierarchyScope.hsPages, out string hierarchyXml);
                var hierarchy = XDocument.Parse(hierarchyXml);
                return FindNotebookIdForPage(hierarchy, currentPageId);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取最近修改的页面列表（仅解析 Hierarchy XML，不逐页获取内容，性能极快）
        /// </summary>
        public List<PageResult> GetRecentPages(int maxCount = 100,
            bool currentNotebookOnly = false)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            _app.GetHierarchy(null, HierarchyScope.hsPages, out string hierarchyXml);
            var hierarchy = XDocument.Parse(hierarchyXml);

            string? currentNotebookId = null;
            if (currentNotebookOnly)
            {
                string? currentPageId = GetCurrentPageId();
                currentNotebookId = string.IsNullOrEmpty(currentPageId)
                    ? null
                    : FindNotebookIdForPage(hierarchy, currentPageId);
            }

            var pages = new List<(PageResult Result, DateTime LastModified)>();

            foreach (var pageEl in hierarchy.Descendants(NS + "Page"))
            {
                // 跳过回收站中的页面
                if (pageEl.Attribute("isInRecycleBin")?.Value == "true") continue;

                var sectionEl = pageEl.Parent;
                if (sectionEl == null) continue;

                // 跳过锁定的分区
                if (sectionEl.Attribute("locked")?.Value == "true") continue;
                if (sectionEl.Attribute("isInRecycleBin")?.Value == "true") continue;

                var notebookEl = sectionEl.Parent;
                if (notebookEl == null) continue;

                string nbId = notebookEl.Attribute("ID")?.Value ?? string.Empty;
                string nbName = DecodeOneNoteName(notebookEl.Attribute("name")?.Value ?? Loc.Get("UnnamedNotebook"));

                // 如果限定当前笔记本，则过滤
                if (currentNotebookOnly && !string.IsNullOrEmpty(currentNotebookId)
                    && nbId != currentNotebookId)
                {
                    continue;
                }

                string pageId = pageEl.Attribute("ID")?.Value ?? string.Empty;
                if (string.IsNullOrEmpty(pageId)) continue;

                string pageName = DecodeOneNoteName(pageEl.Attribute("name")?.Value ?? Loc.Get("UnnamedPage"));
                string secName  = DecodeOneNoteName(sectionEl.Attribute("name")?.Value ?? Loc.Get("UnnamedSection"));

                string lastModifiedStr = pageEl.Attribute("lastModifiedTime")?.Value ?? string.Empty;
                DateTime lastModified = DateTime.MinValue;
                if (!string.IsNullOrEmpty(lastModifiedStr))
                {
                    DateTime.TryParse(lastModifiedStr, null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out lastModified);
                }

                pages.Add((new PageResult
                {
                    NotebookName     = nbName,
                    SectionName      = secName,
                    PageName         = pageName,
                    PageId           = pageId,
                    LastModifiedTime = lastModified,
                    Snippets         = new List<string>(),
                    HitObjectIds     = new List<string>(),
                }, lastModified));
            }

            return pages
                .OrderByDescending(p => p.LastModified)
                .Take(maxCount)
                .Select(p => p.Result)
                .ToList();
        }

        /// <summary>
        /// 提取页面开头文本作为预览（跳过 Title，取正文第一个文本段落）
        /// </summary>
        public string? ExtractPagePreview(string pageId, int maxLength = 120)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            try
            {
                _app.GetPageContent(pageId, out string pageXml,
                    PageInfo.piBasic, XMLSchema.xs2013);

                using var stringReader = new StringReader(pageXml);
                using var xmlReader = XmlReader.Create(stringReader, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                });

                int titleDepth = -1;

                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Element)
                    {
                        if (xmlReader.LocalName == "Title")
                        {
                            titleDepth = xmlReader.Depth;
                            continue;
                        }

                        if (xmlReader.LocalName != "T")
                            continue;

                        // 跳过 Title 内的文本元素
                        if (titleDepth >= 0 && xmlReader.Depth > titleDepth)
                        {
                            xmlReader.Skip();
                            continue;
                        }
                    }
                    else if (xmlReader.NodeType == XmlNodeType.EndElement)
                    {
                        if (xmlReader.LocalName == "Title" && titleDepth >= 0)
                        {
                            titleDepth = -1;
                        }
                        continue;
                    }
                    else
                    {
                        continue;
                    }

                    // 到达此处的是非 Title 内的 <T> 元素
                    string rawText = xmlReader.ReadInnerXml();
                    if (string.IsNullOrWhiteSpace(rawText))
                        continue;

                    string text = BuildSearchableTextMirror(rawText, fastSearch: true);
                    text = NormalizeWhitespace(text);

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    // 跳过无意义的占位段落（如仅含 "Def." 的缩写标记）
                    if (text.Equals("Def.", StringComparison.OrdinalIgnoreCase)
                        || text.Equals("Def", StringComparison.OrdinalIgnoreCase)
                        || text.Equals("Content", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (text.Length > maxLength)
                        text = text.Substring(0, maxLength) + "…";

                    return text;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 执行页面搜索（XML 级文本匹配）
        /// </summary>
        public List<PageResult> Search(string query, bool currentNotebookOnly = false,
            bool fastSearch = false, Action<string>? progress = null,
            CancellationToken cancellationToken = default)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));
            if (string.IsNullOrWhiteSpace(query)) return new List<PageResult>();

            var results = new List<PageResult>();
            string normalizedQuery = NormalizeWhitespace(query);

            _app.GetHierarchy(null, HierarchyScope.hsPages, out string hierarchyXml);
            var hierarchy = XDocument.Parse(hierarchyXml);

            string? currentNotebookId = null;
            if (currentNotebookOnly)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string? currentPageId = GetCurrentPageId();
                currentNotebookId = string.IsNullOrEmpty(currentPageId)
                    ? null
                    : FindNotebookIdForPage(hierarchy, currentPageId);

                if (string.IsNullOrEmpty(currentNotebookId))
                {
                    progress?.Invoke(Loc.Get("CannotGetCurrentNotebook"));
                    currentNotebookOnly = false;
                }
            }

            foreach (var notebook in hierarchy.Descendants(NS + "Notebook"))
            {
                cancellationToken.ThrowIfCancellationRequested();

                string nbId = notebook.Attribute("ID")?.Value ?? string.Empty;
                string nbName = DecodeOneNoteName(notebook.Attribute("name")?.Value ?? Loc.Get("UnnamedNotebook"));

                if (currentNotebookOnly && !string.IsNullOrEmpty(currentNotebookId) && nbId != currentNotebookId)
                {
                    continue;
                }

                progress?.Invoke(Loc.Fmt("ScanningNotebook", nbName));

                foreach (var section in notebook.Descendants(NS + "Section"))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (section.Attribute("locked")?.Value == "true") continue;
                    if (section.Attribute("isInRecycleBin")?.Value == "true") continue;

                    string secName = DecodeOneNoteName(section.Attribute("name")?.Value ?? Loc.Get("UnnamedSection"));

                    foreach (var page in section.Elements(NS + "Page"))
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string pageId   = page.Attribute("ID")?.Value ?? string.Empty;
                        string pageName = DecodeOneNoteName(page.Attribute("name")?.Value ?? Loc.Get("UnnamedPage"));

                        if (string.IsNullOrEmpty(pageId)) continue;

                        try
                        {
                            // 获取页面完整 XML（包含所有文本内容）
                            _app.GetPageContent(pageId, out string pageXml,
                                PageInfo.piAll, XMLSchema.xs2013);
                            var snippets = new List<string>();
                            var hitObjectIds = new List<string>();

                            if (fastSearch)
                            {
                                ExtractTextMatchesFast(pageXml, normalizedQuery, snippets, hitObjectIds);
                            }
                            else
                            {
                                var pageDoc = XDocument.Parse(pageXml);
                                ExtractTextMatches(pageDoc, normalizedQuery, snippets, hitObjectIds, fastSearch: false);
                            }

                            if (snippets.Count > 0)
                            {
                                results.Add(new PageResult
                                {
                                    NotebookName = nbName,
                                    SectionName  = secName,
                                    PageName     = pageName,
                                    PageId       = pageId,
                                    Snippets     = snippets,
                                    HitObjectIds = hitObjectIds,
                                });
                            }
                        }
                        catch (Exception)
                        {
                            // Skip this page on error (locked, corrupted, or insufficient permissions)
                        }
                    }
                }

            }

            return results;
        }

        private void ExtractTextMatchesFast(string pageXml, string query,
            List<string> snippets, List<string> hitObjectIds)
        {
            using var stringReader = new StringReader(pageXml);
            using var xmlReader = XmlReader.Create(stringReader, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = false,
            });

            var oeStack = new Stack<(int Depth, string? ObjectId)>();

            while (xmlReader.Read())
            {
                switch (xmlReader.NodeType)
                {
                    case XmlNodeType.Element:
                        if (xmlReader.LocalName == "OE")
                        {
                            oeStack.Push((xmlReader.Depth, xmlReader.GetAttribute("objectID")));

                            if (xmlReader.IsEmptyElement)
                            {
                                oeStack.Pop();
                            }

                            continue;
                        }

                        if (xmlReader.LocalName != "T")
                        {
                            continue;
                        }

                        string rawText = xmlReader.ReadInnerXml();
                        if (!MightContainQueryFast(rawText, query))
                        {
                            continue;
                        }

                        string text = BuildSearchableTextMirror(rawText, fastSearch: true);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        int index = text.IndexOf(query, SearchComparison);
                        if (index < 0)
                        {
                            continue;
                        }

                        snippets.Add(ExtractSnippet(text, index, query.Length));

                        string? objectId = GetCurrentObjectId(oeStack);
                        if (!string.IsNullOrEmpty(objectId))
                        {
                            hitObjectIds.Add(objectId);
                        }

                        if (snippets.Count >= 5)
                        {
                            return;
                        }

                        continue;

                    case XmlNodeType.EndElement:
                        if (xmlReader.LocalName == "OE")
                        {
                            while (oeStack.Count > 0 && oeStack.Peek().Depth >= xmlReader.Depth)
                            {
                                oeStack.Pop();
                            }
                        }

                        break;
                }
            }
        }

        private void ExtractTextMatches(XDocument pageDoc, string query,
            List<string> snippets, List<string> hitObjectIds, bool fastSearch)
        {
            foreach (var textElement in pageDoc.Descendants(NS + "T"))
            {
                if (fastSearch && !MightContainQueryFast(textElement, query))
                {
                    continue;
                }

                string text = BuildSearchableTextMirror(textElement, fastSearch);
                if (string.IsNullOrWhiteSpace(text)) continue;

                int index = text.IndexOf(query, SearchComparison);

                if (index >= 0)
                {
                    string snippet = ExtractSnippet(text, index, query.Length);
                    snippets.Add(snippet);

                    var oeElement = textElement.Ancestors(NS + "OE").FirstOrDefault();
                    if (oeElement != null)
                    {
                        string? objectId = oeElement.Attribute("objectID")?.Value;
                        if (!string.IsNullOrEmpty(objectId))
                            hitObjectIds.Add(objectId);
                    }

                    if (snippets.Count >= 5) break;
                }
            }
        }

        private string ExtractSnippet(string text, int matchIndex, int matchLength, int contextLength = 30)
        {
            int start = Math.Max(0, matchIndex - contextLength);
            int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

            string prefix = start > 0 ? "…" : "";
            string suffix = end < text.Length ? "…" : "";

            string snippet = text.Substring(start, end - start);

            int highlightStart = matchIndex - start;
            int highlightEnd = highlightStart + matchLength;

            if (highlightStart >= 0 && highlightEnd <= snippet.Length)
            {
                snippet = snippet.Substring(0, highlightStart) +
                         "[" + snippet.Substring(highlightStart, matchLength) + "]" +
                         snippet.Substring(highlightEnd);
            }

            return NormalizeWhitespace(prefix + snippet + suffix);
        }

        private string BuildSearchableTextMirror(XElement textElement, bool fastSearch)
        {
            if (fastSearch && !textElement.HasElements)
            {
                return CleanFragmentToPlainText(textElement.Value, fastSearch: true);
            }

            var builder = new StringBuilder();
            foreach (var node in textElement.Nodes())
            {
                AppendNodePlainText(node, builder, fastSearch);
            }

            if (builder.Length == 0)
            {
                AppendPlainTextFragment(textElement.Value, builder, fastSearch);
            }

            return NormalizeWhitespace(builder.ToString());
        }

        private string BuildSearchableTextMirror(string rawText, bool fastSearch)
        {
            return CleanFragmentToPlainText(rawText, fastSearch);
        }

        private void AppendNodePlainText(XNode node, StringBuilder builder, bool fastSearch)
        {
            switch (node)
            {
                case XCData cdata:
                    AppendPlainTextFragment(cdata.Value, builder, fastSearch);
                    break;
                case XText text:
                    AppendPlainTextFragment(text.Value, builder, fastSearch);
                    break;
                case XElement element:
                    foreach (var child in element.Nodes())
                    {
                        AppendNodePlainText(child, builder, fastSearch);
                    }
                    break;
            }
        }

        private void AppendPlainTextFragment(string fragment, StringBuilder builder, bool fastSearch)
        {
            string plainText = CleanFragmentToPlainText(fragment, fastSearch);
            if (string.IsNullOrWhiteSpace(plainText)) return;

            if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1]) && !char.IsWhiteSpace(plainText[0]))
            {
                builder.Append(' ');
            }

            builder.Append(plainText);
        }

        private string CleanFragmentToPlainText(string text, bool fastSearch)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            if (fastSearch)
            {
                return CleanFragmentToPlainTextFast(text);
            }

            text = StripLiteralCDataMarkers(text);
            text = HtmlDecodeRepeatedly(text);

            string parsedText = TryConvertMarkupToPlainText(text);
            if (!string.IsNullOrEmpty(parsedText))
            {
                return NormalizeWhitespace(parsedText);
            }

            text = BlockBreakTagRegex.Replace(text, " ");
            text = BlockClosingTagRegex.Replace(text, " ");
            text = HtmlTagRegex.Replace(text, string.Empty);
            text = HtmlDecodeRepeatedly(text);
            text = StripLiteralCDataMarkers(text);

            return NormalizeWhitespace(text);
        }

        private string CleanFragmentToPlainTextFast(string text)
        {
            if (CanUsePlainTextFastPath(text))
            {
                return NormalizeWhitespace(text);
            }

            if (ContainsLiteralCData(text))
            {
                text = StripLiteralCDataMarkers(text);
            }

            if (ContainsHtmlEntity(text))
            {
                text = HtmlDecodeRepeatedly(text);
            }

            if (LooksLikeMarkup(text))
            {
                text = BlockBreakTagRegex.Replace(text, " ");
                text = BlockClosingTagRegex.Replace(text, " ");
                text = HtmlTagRegex.Replace(text, string.Empty);
            }

            if (ContainsHtmlEntity(text))
            {
                text = HtmlDecodeRepeatedly(text);
            }

            if (ContainsLiteralCData(text))
            {
                text = StripLiteralCDataMarkers(text);
            }

            return NormalizeWhitespace(text);
        }

        private bool MightContainQueryFast(XElement textElement, string query)
        {
            return MightContainQueryFast(textElement.Value, query);
        }

        private bool MightContainQueryFast(string rawText, string query)
        {
            if (string.IsNullOrWhiteSpace(rawText)) return false;

            if (rawText.IndexOf(query, SearchComparison) >= 0)
            {
                return true;
            }

            if (ContainsCleanupSensitiveSyntax(rawText))
            {
                return true;
            }

            if (query.IndexOf(' ') >= 0 || ContainsNonSpaceWhitespace(rawText))
            {
                return NormalizeWhitespace(rawText).IndexOf(query, SearchComparison) >= 0;
            }

            return false;
        }

        private string? GetCurrentObjectId(Stack<(int Depth, string? ObjectId)> oeStack)
        {
            foreach (var (_, objectId) in oeStack)
            {
                if (!string.IsNullOrEmpty(objectId))
                {
                    return objectId;
                }
            }

            return null;
        }

        private bool CanUsePlainTextFastPath(string text)
        {
            return !ContainsCleanupSensitiveSyntax(text);
        }

        private bool ContainsCleanupSensitiveSyntax(string text)
        {
            return LooksLikeMarkup(text)
                || ContainsHtmlEntity(text)
                || ContainsLiteralCData(text);
        }

        private bool LooksLikeMarkup(string text)
        {
            return text.IndexOf('<') >= 0 && text.IndexOf('>') >= 0;
        }

        private bool ContainsHtmlEntity(string text)
        {
            return text.IndexOf('&') >= 0;
        }

        private bool ContainsLiteralCData(string text)
        {
            return text.IndexOf("CDATA", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool ContainsNonSpaceWhitespace(string text)
        {
            foreach (char ch in text)
            {
                if (char.IsWhiteSpace(ch) && ch != ' ')
                {
                    return true;
                }
            }

            return false;
        }

        private string TryConvertMarkupToPlainText(string text)
        {
            if (text.IndexOf('<') < 0 || text.IndexOf('>') < 0) return string.Empty;

            try
            {
                var root = XElement.Parse($"<root>{text}</root>", LoadOptions.PreserveWhitespace);
                var builder = new StringBuilder();
                AppendElementPlainText(root, builder);
                return builder.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        private void AppendElementPlainText(XElement element, StringBuilder builder)
        {
            bool isBlockElement = IsBlockLikeElement(element.Name.LocalName);
            if (isBlockElement && builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1]))
            {
                builder.Append(' ');
            }

            foreach (var node in element.Nodes())
            {
                switch (node)
                {
                    case XCData cdataNode:
                        AppendPlainTextFragment(cdataNode.Value, builder, false);
                        break;
                    case XText textNode:
                        builder.Append(textNode.Value);
                        break;
                    case XElement childElement:
                        AppendElementPlainText(childElement, builder);
                        break;
                }
            }

            if ((isBlockElement || string.Equals(element.Name.LocalName, "br", StringComparison.OrdinalIgnoreCase))
                && builder.Length > 0
                && !char.IsWhiteSpace(builder[builder.Length - 1]))
            {
                builder.Append(' ');
            }
        }

        private bool IsBlockLikeElement(string elementName)
        {
            return elementName.Equals("p", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("div", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("li", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("tr", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("td", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("th", StringComparison.OrdinalIgnoreCase)
                || elementName.Equals("br", StringComparison.OrdinalIgnoreCase)
                || HeadingElementNameRegex.IsMatch(elementName);
        }

        private string StripLiteralCDataMarkers(string text)
        {
            string previous;
            do
            {
                previous = text;
                text = LiteralCDataRegex.Replace(text, "$1");
            }
            while (!string.Equals(previous, text, StringComparison.Ordinal));

            return text.Replace("<![CDATA[", string.Empty)
                       .Replace("]]>", string.Empty);
        }

        private string HtmlDecodeRepeatedly(string text)
        {
            for (int i = 0; i < 3; i++)
            {
                string decoded = WebUtility.HtmlDecode(text);
                if (string.Equals(decoded, text, StringComparison.Ordinal))
                {
                    break;
                }

                text = decoded;
            }

            return text;
        }

        private string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            var builder = new StringBuilder(text.Length);
            bool seenNonWhitespace = false;
            bool pendingSpace = false;

            foreach (char ch in text)
            {
                if (char.IsWhiteSpace(ch))
                {
                    pendingSpace = seenNonWhitespace;
                    continue;
                }

                if (pendingSpace)
                {
                    builder.Append(' ');
                    pendingSpace = false;
                }

                builder.Append(ch);
                seenNonWhitespace = true;
            }

            return builder.ToString();
        }

        private string? GetCurrentPageId()
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            string currentPageId = _app.Windows.CurrentWindow.CurrentPageId;
            return string.IsNullOrEmpty(currentPageId) ? null : currentPageId;
        }

        /// <summary>
        /// 获取当前页面的笔记本 ID
        /// </summary>
        private string? FindNotebookIdForPage(XDocument hierarchy, string pageId)
        {
            foreach (var notebook in hierarchy.Descendants(NS + "Notebook"))
            {
                string notebookId = notebook.Attribute("ID")?.Value ?? string.Empty;
                if (string.IsNullOrEmpty(notebookId)) continue;

                bool containsPage = notebook.Descendants(NS + "Page")
                    .Any(page => string.Equals(page.Attribute("ID")?.Value, pageId, StringComparison.Ordinal));

                if (containsPage)
                {
                    return notebookId;
                }
            }

            return null;
        }

        /// <summary>
        /// 导航到指定的页面及对应的位置
        /// </summary>
        public void NavigateToPage(string pageId, string? objectId = null)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            if (!string.IsNullOrEmpty(objectId))
            {
                try
                {
                    _app.NavigateTo(pageId, objectId);
                    return;
                }
                catch
                {
                }
            }

            _app.NavigateTo(pageId);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_app != null)
                {
                    Log($"OneNoteService.Dispose: calling FinalReleaseComObject on tid={System.Threading.Thread.CurrentThread.ManagedThreadId} apt={System.Threading.Thread.CurrentThread.GetApartmentState()}");
                    int remaining = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_app);
                    Log($"OneNoteService.Dispose: FinalReleaseComObject done, remaining={remaining}");
                    _app = null;
                }
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}
