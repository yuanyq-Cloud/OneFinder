using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
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
        /// 搜索词命中的片段列表（每个片段包含前后文）
        /// </summary>
        public List<string> Snippets { get; set; } = new();

        /// <summary>
        /// 命中对象的 ID 列表（用于页内导航）
        /// </summary>
        public List<string> HitObjectIds { get; set; } = new();

        public override string ToString() =>
            $"{NotebookName}  ›  {SectionName}  ›  {PageName}";

        /// <summary>
        /// 获取格式化的搜索结果摘要（包含片段预览）
        /// </summary>
        public string GetDisplayText()
        {
            string basePath = ToString();
            if (Snippets.Count == 0) return basePath;

            // 显示第一个片段
            string firstSnippet = Snippets[0];
            if (Snippets.Count > 1)
                return $"{basePath}\n    {firstSnippet} … (+{Snippets.Count - 1} 处匹配)";
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

        /// <summary>
        /// 匹配的文本片段
        /// </summary>
        public string Snippet { get; set; } = string.Empty;

        /// <summary>
        /// 匹配对象的 ID（用于页内导航）
        /// </summary>
        public string? ObjectId { get; set; }

        /// <summary>
        /// 该页面中的匹配序号（1-based）
        /// </summary>
        public int MatchIndex { get; set; }

        /// <summary>
        /// 该页面的总匹配数
        /// </summary>
        public int TotalMatches { get; set; }

        public string GetPagePath() =>
            $"{NotebookName}  ›  {SectionName}  ›  {PageName}";

        public string GetMatchInfo() =>
            TotalMatches > 1 ? $"[{MatchIndex}/{TotalMatches}]" : "";
    }

    /// <summary>
    /// 封装对 OneNote COM API 的访问，通过 XML 全文搜索页面。
    /// </summary>
    public class OneNoteService : IDisposable
    {
        private Application? _app;
        private bool _disposed;

        // OneNote XML 命名空间（版本号为 2013+）
        private static readonly XNamespace NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

        public OneNoteService()
        {
            // 如果 OneNote 已经打开则附加，否则启动一个新实例
            _app = new Application();
        }

        /// <summary>
        /// 获取当前打开的笔记本ID
        /// </summary>
        public string? GetCurrentNotebookId()
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            try
            {
                // 获取当前页面ID
                string currentPageId = _app.Windows.CurrentWindow.CurrentPageId;
                if (string.IsNullOrEmpty(currentPageId)) return null;

                // 获取完整层次结构
                _app.GetHierarchy("", HierarchyScope.hsNotebooks, out string xml);
                var doc = XDocument.Parse(xml);

                // 遍历所有笔记本查找包含当前页面的笔记本
                foreach (var notebook in doc.Descendants(NS + "Notebook"))
                {
                    string notebookId = notebook.Attribute("ID")?.Value ?? string.Empty;
                    if (string.IsNullOrEmpty(notebookId)) continue;

                    // 获取笔记本的所有页面
                    _app.GetHierarchy(notebookId, HierarchyScope.hsPages, out string notebookXml);
                    var notebookDoc = XDocument.Parse(notebookXml);

                    // 检查当前页面是否在这个笔记本中
                    var page = notebookDoc.Descendants(NS + "Page")
                        .FirstOrDefault(p => p.Attribute("ID")?.Value == currentPageId);

                    if (page != null)
                    {
                        return notebookId;
                    }
                }
            }
            catch
            {
                // 如果无法获取当前笔记本，返回null
            }

            return null;
        }

        /// <summary>
        /// 遍历所有笔记本 → 节 → 页，在页面 XML 中做大小写不敏感的全文匹配。
        /// </summary>
        /// <param name="query">搜索关键词</param>
        /// <param name="currentNotebookOnly">是否仅搜索当前笔记本</param>
        /// <param name="progress">进度回调</param>
        public List<PageResult> Search(string query, bool currentNotebookOnly = false,
            Action<string>? progress = null, CancellationToken cancellationToken = default)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));
            if (string.IsNullOrWhiteSpace(query)) return new List<PageResult>();

            var results = new List<PageResult>();
            string normalizedQuery = NormalizeWhitespace(query);
            string queryLower = normalizedQuery.ToLowerInvariant();

            // 如果需要仅搜索当前笔记本，获取当前笔记本ID
            string? currentNotebookId = null;
            if (currentNotebookOnly)
            {
                cancellationToken.ThrowIfCancellationRequested();
                currentNotebookId = GetCurrentNotebookId();
                if (string.IsNullOrEmpty(currentNotebookId))
                {
                    progress?.Invoke("无法获取当前笔记本，将搜索所有笔记本");
                    currentNotebookOnly = false;
                }
            }

            // 获取所有笔记本的层次结构 XML
            _app.GetHierarchy(null, HierarchyScope.hsPages, out string hierarchyXml);
            var hierarchy = XDocument.Parse(hierarchyXml);

            foreach (var notebook in hierarchy.Descendants(NS + "Notebook"))
            {
                cancellationToken.ThrowIfCancellationRequested();

                string nbId = notebook.Attribute("ID")?.Value ?? string.Empty;
                string nbName = notebook.Attribute("name")?.Value ?? "(未命名笔记本)";

                // 如果仅搜索当前笔记本，跳过其他笔记本
                if (currentNotebookOnly && !string.IsNullOrEmpty(currentNotebookId) && nbId != currentNotebookId)
                {
                    continue;
                }

                progress?.Invoke($"扫描笔记本：{nbName}");

                foreach (var section in notebook.Descendants(NS + "Section"))
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    // 跳过受密码保护或已加密的节
                    if (section.Attribute("locked")?.Value == "true") continue;
                    if (section.Attribute("isInRecycleBin")?.Value == "true") continue;

                    string secName = section.Attribute("name")?.Value ?? "(未命名节)";

                    foreach (var page in section.Elements(NS + "Page"))
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        string pageId   = page.Attribute("ID")?.Value ?? string.Empty;
                        string pageName = page.Attribute("name")?.Value ?? "(未命名页面)";

                        if (string.IsNullOrEmpty(pageId)) continue;

                        try
                        {
                            // 获取页面完整 XML（包含所有文本内容）
                            _app.GetPageContent(pageId, out string pageXml,
                                PageInfo.piAll, XMLSchema.xs2013);

                            var pageDoc = XDocument.Parse(pageXml);
                            var snippets = new List<string>();
                            var hitObjectIds = new List<string>();

                            // 提取所有 T 元素（文本节点）
                            ExtractTextMatches(pageDoc, queryLower, normalizedQuery.Length, snippets, hitObjectIds);

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
                        catch
                        {
                            // 单页读取失败时跳过，不影响整体搜索
                        }
                    }
                }
            }

            return results;
        }

        /// <summary>
        /// 从页面 XML 中提取包含匹配文本的片段
        /// </summary>
        private void ExtractTextMatches(XDocument pageDoc, string queryLower, int queryLength,
            List<string> snippets, List<string> hitObjectIds)
        {
            foreach (var textElement in pageDoc.Descendants(NS + "T"))
            {
                string text = BuildSearchableTextMirror(textElement);
                if (string.IsNullOrWhiteSpace(text)) continue;

                string textLower = text.ToLowerInvariant();
                int index = textLower.IndexOf(queryLower, StringComparison.Ordinal);

                if (index >= 0)
                {
                    string snippet = ExtractSnippet(text, index, queryLength);
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

        /// <summary>
        /// 从文本中提取匹配位置的上下文片段
        /// </summary>
        private string ExtractSnippet(string text, int matchIndex, int matchLength, 
            int contextLength = 30)
        {
            int start = Math.Max(0, matchIndex - contextLength);
            int end = Math.Min(text.Length, matchIndex + matchLength + contextLength);

            string prefix = start > 0 ? "…" : "";
            string suffix = end < text.Length ? "…" : "";

            string snippet = text.Substring(start, end - start);

            // 在匹配词前后添加标记（在原始文本上）
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

        /// <summary>
        /// 为单个文本节点建立纯文本镜像：逐个子节点提取、清理 CDATA/HTML，再拼接搜索。
        /// </summary>
        private string BuildSearchableTextMirror(XElement textElement)
        {
            var builder = new StringBuilder();
            foreach (var node in textElement.Nodes())
            {
                AppendNodePlainText(node, builder);
            }

            if (builder.Length == 0)
            {
                AppendPlainTextFragment(textElement.Value, builder);
            }

            return NormalizeWhitespace(builder.ToString());
        }

        private void AppendNodePlainText(XNode node, StringBuilder builder)
        {
            switch (node)
            {
                case XCData cdata:
                    AppendPlainTextFragment(cdata.Value, builder);
                    break;
                case XText text:
                    AppendPlainTextFragment(text.Value, builder);
                    break;
                case XElement element:
                    foreach (var child in element.Nodes())
                    {
                        AppendNodePlainText(child, builder);
                    }
                    break;
            }
        }

        private void AppendPlainTextFragment(string fragment, StringBuilder builder)
        {
            string plainText = CleanFragmentToPlainText(fragment);
            if (string.IsNullOrWhiteSpace(plainText)) return;

            if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1]) && !char.IsWhiteSpace(plainText[0]))
            {
                builder.Append(' ');
            }

            builder.Append(plainText);
        }

        private string CleanFragmentToPlainText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            text = StripLiteralCDataMarkers(text);
            text = HtmlDecodeRepeatedly(text);

            string parsedText = TryConvertMarkupToPlainText(text);
            if (!string.IsNullOrEmpty(parsedText))
            {
                return NormalizeWhitespace(parsedText);
            }

            text = Regex.Replace(text, @"<(?:br|hr)\s*/?>", " ", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"</(?:p|div|li|tr|td|th|h[1-6])\s*>", " ", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<[^>]+>", string.Empty, RegexOptions.Singleline);
            text = HtmlDecodeRepeatedly(text);
            text = StripLiteralCDataMarkers(text);

            return NormalizeWhitespace(text);
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
                        AppendPlainTextFragment(cdataNode.Value, builder);
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
                || Regex.IsMatch(elementName, @"^h[1-6]$", RegexOptions.IgnoreCase);
        }

        private string StripLiteralCDataMarkers(string text)
        {
            string previous;
            do
            {
                previous = text;
                text = Regex.Replace(text, @"<!\[CDATA\[(.*?)\]\]>", "$1", RegexOptions.Singleline);
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
            return Regex.Replace(text, @"\s+", " ").Trim();
        }

        /// <summary>
        /// 在 OneNote 中打开指定页面（将焦点跳转到该页）。
        /// </summary>
        public void NavigateToPage(string pageId, string? objectId = null)
        {
            if (_app == null) throw new ObjectDisposedException(nameof(OneNoteService));

            // 如果提供了对象 ID，尝试导航到页内对象
            if (!string.IsNullOrEmpty(objectId))
            {
                try
                {
                    _app.NavigateTo(pageId, objectId);
                    return;
                }
                catch
                {
                    // 如果页内导航失败，降级为普通页面导航
                }
            }

            // 导航到页面
            _app.NavigateTo(pageId);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                if (_app != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);
                    _app = null;
                }
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}
