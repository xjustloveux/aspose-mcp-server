using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for getting Word document content, statistics, and document info
/// </summary>
[McpServerToolType]
public class WordContentTool
{
    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordContentTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordContentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a Word content operation (get_content, get_content_detailed, get_statistics, get_document_info).
    /// </summary>
    /// <param name="operation">The operation to perform: get_content, get_content_detailed, get_statistics, get_document_info.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="includeHeaders">Include headers in content (for get_content_detailed, default: false).</param>
    /// <param name="includeFooters">Include footers in content (for get_content_detailed, default: false).</param>
    /// <param name="includeFootnotes">Include footnotes in statistics (for get_statistics, default: true).</param>
    /// <param name="includeTabStops">Include tab stops in document info (for get_document_info, default: false).</param>
    /// <param name="maxChars">Maximum characters to return (for get_content/get_content_detailed).</param>
    /// <param name="offset">Character offset to start reading from (for get_content/get_content_detailed, default: 0).</param>
    /// <returns>Document content, detailed content, statistics, or document info as string or JSON.</returns>
    /// <exception cref="ArgumentException">Thrown when the operation is unknown.</exception>
    [McpServerTool(Name = "word_content")]
    [Description(
        @"Get Word document content, statistics, and document information. Supports 4 operations: get_content, get_content_detailed, get_statistics, get_document_info.

Usage examples:
- Get content: word_content(operation='get_content', path='doc.docx')
- Get detailed content: word_content(operation='get_content_detailed', path='doc.docx', includeHeaders=true, includeFooters=true)
- Get statistics: word_content(operation='get_statistics', path='doc.docx', includeFootnotes=true)
- Get document info: word_content(operation='get_document_info', path='doc.docx', includeTabStops=true)")]
    public string Execute(
        [Description("Operation: get_content, get_content_detailed, get_statistics, get_document_info")]
        string operation,
        [Description("Word document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Include headers in content (for get_content_detailed, default: false)")]
        bool includeHeaders = false,
        [Description("Include footers in content (for get_content_detailed, default: false)")]
        bool includeFooters = false,
        [Description("Include footnotes in statistics (for get_statistics, default: true)")]
        bool includeFootnotes = true,
        [Description("Include tab stops in document info (for get_document_info, default: false)")]
        bool includeTabStops = false,
        [Description(
            "Maximum characters to return (for get_content/get_content_detailed). Use for large documents to avoid token overflow.")]
        int? maxChars = null,
        [Description("Character offset to start reading from (for get_content/get_content_detailed, default: 0)")]
        int offset = 0)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "get_content" => GetContent(ctx, maxChars, offset),
            "get_content_detailed" => GetContentDetailed(ctx, includeHeaders, includeFooters),
            "get_statistics" => GetStatistics(ctx, includeFootnotes),
            "get_document_info" => GetDocumentInfo(ctx, includeTabStops),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets document content as plain text with optional pagination.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="maxChars">The maximum number of characters to return.</param>
    /// <param name="offset">The character offset to start reading from.</param>
    /// <returns>A string containing the document content.</returns>
    private static string GetContent(DocumentContext<Document> ctx, int? maxChars, int offset)
    {
        var document = ctx.Document;
        var fullText = CleanText(document.GetText());
        var totalLength = fullText.Length;

        string content;
        var hasMore = false;
        if (offset >= totalLength)
        {
            content = "";
        }
        else if (maxChars.HasValue)
        {
            var endIndex = Math.Min(offset + maxChars.Value, totalLength);
            content = fullText.Substring(offset, endIndex - offset);
            hasMore = endIndex < totalLength;
        }
        else
        {
            content = offset > 0 ? fullText.Substring(offset) : fullText;
        }

        var sb = new StringBuilder();
        sb.AppendLine("=== Document Content ===");
        if (maxChars.HasValue || offset > 0)
        {
            sb.AppendLine($"[Showing chars {offset} to {offset + content.Length} of {totalLength}]");
            if (hasMore)
                sb.AppendLine($"[More content available, use offset={offset + content.Length} to continue]");
        }

        sb.AppendLine(content);
        return sb.ToString();
    }

    /// <summary>
    ///     Gets detailed document content including optional headers and footers.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeHeaders">Whether to include header content.</param>
    /// <param name="includeFooters">Whether to include footer content.</param>
    /// <returns>A string containing the detailed document content.</returns>
    private static string GetContentDetailed(DocumentContext<Document> ctx, bool includeHeaders, bool includeFooters)
    {
        var document = ctx.Document;
        var sb = new StringBuilder();
        sb.AppendLine("=== Detailed Document Content ===");

        if (includeHeaders)
        {
            sb.AppendLine("\n--- Headers ---");
            foreach (var section in document.Sections.Cast<Section>())
            foreach (var header in section.HeadersFooters.Cast<HeaderFooter>())
                if (header.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                    header.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                    header.HeaderFooterType == HeaderFooterType.HeaderEven)
                {
                    var headerText = CleanText(header.GetText());
                    if (!string.IsNullOrWhiteSpace(headerText))
                    {
                        sb.AppendLine($"Section {document.Sections.IndexOf(section)} - {header.HeaderFooterType}:");
                        sb.AppendLine(headerText);
                    }
                }
        }

        sb.AppendLine("\n--- Body Content ---");
        foreach (var section in document.Sections.Cast<Section>())
        {
            var bodyText = CleanText(section.Body.GetText());
            if (!string.IsNullOrWhiteSpace(bodyText))
                sb.AppendLine(bodyText);
        }

        if (includeFooters)
        {
            sb.AppendLine("\n--- Footers ---");
            foreach (var section in document.Sections.Cast<Section>())
            foreach (var footer in section.HeadersFooters.Cast<HeaderFooter>())
                if (footer.HeaderFooterType == HeaderFooterType.FooterPrimary ||
                    footer.HeaderFooterType == HeaderFooterType.FooterFirst ||
                    footer.HeaderFooterType == HeaderFooterType.FooterEven)
                {
                    var footerText = CleanText(footer.GetText());
                    if (!string.IsNullOrWhiteSpace(footerText))
                    {
                        sb.AppendLine($"Section {document.Sections.IndexOf(section)} - {footer.HeaderFooterType}:");
                        sb.AppendLine(footerText);
                    }
                }
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Gets document statistics including word count, page count, and element counts.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeFootnotes">Whether to include footnote count in statistics.</param>
    /// <returns>A JSON string containing the document statistics.</returns>
    private static string GetStatistics(DocumentContext<Document> ctx, bool includeFootnotes)
    {
        var document = ctx.Document;
        document.UpdateWordCount();

        var stats = document.BuiltInDocumentProperties;

        var tables = document.GetChildNodes(NodeType.Table, true);
        var shapes = document.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        var images = shapes.Count(s => s.HasImage);

        var result = new
        {
            pages = stats.Pages,
            words = stats.Words,
            characters = stats.Characters,
            charactersWithSpaces = stats.CharactersWithSpaces,
            paragraphs = stats.Paragraphs,
            lines = stats.Lines,
            footnotes = includeFootnotes ? document.GetChildNodes(NodeType.Footnote, true).Count : (int?)null,
            footnotesIncluded = includeFootnotes,
            tables = tables.Count,
            images,
            shapes = shapes.Count,
            statisticsUpdated = true
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets document metadata and properties as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="includeTabStops">Whether to include tab stop information.</param>
    /// <returns>A JSON string containing the document metadata and properties.</returns>
    private static string GetDocumentInfo(DocumentContext<Document> ctx, bool includeTabStops)
    {
        var document = ctx.Document;
        var props = document.BuiltInDocumentProperties;

        List<object>? tabStopsList = null;
        if (includeTabStops)
        {
            tabStopsList = [];
            var sectionIndex = 0;
            foreach (var section in document.Sections.Cast<Section>())
            {
                var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                for (var paraIndex = 0; paraIndex < paragraphs.Count; paraIndex++)
                {
                    var para = paragraphs[paraIndex];
                    if (para.ParagraphFormat.TabStops.Count > 0)
                    {
                        List<object> stops = [];
                        for (var i = 0; i < para.ParagraphFormat.TabStops.Count; i++)
                        {
                            var tabStop = para.ParagraphFormat.TabStops[i];
                            stops.Add(new
                            {
                                position = tabStop.Position,
                                alignment = tabStop.Alignment.ToString()
                            });
                        }

                        tabStopsList.Add(new
                        {
                            sectionIndex,
                            paragraphIndex = paraIndex,
                            tabStops = stops
                        });
                    }
                }

                sectionIndex++;
            }
        }

        var result = new
        {
            title = props.Title,
            author = props.Author,
            subject = props.Subject,
            created = props.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss"),
            modified = props.LastSavedTime.ToString("yyyy-MM-dd HH:mm:ss"),
            pages = props.Pages,
            sections = document.Sections.Count,
            tabStopsIncluded = includeTabStops,
            tabStops = tabStopsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Cleans text by removing control characters and normalizing whitespace.
    /// </summary>
    /// <param name="text">The text to clean.</param>
    /// <returns>The cleaned text with control characters removed and whitespace normalized.</returns>
    private static string CleanText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text;

        var sb = new StringBuilder();
        var lastWasNewline = false;
        var lastWasSpace = false;

        foreach (var c in text)
        {
            if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t')
                continue;

            if (c == '\r')
                continue;

            if (c == '\n')
            {
                if (!lastWasNewline)
                {
                    sb.Append('\n');
                    lastWasNewline = true;
                }
                else
                {
                    if (sb is [.., '\n'] and not [.., '\n', '\n'])
                        sb.Append('\n');
                }

                lastWasSpace = false;
                continue;
            }

            if (c == ' ' || c == '\t')
            {
                if (!lastWasSpace && !lastWasNewline)
                {
                    sb.Append(' ');
                    lastWasSpace = true;
                }

                continue;
            }

            sb.Append(c);
            lastWasNewline = false;
            lastWasSpace = false;
        }

        return sb.ToString().Trim();
    }
}