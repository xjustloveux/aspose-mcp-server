using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for getting content and statistics from PDF documents
/// </summary>
[McpServerToolType]
public class PdfInfoTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfInfoTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfInfoTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_info")]
    [Description(@"Get content and statistics from PDF documents. Supports 2 operations: get_content, get_statistics.

Usage examples:
- Get content from page: pdf_info(operation='get_content', path='doc.pdf', pageIndex=1)
- Get content from all pages: pdf_info(operation='get_content', path='doc.pdf')
- Get content with limit: pdf_info(operation='get_content', path='doc.pdf', maxPages=50)
- Get statistics: pdf_info(operation='get_statistics', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_content': Get text content from page(s) (required params: path)
- 'get_statistics': Get document statistics (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Page index (1-based, optional for get_content, extracts all if not specified)")]
        int? pageIndex = null,
        [Description("Maximum pages to extract (for get_content without pageIndex, default: 100)")]
        int maxPages = 100)
    {
        return operation.ToLower() switch
        {
            "get_content" => GetContent(sessionId, path, pageIndex, maxPages),
            "get_statistics" => GetStatistics(sessionId, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Extracts text content from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="pageIndex">Optional 1-based page index to extract content from a specific page.</param>
    /// <param name="maxPages">Maximum number of pages to extract when extracting all pages.</param>
    /// <returns>A JSON string containing the extracted text content.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    private string GetContent(string? sessionId, string? path, int? pageIndex, int maxPages)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var textAbsorber = new TextAbsorber();
            document.Pages[pageIndex.Value].Accept(textAbsorber);
            var result = new
            {
                pageIndex = pageIndex.Value,
                totalPages = document.Pages.Count,
                content = textAbsorber.Text
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            var pagesToExtract = Math.Min(maxPages, document.Pages.Count);
            var truncated = document.Pages.Count > maxPages;
            var contentBuilder = new StringBuilder();

            for (var i = 1; i <= pagesToExtract; i++)
            {
                var textAbsorber = new TextAbsorber();
                document.Pages[i].Accept(textAbsorber);
                contentBuilder.AppendLine(textAbsorber.Text);
            }

            var result = new
            {
                totalPages = document.Pages.Count,
                extractedPages = pagesToExtract,
                truncated,
                content = contentBuilder.ToString()
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }

    /// <summary>
    ///     Retrieves statistics about the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <returns>A JSON string containing document statistics.</returns>
    /// <exception cref="ArgumentException">Thrown when path is required but not provided.</exception>
    private string GetStatistics(string? sessionId, string? path)
    {
        if (!string.IsNullOrEmpty(sessionId))
        {
            using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
            var document = ctx.Document;

            var totalAnnotations = 0;
            var totalParagraphs = 0;
            for (var i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                totalAnnotations += page.Annotations.Count;
                totalParagraphs += page.Paragraphs.Count;
            }

            var result = new
            {
                totalPages = document.Pages.Count,
                isEncrypted = document.IsEncrypted,
                isLinearized = document.IsLinearized,
                bookmarks = document.Outlines.Count,
                formFields = document.Form?.Count ?? 0,
                totalAnnotations,
                totalParagraphs,
                note = "File size info not available in session mode"
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
        else
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("path is required for get_statistics operation");

            SecurityHelper.ValidateFilePath(path, "path", true);

            using var document = new Document(path);
            var fileInfo = new FileInfo(path);

            var totalAnnotations = 0;
            var totalParagraphs = 0;
            for (var i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                totalAnnotations += page.Annotations.Count;
                totalParagraphs += page.Paragraphs.Count;
            }

            var result = new
            {
                fileSizeBytes = fileInfo.Length,
                fileSizeKb = Math.Round(fileInfo.Length / 1024.0, 2),
                totalPages = document.Pages.Count,
                isEncrypted = document.IsEncrypted,
                isLinearized = document.IsLinearized,
                bookmarks = document.Outlines.Count,
                formFields = document.Form?.Count ?? 0,
                totalAnnotations,
                totalParagraphs
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }
    }
}