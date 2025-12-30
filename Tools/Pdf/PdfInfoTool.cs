using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for getting content and statistics from PDF documents
/// </summary>
public class PdfInfoTool : IAsposeTool
{
    public string Description =>
        @"Get content and statistics from PDF documents. Supports 2 operations: get_content, get_statistics.

Usage examples:
- Get content from page: pdf_info(operation='get_content', path='doc.pdf', pageIndex=1)
- Get content from all pages: pdf_info(operation='get_content', path='doc.pdf')
- Get content with limit: pdf_info(operation='get_content', path='doc.pdf', maxPages=50)
- Get statistics: pdf_info(operation='get_statistics', path='doc.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'get_content': Get text content from page(s) (required params: path)
- 'get_statistics': Get document statistics (required params: path)",
                @enum = new[] { "get_content", "get_statistics" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, optional for get_content, extracts all if not specified)"
            },
            maxPages = new
            {
                type = "number",
                description = "Maximum pages to extract (for get_content without pageIndex, default: 100)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "get_content" => await GetContent(arguments, path),
            "get_statistics" => await GetStatistics(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets PDF content as text
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional pageIndex and maxPages</param>
    /// <param name="path">Input file path</param>
    /// <returns>PDF content as JSON string</returns>
    private Task<string> GetContent(JsonObject? arguments, string path)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");
            var maxPages = ArgumentHelper.GetInt(arguments, "maxPages", 100);

            using var document = new Document(path);

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
        });
    }

    /// <summary>
    ///     Gets PDF statistics
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with statistics</returns>
    private Task<string> GetStatistics(string path)
    {
        return Task.Run(() =>
        {
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
        });
    }
}