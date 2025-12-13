using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfInfoTool : IAsposeTool
{
    public string Description => @"Get content and statistics from PDF documents. Supports 2 operations: get_content, get_statistics.

Usage examples:
- Get content: pdf_info(operation='get_content', path='doc.pdf', pageIndex=1)
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "get_content" => await GetContent(arguments),
            "get_statistics" => await GetStatistics(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> GetContent(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var sb = new StringBuilder();

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var textAbsorber = new Aspose.Pdf.Text.TextAbsorber();
            document.Pages[pageIndex.Value].Accept(textAbsorber);
            sb.AppendLine($"=== Content from Page {pageIndex.Value} ===");
            sb.AppendLine(textAbsorber.Text);
        }
        else
        {
            var textAbsorber = new Aspose.Pdf.Text.TextAbsorber();
            document.Pages.Accept(textAbsorber);
            sb.AppendLine("=== Full Document Content ===");
            sb.AppendLine(textAbsorber.Text);
        }

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetStatistics(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var fileInfo = new FileInfo(path);
        var sb = new StringBuilder();

        sb.AppendLine("=== PDF Statistics ===");
        sb.AppendLine($"File Size: {fileInfo.Length} bytes ({fileInfo.Length / 1024.0:F2} KB)");
        sb.AppendLine($"Total Pages: {document.Pages.Count}");
        sb.AppendLine($"Is Encrypted: {document.IsEncrypted}");
        sb.AppendLine($"Is Linearized: {document.IsLinearized}");
        sb.AppendLine($"Bookmarks: {document.Outlines.Count}");
        sb.AppendLine($"Form Fields: {document.Form?.Count ?? 0}");

        int totalAnnotations = 0;
        for (int i = 1; i <= document.Pages.Count; i++)
            totalAnnotations += document.Pages[i].Annotations.Count;
        sb.AppendLine($"Total Annotations: {totalAnnotations}");

        int totalParagraphs = 0;
        for (int i = 1; i <= document.Pages.Count; i++)
            totalParagraphs += document.Pages[i].Paragraphs.Count;
        sb.AppendLine($"Total Paragraphs: {totalParagraphs}");

        return await Task.FromResult(sb.ToString());
    }
}

