using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfGetStatisticsTool : IAsposeTool
{
    public string Description => "Get PDF document statistics";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var document = new Document(path);
        var fileInfo = new FileInfo(path);
        var sb = new StringBuilder();

        sb.AppendLine("PDF Statistics:");
        sb.AppendLine($"  File Size: {fileInfo.Length} bytes ({fileInfo.Length / 1024.0:F2} KB)");
        sb.AppendLine($"  Total Pages: {document.Pages.Count}");
        sb.AppendLine($"  Is Encrypted: {document.IsEncrypted}");
        sb.AppendLine($"  Is Linearized: {document.IsLinearized}");
        sb.AppendLine($"  Bookmarks: {document.Outlines.Count}");
        sb.AppendLine($"  Form Fields: {document.Form?.Count ?? 0}");

        var totalAnnotations = 0;
        for (int i = 1; i <= document.Pages.Count; i++)
        {
            totalAnnotations += document.Pages[i].Annotations.Count;
        }
        sb.AppendLine($"  Total Annotations: {totalAnnotations}");

        var totalParagraphs = 0;
        for (int i = 1; i <= document.Pages.Count; i++)
        {
            totalParagraphs += document.Pages[i].Paragraphs.Count;
        }
        sb.AppendLine($"  Total Paragraphs: {totalParagraphs}");

        return await Task.FromResult(sb.ToString());
    }
}

