using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfGetPageInfoTool : IAsposeTool
{
    public string Description => "Get page information from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, optional, if not provided returns info for all pages)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        using var document = new Document(path);
        var sb = new StringBuilder();

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            {
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            }

            var page = document.Pages[pageIndex.Value];
            sb.AppendLine($"Page {pageIndex.Value} Information:");
            sb.AppendLine($"  Width: {page.PageInfo.Width}");
            sb.AppendLine($"  Height: {page.PageInfo.Height}");
            sb.AppendLine($"  Rotation: {page.Rotate}");
            sb.AppendLine($"  Paragraphs Count: {page.Paragraphs.Count}");
            sb.AppendLine($"  Annotations Count: {page.Annotations.Count}");
        }
        else
        {
            sb.AppendLine($"Total Pages: {document.Pages.Count}");
            sb.AppendLine();
            for (int i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                sb.AppendLine($"Page {i}:");
                sb.AppendLine($"  Width: {page.PageInfo.Width}, Height: {page.PageInfo.Height}");
                sb.AppendLine($"  Paragraphs: {page.Paragraphs.Count}, Annotations: {page.Annotations.Count}");
                sb.AppendLine();
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

