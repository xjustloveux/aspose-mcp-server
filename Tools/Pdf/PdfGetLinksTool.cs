using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfGetLinksTool : IAsposeTool
{
    public string Description => "Get all links (hyperlinks) from PDF document";

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
                description = "Page index (1-based, optional, if not provided gets all pages)"
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
            sb.AppendLine($"=== Page {pageIndex.Value} Links ===");
            GetLinksFromPage(page, pageIndex.Value, sb);
        }
        else
        {
            sb.AppendLine("=== All Links ===");
            for (int i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                var linkCount = page.Annotations.Count(a => a is LinkAnnotation);
                if (linkCount > 0)
                {
                    sb.AppendLine($"\nPage {i}: {linkCount} link(s)");
                    GetLinksFromPage(page, i, sb);
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private void GetLinksFromPage(Page page, int pageNum, StringBuilder sb)
    {
        int index = 0;
        foreach (var annotation in page.Annotations)
        {
            if (annotation is LinkAnnotation link)
            {
                sb.AppendLine($"  [{index}] Link:");
                sb.AppendLine($"    Position: ({link.Rect.LLX}, {link.Rect.LLY}) Size: ({link.Rect.Width}, {link.Rect.Height})");
                if (link.Action is GoToURIAction uriAction)
                {
                    sb.AppendLine($"    Type: External URL");
                    sb.AppendLine($"    URL: {uriAction.URI}");
                }
                else if (link.Action is GoToAction goToAction)
                {
                    sb.AppendLine($"    Type: Internal link");
                    // Note: Destination structure may vary
                    sb.AppendLine($"    Target: Internal page reference");
                }
                else
                {
                    sb.AppendLine($"    Type: {link.Action?.GetType().Name ?? "Unknown"}");
                }
                index++;
            }
        }
    }
}

