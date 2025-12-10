using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfDeleteLinkTool : IAsposeTool
{
    public string Description => "Delete a link from PDF document";

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
                description = "Page index (1-based)"
            },
            linkIndex = new
            {
                type = "number",
                description = "Link index (0-based, use pdf_get_links to find index)"
            }
        },
        required = new[] { "path", "pageIndex", "linkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var linkIndex = arguments?["linkIndex"]?.GetValue<int>() ?? throw new ArgumentException("linkIndex is required");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();
        
        if (linkIndex < 0 || linkIndex >= links.Count)
        {
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");
        }

        var link = links[linkIndex];
        page.Annotations.Delete(link);
        document.Save(path);

        return await Task.FromResult($"Link deleted from page {pageIndex}: {path}");
    }
}

