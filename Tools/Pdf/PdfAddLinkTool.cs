using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfAddLinkTool : IAsposeTool
{
    public string Description => "Add hyperlink to PDF document";

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
            x = new
            {
                type = "number",
                description = "X position of link area"
            },
            y = new
            {
                type = "number",
                description = "Y position of link area"
            },
            width = new
            {
                type = "number",
                description = "Width of link area"
            },
            height = new
            {
                type = "number",
                description = "Height of link area"
            },
            url = new
            {
                type = "string",
                description = "URL to link to (optional, if not provided uses targetPage)"
            },
            targetPage = new
            {
                type = "number",
                description = "Target page number (1-based, optional, if not provided uses url)"
            }
        },
        required = new[] { "path", "pageIndex", "x", "y", "width", "height" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var x = arguments?["x"]?.GetValue<double>() ?? throw new ArgumentException("x is required");
        var y = arguments?["y"]?.GetValue<double>() ?? throw new ArgumentException("y is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");
        var url = arguments?["url"]?.GetValue<string>();
        var targetPage = arguments?["targetPage"]?.GetValue<int?>();

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var linkRect = new Rectangle(x, y, x + width, y + height);

        LinkAnnotation link;
        if (!string.IsNullOrEmpty(url))
        {
            link = new LinkAnnotation(page, linkRect)
            {
                Action = new GoToURIAction(url)
            };
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
            {
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            }
            link = new LinkAnnotation(page, linkRect)
            {
                Action = new GoToAction(document.Pages[targetPage.Value])
            };
        }
        else
        {
            throw new ArgumentException("Either url or targetPage must be provided");
        }

        page.Annotations.Add(link);
        document.Save(path);
        return await Task.FromResult($"Link added to page {pageIndex}: {path}");
    }
}

