using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfEditLinkTool : IAsposeTool
{
    public string Description => "Edit link properties (URL, target page, position, size)";

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
            },
            url = new
            {
                type = "string",
                description = "New URL (optional, for external links)"
            },
            targetPage = new
            {
                type = "number",
                description = "New target page (1-based, optional, for internal links)"
            },
            x = new
            {
                type = "number",
                description = "New X position (optional)"
            },
            y = new
            {
                type = "number",
                description = "New Y position (optional)"
            },
            width = new
            {
                type = "number",
                description = "New width (optional)"
            },
            height = new
            {
                type = "number",
                description = "New height (optional)"
            }
        },
        required = new[] { "path", "pageIndex", "linkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var linkIndex = arguments?["linkIndex"]?.GetValue<int>() ?? throw new ArgumentException("linkIndex is required");
        var url = arguments?["url"]?.GetValue<string>();
        var targetPage = arguments?["targetPage"]?.GetValue<int?>();
        var x = arguments?["x"]?.GetValue<double?>();
        var y = arguments?["y"]?.GetValue<double?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

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
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(url))
        {
            link.Action = new GoToURIAction(url);
            changes.Add($"URL: {url}");
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
            {
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            }
            link.Action = new GoToAction(document.Pages[targetPage.Value]);
            changes.Add($"Target page: {targetPage.Value}");
        }

        if (x.HasValue || y.HasValue || width.HasValue || height.HasValue)
        {
            var rect = link.Rect;
            var newX = x ?? rect.LLX;
            var newY = y ?? rect.LLY;
            var newWidth = width ?? rect.Width;
            var newHeight = height ?? rect.Height;
            link.Rect = new Rectangle(newX, newY, newX + newWidth, newY + newHeight);
            changes.Add("Position/Size updated");
        }

        document.Save(path);
        return await Task.FromResult($"Link edited: {string.Join(", ", changes)} - {path}");
    }
}

