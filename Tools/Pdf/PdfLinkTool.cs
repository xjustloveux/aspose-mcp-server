using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfLinkTool : IAsposeTool
{
    public string Description => @"Manage links in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add link: pdf_link(operation='add', path='doc.pdf', pageIndex=1, x=100, y=100, width=200, height=30, url='https://example.com')
- Delete link: pdf_link(operation='delete', path='doc.pdf', pageIndex=1, linkIndex=0)
- Edit link: pdf_link(operation='edit', path='doc.pdf', pageIndex=1, linkIndex=0, url='https://newurl.com')
- Get links: pdf_link(operation='get', path='doc.pdf', pageIndex=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a link (required params: path, pageIndex, x, y, width, height, url)
- 'delete': Delete a link (required params: path, pageIndex, linkIndex)
- 'edit': Edit a link (required params: path, pageIndex, linkIndex, url)
- 'get': Get all links (required params: path, pageIndex)",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, delete, edit)"
            },
            linkIndex = new
            {
                type = "number",
                description = "Link index (0-based, required for delete, edit)"
            },
            x = new
            {
                type = "number",
                description = "X position of link area (required for add)"
            },
            y = new
            {
                type = "number",
                description = "Y position of link area (required for add)"
            },
            width = new
            {
                type = "number",
                description = "Width of link area (required for add)"
            },
            height = new
            {
                type = "number",
                description = "Height of link area (required for add)"
            },
            url = new
            {
                type = "string",
                description = "URL to link to (for add, edit)"
            },
            targetPage = new
            {
                type = "number",
                description = "Target page number (1-based, for add, edit)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddLink(arguments),
            "delete" => await DeleteLink(arguments),
            "edit" => await EditLink(arguments),
            "get" => await GetLinks(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddLink(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var x = arguments?["x"]?.GetValue<double>() ?? throw new ArgumentException("x is required");
        var y = arguments?["y"]?.GetValue<double>() ?? throw new ArgumentException("y is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");
        var url = arguments?["url"]?.GetValue<string>();
        var targetPage = arguments?["targetPage"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var rect = new Aspose.Pdf.Rectangle(x, y, x + width, y + height);
        LinkAnnotation link;

        if (!string.IsNullOrEmpty(url))
        {
            link = new LinkAnnotation(page, rect) { Action = new GoToURIAction(url) };
        }
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            link = new LinkAnnotation(page, rect) { Action = new GoToAction(document.Pages[targetPage.Value]) };
        }
        else
            throw new ArgumentException("Either url or targetPage must be provided");

        page.Annotations.Add(link);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully added link to page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> DeleteLink(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var linkIndex = arguments?["linkIndex"]?.GetValue<int>() ?? throw new ArgumentException("linkIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();
        if (linkIndex < 0 || linkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        page.Annotations.Delete(links[linkIndex]);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted link {linkIndex} from page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> EditLink(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var linkIndex = arguments?["linkIndex"]?.GetValue<int>() ?? throw new ArgumentException("linkIndex is required");
        var url = arguments?["url"]?.GetValue<string>();
        var targetPage = arguments?["targetPage"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var links = page.Annotations.OfType<LinkAnnotation>().ToList();
        if (linkIndex < 0 || linkIndex >= links.Count)
            throw new ArgumentException($"linkIndex must be between 0 and {links.Count - 1}");

        var link = links[linkIndex];
        if (!string.IsNullOrEmpty(url))
            link.Action = new GoToURIAction(url);
        else if (targetPage.HasValue)
        {
            if (targetPage.Value < 1 || targetPage.Value > document.Pages.Count)
                throw new ArgumentException($"targetPage must be between 1 and {document.Pages.Count}");
            link.Action = new GoToAction(document.Pages[targetPage.Value]);
        }

        document.Save(outputPath);
        return await Task.FromResult($"Successfully edited link {linkIndex} on page {pageIndex}. Output: {outputPath}");
    }

    private async Task<string> GetLinks(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Links ===");
        sb.AppendLine();

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex.Value];
            var links = page.Annotations.OfType<LinkAnnotation>().ToList();
            sb.AppendLine($"Page {pageIndex.Value} Links ({links.Count}):");
            for (int i = 0; i < links.Count; i++)
            {
                var link = links[i];
                sb.AppendLine($"  [{i}] Position: ({link.Rect.LLX}, {link.Rect.LLY})");
                if (link.Action is GoToURIAction uriAction)
                    sb.AppendLine($"      URL: {uriAction.URI}");
                else if (link.Action is GoToAction goToAction)
                    sb.AppendLine($"      Target: Page");
                sb.AppendLine();
            }
        }
        else
        {
            int totalCount = 0;
            for (int p = 1; p <= document.Pages.Count; p++)
            {
                var page = document.Pages[p];
                var links = page.Annotations.OfType<LinkAnnotation>().ToList();
                if (links.Count > 0)
                {
                    sb.AppendLine($"Page {p} ({links.Count} links):");
                    for (int i = 0; i < links.Count; i++)
                    {
                        var link = links[i];
                        if (link.Action is GoToURIAction uriAction)
                            sb.AppendLine($"  [{i}] URL: {uriAction.URI}");
                        else if (link.Action is GoToAction)
                            sb.AppendLine($"  [{i}] Target: Page");
                    }
                    sb.AppendLine();
                    totalCount += links.Count;
                }
            }
            sb.Insert(0, $"Total Links: {totalCount}\n\n");
        }

        return await Task.FromResult(sb.ToString());
    }
}

