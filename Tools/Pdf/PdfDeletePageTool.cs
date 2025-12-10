using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfDeletePageTool : IAsposeTool
{
    public string Description => "Delete page(s) from PDF document";

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
                description = "Page index to delete (1-based)"
            },
            pageIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of page indices to delete (1-based, optional, overrides pageIndex)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();
        var pageIndicesArray = arguments?["pageIndices"]?.AsArray();

        using var document = new Document(path);
        var totalPages = document.Pages.Count;

        List<int> pagesToDelete;
        if (pageIndicesArray != null && pageIndicesArray.Count > 0)
        {
            pagesToDelete = pageIndicesArray.Select(p => p?.GetValue<int>()).Where(p => p.HasValue).Select(p => p!.Value).OrderByDescending(p => p).ToList();
        }
        else if (pageIndex.HasValue)
        {
            pagesToDelete = new List<int> { pageIndex.Value };
        }
        else
        {
            throw new ArgumentException("Either pageIndex or pageIndices must be provided");
        }

        foreach (var pageNum in pagesToDelete)
        {
            if (pageNum < 1 || pageNum > totalPages)
            {
                continue;
            }
            document.Pages.Delete(pageNum);
        }

        document.Save(path);
        return await Task.FromResult($"Deleted {pagesToDelete.Count} page(s). Remaining pages: {document.Pages.Count}");
    }
}

