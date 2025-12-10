using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfAddPageTool : IAsposeTool
{
    public string Description => "Add new page(s) to PDF document";

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
            count = new
            {
                type = "number",
                description = "Number of pages to add (optional, default: 1)"
            },
            insertAt = new
            {
                type = "number",
                description = "Position to insert pages (1-based, optional, default: append at end)"
            },
            width = new
            {
                type = "number",
                description = "Page width in points (optional)"
            },
            height = new
            {
                type = "number",
                description = "Page height in points (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var count = arguments?["count"]?.GetValue<int?>() ?? 1;
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        using var document = new Document(path);
        var totalPages = document.Pages.Count;

        if (insertAt.HasValue)
        {
            if (insertAt.Value < 1 || insertAt.Value > totalPages + 1)
            {
                throw new ArgumentException($"insertAt must be between 1 and {totalPages + 1}");
            }

            for (int i = 0; i < count; i++)
            {
                var page = document.Pages.Insert(insertAt.Value + i);
                if (width.HasValue && height.HasValue)
                {
                    page.SetPageSize(width.Value, height.Value);
                }
            }
        }
        else
        {
            for (int i = 0; i < count; i++)
            {
                var page = document.Pages.Add();
                if (width.HasValue && height.HasValue)
                {
                    page.SetPageSize(width.Value, height.Value);
                }
            }
        }

        document.Save(path);
        return await Task.FromResult($"Added {count} page(s). Total pages: {document.Pages.Count}");
    }
}

