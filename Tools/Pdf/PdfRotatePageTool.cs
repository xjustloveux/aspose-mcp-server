using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfRotatePageTool : IAsposeTool
{
    public string Description => "Rotate page(s) in PDF document";

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
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (0, 90, 180, or 270)"
            },
            pageIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of page indices to rotate (1-based, optional, overrides pageIndex)"
            }
        },
        required = new[] { "path", "rotation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();
        var rotation = arguments?["rotation"]?.GetValue<int>() ?? throw new ArgumentException("rotation is required");
        var pageIndicesArray = arguments?["pageIndices"]?.AsArray();

        if (rotation != 0 && rotation != 90 && rotation != 180 && rotation != 270)
        {
            throw new ArgumentException("rotation must be 0, 90, 180, or 270");
        }

        using var document = new Document(path);
        Rotation rotationEnum;
        switch (rotation)
        {
            case 90:
                rotationEnum = Rotation.on90;
                break;
            case 180:
                rotationEnum = Rotation.on180;
                break;
            case 270:
                rotationEnum = Rotation.on270;
                break;
            default:
                rotationEnum = Rotation.None;
                break;
        }

        List<int> pagesToRotate;
        if (pageIndicesArray != null && pageIndicesArray.Count > 0)
        {
            pagesToRotate = pageIndicesArray.Select(p => p?.GetValue<int>()).Where(p => p.HasValue).Select(p => p!.Value).ToList();
        }
        else if (pageIndex.HasValue)
        {
            pagesToRotate = new List<int> { pageIndex.Value };
        }
        else
        {
            // Rotate all pages
            pagesToRotate = Enumerable.Range(1, document.Pages.Count).ToList();
        }

        foreach (var pageNum in pagesToRotate)
        {
            if (pageNum < 1 || pageNum > document.Pages.Count)
            {
                continue;
            }
            document.Pages[pageNum].Rotate = rotationEnum;
        }

        document.Save(path);
        return await Task.FromResult($"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees: {path}");
    }
}

