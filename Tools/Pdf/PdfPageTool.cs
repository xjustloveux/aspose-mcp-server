using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfPageTool : IAsposeTool
{
    public string Description => "Manage pages in PDF documents (add, delete, rotate, get details, get info)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add, delete, rotate, get_details, get_info",
                @enum = new[] { "add", "delete", "rotate", "get_details", "get_info" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            count = new
            {
                type = "number",
                description = "Number of pages to add (for add, default: 1)"
            },
            insertAt = new
            {
                type = "number",
                description = "Position to insert pages (1-based, for add, optional, default: append at end)"
            },
            width = new
            {
                type = "number",
                description = "Page width in points (for add, optional)"
            },
            height = new
            {
                type = "number",
                description = "Page height in points (for add, optional)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for delete, rotate, get_details)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required)"
            },
            pageIndices = new
            {
                type = "array",
                description = "Array of page indices to rotate (1-based, for rotate, optional)",
                items = new { type = "number" }
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add" => await AddPage(arguments),
            "delete" => await DeletePage(arguments),
            "rotate" => await RotatePage(arguments),
            "get_details" => await GetPageDetails(arguments),
            "get_info" => await GetPageInfo(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddPage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var count = arguments?["count"]?.GetValue<int?>() ?? 1;
        var insertAt = arguments?["insertAt"]?.GetValue<int?>();
        var width = arguments?["width"]?.GetValue<double?>();
        var height = arguments?["height"]?.GetValue<double?>();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);

        for (int i = 0; i < count; i++)
        {
            var page = document.Pages.Add();
            if (width.HasValue && height.HasValue)
                page.SetPageSize(width.Value, height.Value);
            else
                page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
        }

        if (insertAt.HasValue && insertAt.Value >= 1 && insertAt.Value <= document.Pages.Count)
        {
            // Move pages manually since GetRange may not be available
            var pagesToMove = new List<Page>();
            for (int i = document.Pages.Count - count; i < document.Pages.Count; i++)
            {
                pagesToMove.Add(document.Pages[i]);
            }
            foreach (var page in pagesToMove)
            {
                document.Pages.Remove(page);
                document.Pages.Insert(insertAt.Value - 1, page);
            }
        }

        document.Save(outputPath);
        return await Task.FromResult($"Successfully added {count} page(s). Output: {outputPath}");
    }

    private async Task<string> DeletePage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        document.Pages.Delete(pageIndex);
        document.Save(outputPath);
        return await Task.FromResult($"Successfully deleted page {pageIndex}. Remaining pages: {document.Pages.Count}. Output: {outputPath}");
    }

    private async Task<string> RotatePage(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var rotation = arguments?["rotation"]?.GetValue<int>() ?? throw new ArgumentException("rotation is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();
        var pageIndicesArray = arguments?["pageIndices"]?.AsArray();

        SecurityHelper.ValidateFilePath(path, "path");
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        if (rotation != 0 && rotation != 90 && rotation != 180 && rotation != 270)
            throw new ArgumentException("rotation must be 0, 90, 180, or 270");

        using var document = new Document(path);
        var rotationEnum = rotation switch
        {
            90 => Rotation.on90,
            180 => Rotation.on180,
            270 => Rotation.on270,
            _ => Rotation.None
        };

        List<int> pagesToRotate;
        if (pageIndicesArray != null && pageIndicesArray.Count > 0)
            pagesToRotate = pageIndicesArray.Select(p => p?.GetValue<int>()).Where(p => p.HasValue).Select(p => p!.Value).ToList();
        else if (pageIndex.HasValue)
            pagesToRotate = new List<int> { pageIndex.Value };
        else
            pagesToRotate = Enumerable.Range(1, document.Pages.Count).ToList();

        foreach (var pageNum in pagesToRotate)
        {
            if (pageNum >= 1 && pageNum <= document.Pages.Count)
                document.Pages[pageNum].Rotate = rotationEnum;
        }

        document.Save(outputPath);
        return await Task.FromResult($"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees. Output: {outputPath}");
    }

    private async Task<string> GetPageDetails(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var sb = new StringBuilder();
        sb.AppendLine($"=== Page {pageIndex} Details ===");
        sb.AppendLine($"Size: {page.Rect.Width} x {page.Rect.Height}");
        sb.AppendLine($"Rotation: {page.Rotate}°");
        sb.AppendLine($"MediaBox: ({page.MediaBox.LLX}, {page.MediaBox.LLY}) to ({page.MediaBox.URX}, {page.MediaBox.URY})");
        sb.AppendLine($"CropBox: ({page.CropBox.LLX}, {page.CropBox.LLY}) to ({page.CropBox.URX}, {page.CropBox.URY})");
        sb.AppendLine($"Annotations: {page.Annotations.Count}");
        sb.AppendLine($"Paragraphs: {page.Paragraphs.Count}");
        sb.AppendLine($"Images: {page.Resources?.Images?.Count ?? 0}");

        return await Task.FromResult(sb.ToString());
    }

    private async Task<string> GetPageInfo(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        SecurityHelper.ValidateFilePath(path, "path");

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine($"=== PDF Page Info ===");
        sb.AppendLine($"Total Pages: {document.Pages.Count}");
        sb.AppendLine();

        for (int i = 1; i <= Math.Min(document.Pages.Count, 10); i++)
        {
            var page = document.Pages[i];
            sb.AppendLine($"Page {i}: {page.Rect.Width} x {page.Rect.Height}, Rotation: {page.Rotate}°");
        }

        if (document.Pages.Count > 10)
            sb.AppendLine($"... ({document.Pages.Count - 10} more pages)");

        return await Task.FromResult(sb.ToString());
    }
}

