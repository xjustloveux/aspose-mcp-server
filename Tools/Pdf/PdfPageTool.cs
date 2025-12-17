using System.Text;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing pages in PDF documents (add, delete, insert, extract, rotate, resize)
/// </summary>
public class PdfPageTool : IAsposeTool
{
    public string Description =>
        @"Manage pages in PDF documents. Supports 5 operations: add, delete, rotate, get_details, get_info.

Usage examples:
- Add page: pdf_page(operation='add', path='doc.pdf', count=1)
- Delete page: pdf_page(operation='delete', path='doc.pdf', pageIndex=1)
- Rotate page: pdf_page(operation='rotate', path='doc.pdf', pageIndex=1, angle=90)
- Get page details: pdf_page(operation='get_details', path='doc.pdf', pageIndex=1)
- Get page info: pdf_page(operation='get_info', path='doc.pdf')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add page(s) (required params: path)
- 'delete': Delete a page (required params: path, pageIndex)
- 'rotate': Rotate a page (required params: path, pageIndex, angle)
- 'get_details': Get page details (required params: path, pageIndex)
- 'get_info': Get all pages info (required params: path)",
                @enum = new[] { "add", "delete", "rotate", "get_details", "get_info" }
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
        var operation = ArgumentHelper.GetString(arguments, "operation");

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

    /// <summary>
    ///     Adds a page to the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, optional pageIndex, outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddPage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var count = ArgumentHelper.GetInt(arguments, "count", 1);
        var insertAt = ArgumentHelper.GetIntNullable(arguments, "insertAt");
        var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
        var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);

        for (var i = 0; i < count; i++)
        {
            var page = document.Pages.Add();
            if (width.HasValue && height.HasValue)
                page.SetPageSize(width.Value, height.Value);
            else
                page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
        }

        if (insertAt is >= 1 && insertAt.Value <= document.Pages.Count)
        {
            // Move pages manually since GetRange may not be available
            var pagesToMove = new List<Page>();
            for (var i = document.Pages.Count - count; i < document.Pages.Count; i++)
                pagesToMove.Add(document.Pages[i]);
            foreach (var page in pagesToMove)
            {
                document.Pages.Remove(page);
                document.Pages.Insert(insertAt.Value - 1, page);
            }
        }

        document.Save(outputPath);
        return await Task.FromResult($"Successfully added {count} page(s). Output: {outputPath}");
    }

    /// <summary>
    ///     Deletes a page from the PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> DeletePage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        document.Pages.Delete(pageIndex);
        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully deleted page {pageIndex}. Remaining pages: {document.Pages.Count}. Output: {outputPath}");
    }

    /// <summary>
    ///     Rotates a page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, angle, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> RotatePage(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var rotation = ArgumentHelper.GetInt(arguments, "rotation");
        var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");
        var pageIndicesArray = ArgumentHelper.GetArray(arguments, "pageIndices", false);

        SecurityHelper.ValidateFilePath(path);
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
        if (pageIndicesArray is { Count: > 0 })
            pagesToRotate = pageIndicesArray.Select(p => p?.GetValue<int>()).Where(p => p.HasValue)
                .Select(p => p!.Value).ToList();
        else if (pageIndex.HasValue)
            pagesToRotate = [pageIndex.Value];
        else
            pagesToRotate = Enumerable.Range(1, document.Pages.Count).ToList();

        foreach (var pageNum in pagesToRotate)
            if (pageNum >= 1 && pageNum <= document.Pages.Count)
                document.Pages[pageNum].Rotate = rotationEnum;

        document.Save(outputPath);
        return await Task.FromResult(
            $"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees. Output: {outputPath}");
    }

    /// <summary>
    ///     Gets detailed information about a page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex</param>
    /// <returns>Formatted string with page details</returns>
    private async Task<string> GetPageDetails(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

        SecurityHelper.ValidateFilePath(path);

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var sb = new StringBuilder();
        sb.AppendLine($"=== Page {pageIndex} Details ===");
        sb.AppendLine($"Size: {page.Rect.Width} x {page.Rect.Height}");
        sb.AppendLine($"Rotation: {page.Rotate}°");
        sb.AppendLine(
            $"MediaBox: ({page.MediaBox.LLX}, {page.MediaBox.LLY}) to ({page.MediaBox.URX}, {page.MediaBox.URY})");
        sb.AppendLine($"CropBox: ({page.CropBox.LLX}, {page.CropBox.LLY}) to ({page.CropBox.URX}, {page.CropBox.URY})");
        sb.AppendLine($"Annotations: {page.Annotations.Count}");
        sb.AppendLine($"Paragraphs: {page.Paragraphs.Count}");
        sb.AppendLine($"Images: {page.Resources?.Images?.Count ?? 0}");

        return await Task.FromResult(sb.ToString());
    }

    /// <summary>
    ///     Gets information about all pages
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <returns>Formatted string with page information</returns>
    private async Task<string> GetPageInfo(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Page Info ===");
        sb.AppendLine($"Total Pages: {document.Pages.Count}");
        sb.AppendLine();

        for (var i = 1; i <= Math.Min(document.Pages.Count, 10); i++)
        {
            var page = document.Pages[i];
            sb.AppendLine($"Page {i}: {page.Rect.Width} x {page.Rect.Height}, Rotation: {page.Rotate}°");
        }

        if (document.Pages.Count > 10)
            sb.AppendLine($"... ({document.Pages.Count - 10} more pages)");

        return await Task.FromResult(sb.ToString());
    }
}