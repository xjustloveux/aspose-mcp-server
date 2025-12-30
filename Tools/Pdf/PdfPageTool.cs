using System.Text.Json;
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        // Only get outputPath for operations that modify the document
        string? outputPath = null;
        if (operation.ToLower() != "get_details" && operation.ToLower() != "get_info")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddPage(path, outputPath!, arguments),
            "delete" => await DeletePage(path, outputPath!, arguments),
            "rotate" => await RotatePage(path, outputPath!, arguments),
            "get_details" => await GetPageDetails(path, arguments),
            "get_info" => await GetPageInfo(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a page to the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional count, insertAt, width, height</param>
    /// <returns>Success message</returns>
    private Task<string> AddPage(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var count = ArgumentHelper.GetInt(arguments, "count", 1);
            var insertAt = ArgumentHelper.GetIntNullable(arguments, "insertAt");
            var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
            var height = ArgumentHelper.GetDoubleNullable(arguments, "height");

            using var document = new Document(path);

            if (insertAt is >= 1 && insertAt.Value <= document.Pages.Count)
                // Insert pages at specific position
                for (var i = 0; i < count; i++)
                {
                    var page = document.Pages.Insert(insertAt.Value + i);
                    if (width.HasValue && height.HasValue)
                        page.SetPageSize(width.Value, height.Value);
                    else
                        page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
                }
            else
                // Append pages at the end
                for (var i = 0; i < count; i++)
                {
                    var page = document.Pages.Add();
                    if (width.HasValue && height.HasValue)
                        page.SetPageSize(width.Value, height.Value);
                    else
                        page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
                }

            document.Save(outputPath);
            return $"Added {count} page(s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a page from the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeletePage(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            document.Pages.Delete(pageIndex);
            document.Save(outputPath);
            return $"Deleted page {pageIndex} (remaining: {document.Pages.Count}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Rotates a page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing rotation, optional pageIndex, pageIndices</param>
    /// <returns>Success message</returns>
    private Task<string> RotatePage(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rotation = ArgumentHelper.GetInt(arguments, "rotation");
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");
            var pageIndicesArray = ArgumentHelper.GetArray(arguments, "pageIndices", false);

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
            return $"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets detailed information about a page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex</param>
    /// <returns>JSON string with page details</returns>
    private Task<string> GetPageDetails(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var result = new
            {
                pageIndex,
                width = page.Rect.Width,
                height = page.Rect.Height,
                rotation = page.Rotate.ToString(),
                mediaBox = new
                {
                    llx = page.MediaBox.LLX,
                    lly = page.MediaBox.LLY,
                    urx = page.MediaBox.URX,
                    ury = page.MediaBox.URY
                },
                cropBox = new
                {
                    llx = page.CropBox.LLX,
                    lly = page.CropBox.LLY,
                    urx = page.CropBox.URX,
                    ury = page.CropBox.URY
                },
                annotations = page.Annotations.Count,
                paragraphs = page.Paragraphs.Count,
                images = page.Resources?.Images?.Count ?? 0
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets information about all pages
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <returns>JSON string with page information</returns>
    private Task<string> GetPageInfo(string path)
    {
        return Task.Run(() =>
        {
            using var document = new Document(path);
            var pageList = new List<object>();

            for (var i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                pageList.Add(new
                {
                    pageIndex = i,
                    width = page.Rect.Width,
                    height = page.Rect.Height,
                    rotation = page.Rotate.ToString()
                });
            }

            var result = new
            {
                count = document.Pages.Count,
                items = pageList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}