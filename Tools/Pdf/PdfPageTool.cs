using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing pages in PDF documents (add, delete, insert, extract, rotate, resize)
/// </summary>
[McpServerToolType]
public class PdfPageTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfPageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfPageTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_page")]
    [Description(@"Manage pages in PDF documents. Supports 5 operations: add, delete, rotate, get_details, get_info.

Usage examples:
- Add page: pdf_page(operation='add', path='doc.pdf', count=1)
- Delete page: pdf_page(operation='delete', path='doc.pdf', pageIndex=1)
- Rotate page: pdf_page(operation='rotate', path='doc.pdf', pageIndex=1, rotation=90)
- Get page details: pdf_page(operation='get_details', path='doc.pdf', pageIndex=1)
- Get page info: pdf_page(operation='get_info', path='doc.pdf')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add page(s) (required params: path)
- 'delete': Delete a page (required params: path, pageIndex)
- 'rotate': Rotate a page (required params: path, pageIndex, rotation)
- 'get_details': Get page details (required params: path, pageIndex)
- 'get_info': Get all pages info (required params: path)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to overwrite input)")]
        string? outputPath = null,
        [Description("Number of pages to add (for add, default: 1)")]
        int count = 1,
        [Description("Position to insert pages (1-based, for add, optional, default: append at end)")]
        int? insertAt = null,
        [Description("Page width in points (for add, optional)")]
        double? width = null,
        [Description("Page height in points (for add, optional)")]
        double? height = null,
        [Description("Page index (1-based, required for delete, rotate, get_details)")]
        int pageIndex = 0,
        [Description("Rotation angle in degrees: 0, 90, 180, 270 (for rotate, required)")]
        int rotation = 0,
        [Description("Array of page indices to rotate (1-based, for rotate, optional)")]
        int[]? pageIndices = null)
    {
        return operation.ToLower() switch
        {
            "add" => AddPage(sessionId, path, outputPath, count, insertAt, width, height),
            "delete" => DeletePage(sessionId, path, outputPath, pageIndex),
            "rotate" => RotatePage(sessionId, path, outputPath, rotation, pageIndex, pageIndices),
            "get_details" => GetPageDetails(sessionId, path, pageIndex),
            "get_info" => GetPageInfo(sessionId, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds one or more pages to the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="count">Number of pages to add.</param>
    /// <param name="insertAt">Optional 1-based position to insert pages at.</param>
    /// <param name="width">Optional page width in points.</param>
    /// <param name="height">Optional page height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private string AddPage(string? sessionId, string? path, string? outputPath, int count, int? insertAt, double? width,
        double? height)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (insertAt is >= 1 && insertAt.Value <= document.Pages.Count)
            for (var i = 0; i < count; i++)
            {
                var page = document.Pages.Insert(insertAt.Value + i);
                if (width.HasValue && height.HasValue)
                    page.SetPageSize(width.Value, height.Value);
                else
                    page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            }
        else
            for (var i = 0; i < count; i++)
            {
                var page = document.Pages.Add();
                if (width.HasValue && height.HasValue)
                    page.SetPageSize(width.Value, height.Value);
                else
                    page.SetPageSize(PageSize.A4.Width, PageSize.A4.Height);
            }

        ctx.Save(outputPath);

        return $"Added {count} page(s). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a page from the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is invalid.</exception>
    private string DeletePage(string? sessionId, string? path, string? outputPath, int pageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        document.Pages.Delete(pageIndex);

        ctx.Save(outputPath);

        return $"Deleted page {pageIndex} (remaining: {document.Pages.Count}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Rotates one or more pages in the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="rotation">Rotation angle in degrees (0, 90, 180, 270).</param>
    /// <param name="pageIndex">Optional 1-based page index to rotate a single page.</param>
    /// <param name="pageIndices">Optional array of 1-based page indices to rotate multiple pages.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the rotation value is invalid.</exception>
    private string RotatePage(string? sessionId, string? path, string? outputPath, int rotation, int? pageIndex,
        int[]? pageIndices)
    {
        if (rotation != 0 && rotation != 90 && rotation != 180 && rotation != 270)
            throw new ArgumentException("rotation must be 0, 90, 180, or 270");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        var rotationEnum = rotation switch
        {
            90 => Rotation.on90,
            180 => Rotation.on180,
            270 => Rotation.on270,
            _ => Rotation.None
        };

        List<int> pagesToRotate;
        if (pageIndices is { Length: > 0 })
            pagesToRotate = pageIndices.ToList();
        else if (pageIndex is > 0)
            pagesToRotate = [pageIndex.Value];
        else
            pagesToRotate = Enumerable.Range(1, document.Pages.Count).ToList();

        foreach (var pageNum in pagesToRotate)
            if (pageNum >= 1 && pageNum <= document.Pages.Count)
                document.Pages[pageNum].Rotate = rotationEnum;

        ctx.Save(outputPath);

        return $"Rotated {pagesToRotate.Count} page(s) by {rotation} degrees. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves detailed information about a specific page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <returns>A JSON string containing detailed page information.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is invalid.</exception>
    private string GetPageDetails(string? sessionId, string? path, int pageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

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
    }

    /// <summary>
    ///     Retrieves basic information about all pages in the PDF document.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <returns>A JSON string containing page information for all pages.</returns>
    private string GetPageInfo(string? sessionId, string? path)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
        List<object> pageList = [];

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
    }
}