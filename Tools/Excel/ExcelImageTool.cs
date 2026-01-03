using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel images (add, delete, get, extract).
/// </summary>
[McpServerToolType]
public class ExcelImageTool
{
    /// <summary>
    ///     Set of supported image file extensions.
    /// </summary>
    private static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".emf", ".wmf"
    };

    /// <summary>
    ///     Mapping of file extensions to Aspose.Cells ImageType enum values.
    /// </summary>
    private static readonly Dictionary<string, ImageType> ExtensionToImageType = new(StringComparer.OrdinalIgnoreCase)
    {
        { ".png", ImageType.Png },
        { ".jpg", ImageType.Jpeg },
        { ".jpeg", ImageType.Jpeg },
        { ".gif", ImageType.Gif },
        { ".bmp", ImageType.Bmp },
        { ".tiff", ImageType.Tiff },
        { ".tif", ImageType.Tiff },
        { ".emf", ImageType.Emf },
        { ".wmf", ImageType.Wmf }
    };

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelImageTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelImageTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_image")]
    [Description(@"Manage Excel images. Supports 4 operations: add, delete, get, extract.

Usage examples:
- Add image: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, height=150)
- Add image with aspect ratio: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, keepAspectRatio=true)
- Delete image: excel_image(operation='delete', path='book.xlsx', imageIndex=0)
- Get images: excel_image(operation='get', path='book.xlsx')
- Extract image: excel_image(operation='extract', path='book.xlsx', imageIndex=0, exportPath='extracted.png')

Note: When deleting images, the indices of remaining images will be re-ordered.")]
    public string Execute(
        [Description("Operation: add, delete, get, extract")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Path to the image file. Supported formats: png, jpg, jpeg, gif, bmp, tiff, emf, wmf (required for add)")]
        string? imagePath = null,
        [Description("Top-left cell reference (e.g., 'A1', required for add)")]
        string? cell = null,
        [Description("Image width in pixels (optional for add)")]
        int? width = null,
        [Description("Image height in pixels (optional for add)")]
        int? height = null,
        [Description(
            "Keep aspect ratio when resizing. If true and only width or height is specified, the other dimension is calculated proportionally (optional for add, default: true)")]
        bool keepAspectRatio = true,
        [Description("Image index (0-based, required for delete/extract). Note: indices are re-ordered after deletion")]
        int? imageIndex = null,
        [Description(
            "Path to export the extracted image (required for extract). Format determined by file extension (png, jpg, gif, bmp, tiff)")]
        string? exportPath = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddImage(ctx, outputPath, sheetIndex, imagePath, cell, width, height, keepAspectRatio),
            "delete" => DeleteImage(ctx, outputPath, sheetIndex, imageIndex),
            "get" => GetImages(ctx, sheetIndex),
            "extract" => ExtractImage(ctx, sheetIndex, imageIndex, exportPath),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="imagePath">The path to the image file.</param>
    /// <param name="cell">The cell address for the image location.</param>
    /// <param name="width">The image width in pixels.</param>
    /// <param name="height">The image height in pixels.</param>
    /// <param name="keepAspectRatio">Whether to maintain the aspect ratio when resizing.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when imagePath or cell is not provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the image file does not exist.</exception>
    private static string AddImage(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? imagePath,
        string? cell, int? width, int? height, bool keepAspectRatio)
    {
        if (string.IsNullOrEmpty(imagePath))
            throw new ArgumentException("imagePath is required for add operation");
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");

        SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);

        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"Image file not found: {imagePath}");

        ValidateImageFormat(imagePath);

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        var pictureIndex = worksheet.Pictures.Add(cellObj.Row, cellObj.Column, imagePath);
        var picture = worksheet.Pictures[pictureIndex];

        if (width.HasValue || height.HasValue)
        {
            picture.IsLockAspectRatio = keepAspectRatio;
            if (width.HasValue) picture.Width = width.Value;
            if (height.HasValue) picture.Height = height.Value;
        }

        ctx.Save(outputPath);
        return
            $"Image added to cell {cell} (size: {picture.Width}x{picture.Height}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes an image from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="imageIndex">The index of the image to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when imageIndex is not provided or is out of range.</exception>
    private static string DeleteImage(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? imageIndex)
    {
        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for delete operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;

        if (imageIndex.Value < 0 || imageIndex.Value >= pictures.Count)
            throw new ArgumentException(
                $"Image index {imageIndex.Value} is out of range. Worksheet has {pictures.Count} images (valid indices: 0-{pictures.Count - 1}).");

        pictures.RemoveAt(imageIndex.Value);

        ctx.Save(outputPath);

        var warning = pictures.Count > 0
            ? " Note: remaining image indices have been re-ordered."
            : "";
        return
            $"Image #{imageIndex.Value} deleted. {pictures.Count} images remaining.{warning} {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all images from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the image information.</returns>
    private static string GetImages(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;

        if (pictures.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No images found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> imageList = [];
        for (var i = 0; i < pictures.Count; i++)
        {
            var picture = pictures[i];
            var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);
            var lowerRightCell = CellsHelper.CellIndexToName(picture.LowerRightRow, picture.LowerRightColumn);

            imageList.Add(new
            {
                index = i,
                name = picture.Name,
                alternativeText = picture.AlternativeText,
                imageType = picture.ImageType.ToString(),
                location = new
                {
                    upperLeftCell,
                    lowerRightCell,
                    upperLeftRow = picture.UpperLeftRow,
                    upperLeftColumn = picture.UpperLeftColumn,
                    lowerRightRow = picture.LowerRightRow,
                    lowerRightColumn = picture.LowerRightColumn
                },
                width = picture.Width,
                height = picture.Height,
                isLockAspectRatio = picture.IsLockAspectRatio
            });
        }

        var result = new
        {
            count = pictures.Count,
            worksheetName = worksheet.Name,
            items = imageList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Extracts an image from the worksheet and saves it to a file.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="imageIndex">The index of the image to extract.</param>
    /// <param name="exportPath">The path to save the extracted image.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when imageIndex or exportPath is not provided, the export format is
    ///     unsupported, or the image index is out of range.
    /// </exception>
    private static string ExtractImage(DocumentContext<Workbook> ctx, int sheetIndex, int? imageIndex,
        string? exportPath)
    {
        if (!imageIndex.HasValue)
            throw new ArgumentException("imageIndex is required for extract operation");
        if (string.IsNullOrEmpty(exportPath))
            throw new ArgumentException("exportPath is required for extract operation");

        SecurityHelper.ValidateFilePath(exportPath, "exportPath", true);

        var extension = Path.GetExtension(exportPath);
        if (string.IsNullOrEmpty(extension) || !ExtensionToImageType.TryGetValue(extension, out var imageType))
            throw new ArgumentException(
                $"Unsupported export format: '{extension}'. Supported formats: {string.Join(", ", ExtensionToImageType.Keys)}");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;

        if (imageIndex.Value < 0 || imageIndex.Value >= pictures.Count)
            throw new ArgumentException(
                $"Image index {imageIndex.Value} is out of range. Worksheet has {pictures.Count} images (valid indices: 0-{pictures.Count - 1}).");

        var picture = pictures[imageIndex.Value];
        var upperLeftCell = CellsHelper.CellIndexToName(picture.UpperLeftRow, picture.UpperLeftColumn);

        var exportDir = Path.GetDirectoryName(exportPath);
        if (!string.IsNullOrEmpty(exportDir) && !Directory.Exists(exportDir))
            Directory.CreateDirectory(exportDir);

        var options = new ImageOrPrintOptions
        {
            ImageType = imageType
        };
        picture.ToImage(exportPath, options);

        var fileInfo = new FileInfo(exportPath);
        return
            $"Image #{imageIndex.Value} (at {upperLeftCell}) extracted to: {exportPath} ({fileInfo.Length} bytes, {picture.Width}x{picture.Height})";
    }

    /// <summary>
    ///     Validates that the image file has a supported format.
    /// </summary>
    /// <param name="imagePath">The path to the image file to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the image format is not supported.</exception>
    private static void ValidateImageFormat(string imagePath)
    {
        var extension = Path.GetExtension(imagePath);
        if (string.IsNullOrEmpty(extension) || !SupportedImageExtensions.Contains(extension))
            throw new ArgumentException(
                $"Unsupported image format: '{extension}'. Supported formats: {string.Join(", ", SupportedImageExtensions)}");
    }
}