using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel images (add, delete, get, extract).
/// </summary>
public class ExcelImageTool : IAsposeTool
{
    private static readonly HashSet<string> SupportedImageExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".emf", ".wmf"
    };

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
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel images. Supports 4 operations: add, delete, get, extract.

Usage examples:
- Add image: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, height=150)
- Add image with aspect ratio: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, keepAspectRatio=true)
- Delete image: excel_image(operation='delete', path='book.xlsx', imageIndex=0)
- Get images: excel_image(operation='get', path='book.xlsx')
- Extract image: excel_image(operation='extract', path='book.xlsx', imageIndex=0, exportPath='extracted.png')

Note: When deleting images, the indices of remaining images will be re-ordered.";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add an image (required params: path, imagePath, cell)
- 'delete': Delete an image (required params: path, imageIndex)
- 'get': Get all images (required params: path)
- 'extract': Extract image to file (required params: path, imageIndex, exportPath)",
                @enum = new[] { "add", "delete", "get", "extract" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            imagePath = new
            {
                type = "string",
                description =
                    "Path to the image file. Supported formats: png, jpg, jpeg, gif, bmp, tiff, emf, wmf (required for add)"
            },
            cell = new
            {
                type = "string",
                description = "Top-left cell reference (e.g., 'A1', required for add)"
            },
            width = new
            {
                type = "number",
                description = "Image width in pixels (optional for add)"
            },
            height = new
            {
                type = "number",
                description = "Image height in pixels (optional for add)"
            },
            keepAspectRatio = new
            {
                type = "boolean",
                description =
                    "Keep aspect ratio when resizing. If true and only width or height is specified, the other dimension is calculated proportionally (optional for add, default: true)"
            },
            imageIndex = new
            {
                type = "number",
                description =
                    "Image index (0-based, required for delete/extract). Note: indices are re-ordered after deletion"
            },
            outputPath = new
            {
                type = "string",
                description = "Output Excel file path (optional for add/delete, defaults to input path)"
            },
            exportPath = new
            {
                type = "string",
                description =
                    "Path to export the extracted image (required for extract). Format determined by file extension (png, jpg, gif, bmp, tiff)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteImageAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetImagesAsync(path, sheetIndex),
            "extract" => await ExtractImageAsync(path, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing imagePath, cell, optional width, height, keepAspectRatio.</param>
    /// <returns>Success message with image details.</returns>
    /// <exception cref="FileNotFoundException">Thrown when image file is not found.</exception>
    /// <exception cref="ArgumentException">Thrown when image format is not supported.</exception>
    private Task<string> AddImageAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var width = ArgumentHelper.GetIntNullable(arguments, "width");
            var height = ArgumentHelper.GetIntNullable(arguments, "height");
            var keepAspectRatio = ArgumentHelper.GetBool(arguments, "keepAspectRatio", true);

            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Image file not found: {imagePath}");

            ValidateImageFormat(imagePath);

            using var workbook = new Workbook(path);
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

            workbook.Save(outputPath);

            return $"Image added to cell {cell} (size: {picture.Width}x{picture.Height}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an image from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing imageIndex.</param>
    /// <returns>Success message with remaining image count.</returns>
    /// <exception cref="ArgumentException">Thrown when image index is out of range.</exception>
    private Task<string> DeleteImageAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pictures = worksheet.Pictures;

            if (imageIndex < 0 || imageIndex >= pictures.Count)
                throw new ArgumentException(
                    $"Image index {imageIndex} is out of range. Worksheet has {pictures.Count} images (valid indices: 0-{pictures.Count - 1}).");

            pictures.RemoveAt(imageIndex);
            workbook.Save(outputPath);

            var warning = pictures.Count > 0
                ? " Note: remaining image indices have been re-ordered."
                : "";
            return $"Image #{imageIndex} deleted. {pictures.Count} images remaining.{warning} Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all images from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string containing all image details.</returns>
    private Task<string> GetImagesAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
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

            var imageList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Extracts an image from the worksheet and saves it to a file.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing imageIndex and exportPath.</param>
    /// <returns>Success message with export details.</returns>
    /// <exception cref="ArgumentException">Thrown when image index is out of range or export format is unsupported.</exception>
    private Task<string> ExtractImageAsync(string path, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex");
            var exportPath = ArgumentHelper.GetString(arguments, "exportPath");
            SecurityHelper.ValidateFilePath(exportPath, "exportPath", true);

            var extension = Path.GetExtension(exportPath);
            if (string.IsNullOrEmpty(extension) || !ExtensionToImageType.TryGetValue(extension, out var imageType))
                throw new ArgumentException(
                    $"Unsupported export format: '{extension}'. Supported formats: {string.Join(", ", ExtensionToImageType.Keys)}");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pictures = worksheet.Pictures;

            if (imageIndex < 0 || imageIndex >= pictures.Count)
                throw new ArgumentException(
                    $"Image index {imageIndex} is out of range. Worksheet has {pictures.Count} images (valid indices: 0-{pictures.Count - 1}).");

            var picture = pictures[imageIndex];
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
                $"Image #{imageIndex} (at {upperLeftCell}) extracted to: {exportPath} ({fileInfo.Length} bytes, {picture.Width}x{picture.Height})";
        });
    }

    /// <summary>
    ///     Validates that the image file has a supported format.
    /// </summary>
    /// <param name="imagePath">Path to the image file.</param>
    /// <exception cref="ArgumentException">Thrown when image format is not supported.</exception>
    private static void ValidateImageFormat(string imagePath)
    {
        var extension = Path.GetExtension(imagePath);
        if (string.IsNullOrEmpty(extension) || !SupportedImageExtensions.Contains(extension))
            throw new ArgumentException(
                $"Unsupported image format: '{extension}'. Supported formats: {string.Join(", ", SupportedImageExtensions)}");
    }
}