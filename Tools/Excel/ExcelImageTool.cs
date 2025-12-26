using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel images (add, delete, get)
///     Merges: ExcelAddImageTool, ExcelDeleteImageTool, ExcelGetImagesTool
/// </summary>
public class ExcelImageTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel images. Supports 3 operations: add, delete, get.

Usage examples:
- Add image: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, height=150)
- Delete image: excel_image(operation='delete', path='book.xlsx', imageIndex=0)
- Get images: excel_image(operation='get', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
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
- 'get': Get all images (required params: path)",
                @enum = new[] { "add", "delete", "get" }
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
                description = "Path to the image file (required for add)"
            },
            cell = new
            {
                type = "string",
                description = "Top-left cell reference (e.g., 'A1', required for add)"
            },
            width = new
            {
                type = "number",
                description = "Image width in pixels (optional)"
            },
            height = new
            {
                type = "number",
                description = "Image height in pixels (optional)"
            },
            imageIndex = new
            {
                type = "number",
                description = "Image index (0-based, required for delete)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteImageAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetImagesAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an image to the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing imagePath, cell, optional width, height</param>
    /// <returns>Success message</returns>
    private Task<string> AddImageAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var imagePath = ArgumentHelper.GetString(arguments, "imagePath");
            SecurityHelper.ValidateFilePath(imagePath, "imagePath", true);
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var width = ArgumentHelper.GetIntNullable(arguments, "width");
            var height = ArgumentHelper.GetIntNullable(arguments, "height");

            if (!File.Exists(imagePath)) throw new FileNotFoundException($"Image file not found: {imagePath}");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
            var cellObj = worksheet.Cells[cell];

            var pictureIndex = worksheet.Pictures.Add(cellObj.Row, cellObj.Column, imagePath);

            if (width.HasValue || height.HasValue)
            {
                var picture = worksheet.Pictures[pictureIndex];
                if (width.HasValue) picture.Width = width.Value;
                if (height.HasValue) picture.Height = height.Value;
            }

            workbook.Save(outputPath);

            return $"Image added to cell {cell}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes an image from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing imageIndex</param>
    /// <returns>Success message</returns>
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
                    $"Image index {imageIndex} is out of range (worksheet has {pictures.Count} images)");

            pictures.RemoveAt(imageIndex);
            workbook.Save(outputPath);

            return $"Image #{imageIndex} deleted, {pictures.Count} remaining. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all images from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with all images</returns>
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
                imageList.Add(new
                {
                    index = i,
                    location = new
                    {
                        upperLeftRow = picture.UpperLeftRow,
                        lowerRightRow = picture.LowerRightRow,
                        upperLeftColumn = picture.UpperLeftColumn,
                        lowerRightColumn = picture.LowerRightColumn
                    },
                    width = picture.Width,
                    height = picture.Height
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
}