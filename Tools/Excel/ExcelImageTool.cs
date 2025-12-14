using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel images (add, delete, get)
/// Merges: ExcelAddImageTool, ExcelDeleteImageTool, ExcelGetImagesTool
/// </summary>
public class ExcelImageTool : IAsposeTool
{
    public string Description => @"Manage Excel images. Supports 3 operations: add, delete, get.

Usage examples:
- Add image: excel_image(operation='add', path='book.xlsx', imagePath='image.png', cell='A1', width=200, height=150)
- Delete image: excel_image(operation='delete', path='book.xlsx', imageIndex=0)
- Get images: excel_image(operation='get', path='book.xlsx')";

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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddImageAsync(arguments, path, sheetIndex),
            "delete" => await DeleteImageAsync(arguments, path, sheetIndex),
            "get" => await GetImagesAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds an image to the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing imagePath, cell, optional width, height</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddImageAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var imagePath = ArgumentHelper.GetString(arguments, "imagePath", "imagePath");
        SecurityHelper.ValidateFilePath(imagePath, "imagePath");
        var cell = ArgumentHelper.GetString(arguments, "cell", "cell");
        var width = arguments?["width"]?.GetValue<int?>();
        var height = arguments?["height"]?.GetValue<int?>();

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"圖片檔案不存在: {imagePath}");
        }

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

        workbook.Save(path);

        return await Task.FromResult($"圖片已添加到單元格 {cell}: {path}");
    }

    /// <summary>
    /// Deletes an image from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing imageIndex</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteImageAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var imageIndex = ArgumentHelper.GetInt(arguments, "imageIndex", "imageIndex");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;
        
        if (imageIndex < 0 || imageIndex >= pictures.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (工作表共有 {pictures.Count} 個圖片)");
        }

        pictures.RemoveAt(imageIndex);
        workbook.Save(path);
        
        var remainingCount = pictures.Count;
        
        return await Task.FromResult($"成功刪除圖片 #{imageIndex}\n工作表剩餘圖片數: {remainingCount}\n輸出: {path}");
    }

    /// <summary>
    /// Gets all images from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all images</returns>
    private async Task<string> GetImagesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pictures = worksheet.Pictures;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的圖片資訊 ===\n");
        result.AppendLine($"總圖片數: {pictures.Count}\n");

        if (pictures.Count == 0)
        {
            result.AppendLine("未找到圖片");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < pictures.Count; i++)
        {
            var picture = pictures[i];
            result.AppendLine($"【圖片 {i}】");
            result.AppendLine($"位置: 行 {picture.UpperLeftRow}-{picture.LowerRightRow}, 列 {picture.UpperLeftColumn}-{picture.LowerRightColumn}");
            result.AppendLine($"寬度: {picture.Width} 像素");
            result.AppendLine($"高度: {picture.Height} 像素");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

