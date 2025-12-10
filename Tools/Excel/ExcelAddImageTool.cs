using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeMcpServer.Tools;

public class ExcelAddImageTool : IAsposeTool
{
    public string Description => "Add an image to an Excel worksheet";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            imagePath = new
            {
                type = "string",
                description = "Path to the image file"
            },
            cell = new
            {
                type = "string",
                description = "Top-left cell reference (e.g., 'A1')"
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
            }
        },
        required = new[] { "path", "imagePath", "cell" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var imagePath = arguments?["imagePath"]?.GetValue<string>() ?? throw new ArgumentException("imagePath is required");
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var width = arguments?["width"]?.GetValue<int?>();
        var height = arguments?["height"]?.GetValue<int?>();

        if (!File.Exists(imagePath))
        {
            throw new FileNotFoundException($"圖片文件不存在: {imagePath}");
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
}

