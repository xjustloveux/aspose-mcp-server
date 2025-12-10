using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelSetSheetBackgroundTool : IAsposeTool
{
    public string Description => "Set worksheet background image in Excel";

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
                description = "Background image file path"
            },
            removeBackground = new
            {
                type = "boolean",
                description = "Remove background image (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>() ?? 0;
        var imagePath = arguments?["imagePath"]?.GetValue<string>();
        var removeBackground = arguments?["removeBackground"]?.GetValue<bool?>() ?? false;

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];

        if (removeBackground)
        {
            worksheet.BackgroundImage = null;
        }
        else if (!string.IsNullOrEmpty(imagePath))
        {
            if (!File.Exists(imagePath))
            {
                throw new FileNotFoundException($"Image file not found: {imagePath}");
            }
            var imageBytes = File.ReadAllBytes(imagePath);
            worksheet.BackgroundImage = imageBytes;
        }
        else
        {
            throw new ArgumentException("Either imagePath or removeBackground must be provided");
        }

        workbook.Save(path);
        return await Task.FromResult(removeBackground 
            ? $"Background image removed from sheet {sheetIndex}: {path}"
            : $"Background image set for sheet {sheetIndex}: {path}");
    }
}

