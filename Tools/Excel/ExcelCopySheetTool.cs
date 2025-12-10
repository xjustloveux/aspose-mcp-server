using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCopySheetTool : IAsposeTool
{
    public string Description => "Copy a worksheet within the same workbook or to another workbook";

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
            sourceSheetIndex = new
            {
                type = "number",
                description = "Source sheet index to copy (0-based)"
            },
            newSheetName = new
            {
                type = "string",
                description = "Name for the copied sheet"
            },
            targetPath = new
            {
                type = "string",
                description = "Target workbook path (optional, if not provided, copies within same workbook)"
            }
        },
        required = new[] { "path", "sourceSheetIndex", "newSheetName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sourceSheetIndex = arguments?["sourceSheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceSheetIndex is required");
        var newSheetName = arguments?["newSheetName"]?.GetValue<string>() ?? throw new ArgumentException("newSheetName is required");
        var targetPath = arguments?["targetPath"]?.GetValue<string>();

        using var sourceWorkbook = new Workbook(path);

        if (sourceSheetIndex < 0 || sourceSheetIndex >= sourceWorkbook.Worksheets.Count)
        {
            throw new ArgumentException($"源工作表索引 {sourceSheetIndex} 超出範圍 (共有 {sourceWorkbook.Worksheets.Count} 個工作表)");
        }

        if (string.IsNullOrEmpty(targetPath))
        {
            // Copy within same workbook
            var newSheetIndex = sourceWorkbook.Worksheets.AddCopy(sourceSheetIndex);
            var newSheet = sourceWorkbook.Worksheets[newSheetIndex];
            newSheet.Name = newSheetName;
            sourceWorkbook.Save(path);
            return await Task.FromResult($"工作表已複製到同一工作簿，新名稱: '{newSheetName}': {path}");
        }
        else
        {
            // Copy to another workbook
            using var targetWorkbook = new Workbook(targetPath);
            var sourceSheet = sourceWorkbook.Worksheets[sourceSheetIndex];
            var newSheet = targetWorkbook.Worksheets.Add(sourceSheet.Name);
            newSheet.Copy(sourceSheet);
            newSheet.Name = newSheetName;
            targetWorkbook.Save(targetPath);
            return await Task.FromResult($"工作表已複製到 '{targetPath}'，新名稱: '{newSheetName}'");
        }
    }
}
