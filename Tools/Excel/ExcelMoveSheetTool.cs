using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelMoveSheetTool : IAsposeTool
{
    public string Description => "Move a worksheet to a different position in the workbook";

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
                description = "Sheet index to move (0-based)"
            },
            targetIndex = new
            {
                type = "number",
                description = "Target position index (0-based)"
            }
        },
        required = new[] { "path", "sheetIndex", "targetIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? throw new ArgumentException("sheetIndex is required");
        var targetIndex = arguments?["targetIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetIndex is required");

        using var workbook = new Workbook(path);

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        if (targetIndex < 0 || targetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"目標位置索引 {targetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        var sheet = workbook.Worksheets[sheetIndex];
        
        // Create a copy at target position, then remove original
        var tempSheetIndex = workbook.Worksheets.AddCopy(sheetIndex);
        var tempSheet = workbook.Worksheets[tempSheetIndex];
        tempSheet.Name = sheetName + "_temp";
        
        workbook.Worksheets.Insert(targetIndex, SheetType.Worksheet);
        var movedSheet = workbook.Worksheets[targetIndex];
        movedSheet.Copy(tempSheet);
        movedSheet.Name = sheetName;
        
        // Remove original and temp
        var originalIndex = sheetIndex < targetIndex ? sheetIndex : sheetIndex + 1;
        workbook.Worksheets.RemoveAt(originalIndex);
        var tempIndex = tempSheetIndex < targetIndex ? tempSheetIndex : tempSheetIndex + 1;
        workbook.Worksheets.RemoveAt(tempIndex);
        
        workbook.Save(path);

        return await Task.FromResult($"工作表 '{sheetName}' 已從位置 {sheetIndex} 移動到位置 {targetIndex}: {path}");
    }
}

