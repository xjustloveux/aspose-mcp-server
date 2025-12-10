using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteNamedRangeTool : IAsposeTool
{
    public string Description => "Delete a named range from an Excel workbook";

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
            name = new
            {
                type = "string",
                description = "Name of the range to delete"
            }
        },
        required = new[] { "path", "name" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");

        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        
        Name? namedRange = null;
        try
        {
            namedRange = names[name];
        }
        catch
        {
            throw new ArgumentException($"名稱範圍 '{name}' 不存在");
        }
        
        if (namedRange == null)
        {
            throw new ArgumentException($"名稱範圍 '{name}' 不存在");
        }
        
        var refersTo = namedRange.RefersTo;
        // Find the index of the named range
        int indexToRemove = -1;
        for (int i = 0; i < names.Count; i++)
        {
            if (names[i] == namedRange)
            {
                indexToRemove = i;
                break;
            }
        }
        
        if (indexToRemove >= 0)
        {
            names.RemoveAt(indexToRemove);
        }
        workbook.Save(path);
        
        var remainingCount = names.Count;
        
        return await Task.FromResult($"成功刪除名稱範圍 '{name}'\n原引用: {refersTo}\n工作簿剩餘名稱範圍數: {remainingCount}\n輸出: {path}");
    }
}

