using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelAddNamedRangeTool : IAsposeTool
{
    public string Description => "Add a named range to an Excel workbook";

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
                description = "Name for the range"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10') or formula"
            },
            comment = new
            {
                type = "string",
                description = "Comment for the named range (optional)"
            }
        },
        required = new[] { "path", "name", "range" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var name = arguments?["name"]?.GetValue<string>() ?? throw new ArgumentException("name is required");
        var range = arguments?["range"]?.GetValue<string>() ?? throw new ArgumentException("range is required");
        var comment = arguments?["comment"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        
        // Check if name already exists
        try
        {
            var existingName = names[name];
            if (existingName != null)
            {
                throw new ArgumentException($"名稱範圍 '{name}' 已存在");
            }
        }
        catch
        {
            // Name doesn't exist, continue
        }
        
        var nameIndex = names.Add(range);
        var namedRange = names[nameIndex];
        namedRange.Text = name;
        if (!string.IsNullOrEmpty(comment))
        {
            namedRange.Comment = comment;
        }
        
        workbook.Save(path);
        
        return await Task.FromResult($"成功添加名稱範圍 '{name}'\n引用: {range}\n輸出: {path}");
    }
}
