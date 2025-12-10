using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetNamedRangesTool : IAsposeTool
{
    public string Description => "Get all named ranges from an Excel workbook";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        var result = new StringBuilder();

        result.AppendLine("=== Excel 工作簿的名稱範圍資訊 ===\n");
        result.AppendLine($"總名稱範圍數: {names.Count}\n");

        if (names.Count == 0)
        {
            result.AppendLine("未找到名稱範圍");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < names.Count; i++)
        {
            var name = names[i];
            result.AppendLine($"【名稱範圍 {i}】");
            result.AppendLine($"名稱: {name.Text}");
            result.AppendLine($"引用: {name.RefersTo}");
            result.AppendLine($"註解: {name.Comment ?? "(無)"}");
            result.AppendLine($"是否可見: {name.IsVisible}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

