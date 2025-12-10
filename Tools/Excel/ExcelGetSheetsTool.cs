using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetSheetsTool : IAsposeTool
{
    public string Description => "Get a simple list of all worksheets in an Excel workbook";

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
        var result = new StringBuilder();

        result.AppendLine($"=== 工作簿 '{Path.GetFileName(path)}' 的工作表列表 ===\n");
        result.AppendLine($"總工作表數: {workbook.Worksheets.Count}\n");

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            var worksheet = workbook.Worksheets[i];
            result.AppendLine($"{i}. {worksheet.Name} (可見性: {worksheet.VisibilityType})");
        }

        return await Task.FromResult(result.ToString());
    }
}

