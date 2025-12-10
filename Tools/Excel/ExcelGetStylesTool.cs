using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetStylesTool : IAsposeTool
{
    public string Description => "Get all styles from Excel workbook";

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
        var sb = new StringBuilder();

        sb.AppendLine("Note: Aspose.Cells doesn't support named styles directly.");
        sb.AppendLine("Styles are applied to cells/ranges, not stored as named styles in workbook.");
        sb.AppendLine("To get style information, use excel_get_cell_format on specific cells.");

        return await Task.FromResult(sb.ToString());
    }
}

