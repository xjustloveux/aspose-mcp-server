using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelUnprotectTool : IAsposeTool
{
    public string Description => "Remove protection from an Excel workbook or worksheet";

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
            password = new
            {
                type = "string",
                description = "Password used for protection (if any)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index to unprotect (0-based, optional; unprotects workbook when omitted)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>();
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);

        if (sheetIndex.HasValue)
        {
            if (sheetIndex.Value < 0 || sheetIndex.Value >= workbook.Worksheets.Count)
            {
                throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
            }

            var worksheet = workbook.Worksheets[sheetIndex.Value];
            var wasProtected = worksheet.IsProtected;
            worksheet.Unprotect(password);

            workbook.Save(path);
            return await Task.FromResult($"工作表解除保護完成: {worksheet.Name}\n原狀態: {(wasProtected ? "已保護" : "未保護")}\n輸出: {path}");
        }
        else
        {
            workbook.Unprotect(password);
            workbook.Save(path);
            return await Task.FromResult($"工作簿保護已解除\n輸出: {path}");
        }
    }
}

