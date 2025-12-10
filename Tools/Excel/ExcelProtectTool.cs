using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelProtectTool : IAsposeTool
{
    public string Description => "Protect an Excel workbook or worksheet with password";

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
                description = "Protection password"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index to protect (optional, protects workbook if not specified)"
            }
        },
        required = new[] { "path", "password" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>();

        using var workbook = new Workbook(path);

        if (sheetIndex.HasValue)
        {
            workbook.Worksheets[sheetIndex.Value].Protect(ProtectionType.All, password, null);
        }
        else
        {
            workbook.Protect(ProtectionType.All, password);
        }

        workbook.Save(path);

        return await Task.FromResult($"Excel protected with password: {path}");
    }
}

