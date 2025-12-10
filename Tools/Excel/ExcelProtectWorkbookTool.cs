using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelProtectWorkbookTool : IAsposeTool
{
    public string Description => "Protect Excel workbook structure and windows";

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
            protectStructure = new
            {
                type = "boolean",
                description = "Protect workbook structure (default: true)"
            },
            protectWindows = new
            {
                type = "boolean",
                description = "Protect workbook windows (default: false)"
            }
        },
        required = new[] { "path", "password" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>() ?? throw new ArgumentException("password is required");
        var protectStructure = arguments?["protectStructure"]?.GetValue<bool>() ?? true;
        var protectWindows = arguments?["protectWindows"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        
        var protectionType = ProtectionType.None;
        if (protectStructure && protectWindows)
        {
            protectionType = ProtectionType.All;
        }
        else if (protectStructure)
        {
            protectionType = ProtectionType.Structure;
        }
        else if (protectWindows)
        {
            protectionType = ProtectionType.Windows;
        }
        
        workbook.Protect(protectionType, password);
        workbook.Save(path);
        
        var result = $"工作簿保護已設定\n";
        result += $"保護結構: {protectStructure}\n";
        result += $"保護視窗: {protectWindows}\n";
        result += $"輸出: {path}";
        
        return await Task.FromResult(result);
    }
}

