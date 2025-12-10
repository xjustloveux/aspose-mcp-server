using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordUnprotectTool : IAsposeTool
{
    public string Description => "Remove document protection from a Word file";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            password = new
            {
                type = "string",
                description = "Password used for protection (if any)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var password = arguments?["password"]?.GetValue<string>();
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;

        var doc = new Document(path);
        var wasProtected = doc.ProtectionType != ProtectionType.NoProtection;

        if (!wasProtected)
        {
            if (!string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                doc.Save(outputPath);
                return await Task.FromResult($"文檔未受保護，已另存到: {outputPath}");
            }

            return await Task.FromResult("文檔未受保護，無需解除");
        }

        doc.Unprotect(password);

        if (doc.ProtectionType != ProtectionType.NoProtection)
        {
            throw new InvalidOperationException("解除保護失敗，可能是密碼錯誤或文檔被限制");
        }

        doc.Save(outputPath);
        return await Task.FromResult($"解除保護完成\n輸出: {outputPath}");
    }
}

