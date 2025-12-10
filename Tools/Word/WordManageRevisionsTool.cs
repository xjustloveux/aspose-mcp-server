using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordManageRevisionsTool : IAsposeTool
{
    public string Description => "Accept or reject all tracked revisions in a Word document";

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
            action = new
            {
                type = "string",
                description = "Action to take: accept | reject (default: accept)"
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
        var action = arguments?["action"]?.GetValue<string>()?.ToLowerInvariant() ?? "accept";
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;

        var doc = new Document(path);
        var revisionsCount = doc.Revisions.Count;

        if (revisionsCount == 0)
        {
            if (!string.Equals(path, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                doc.Save(outputPath);
                return await Task.FromResult($"文檔沒有變更記錄，已另存到: {outputPath}");
            }

            return await Task.FromResult("文檔沒有變更記錄");
        }

        switch (action)
        {
            case "accept":
                doc.AcceptAllRevisions();
                break;
            case "reject":
                doc.Revisions.RejectAll();
                break;
            default:
                throw new ArgumentException("action must be 'accept' or 'reject'");
        }

        doc.Save(outputPath);
        return await Task.FromResult($"處理變更完成\n原始變更數: {revisionsCount}\n操作: {action}\n輸出: {outputPath}");
    }
}

