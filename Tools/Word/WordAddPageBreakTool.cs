using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddPageBreakTool : IAsposeTool
{
    public string Description => "Add a page break to a Word document";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        
        // Move to end of document
        builder.MoveToDocumentEnd();
        
        // Insert page break
        builder.InsertBreak(BreakType.PageBreak);

        doc.Save(outputPath);

        return await Task.FromResult($"成功添加分頁符號\n輸出: {outputPath}");
    }
}

