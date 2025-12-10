using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Comparing;

namespace AsposeMcpServer.Tools;

public class WordCompareDocumentsTool : IAsposeTool
{
    public string Description => "Compare two Word documents and generate comparison document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            originalPath = new
            {
                type = "string",
                description = "Original document file path"
            },
            revisedPath = new
            {
                type = "string",
                description = "Revised document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output comparison document path"
            },
            authorName = new
            {
                type = "string",
                description = "Author name for revisions (optional, default: 'Comparison')"
            }
        },
        required = new[] { "originalPath", "revisedPath", "outputPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var originalPath = arguments?["originalPath"]?.GetValue<string>() ?? throw new ArgumentException("originalPath is required");
        var revisedPath = arguments?["revisedPath"]?.GetValue<string>() ?? throw new ArgumentException("revisedPath is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var authorName = arguments?["authorName"]?.GetValue<string>() ?? "Comparison";

        var originalDoc = new Document(originalPath);
        var revisedDoc = new Document(revisedPath);

        originalDoc.Compare(revisedDoc, authorName, DateTime.Now);

        originalDoc.Save(outputPath);
        return await Task.FromResult($"Comparison document created: {outputPath}");
    }
}

