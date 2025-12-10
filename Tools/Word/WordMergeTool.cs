using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordMergeTool : IAsposeTool
{
    public string Description => "Merge multiple Word documents into one";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPaths = new
            {
                type = "array",
                description = "Array of input file paths to merge",
                items = new { type = "string" }
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path"
            }
        },
        required = new[] { "inputPaths", "outputPath" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => p != null).ToList();
        
        if (inputPaths.Count == 0)
            throw new ArgumentException("At least one input path is required");

        var mergedDoc = new Document(inputPaths[0]);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            var doc = new Document(inputPaths[i]);
            mergedDoc.AppendDocument(doc, ImportFormatMode.KeepSourceFormatting);
        }

        mergedDoc.Save(outputPath);

        return await Task.FromResult($"Merged {inputPaths.Count} documents into: {outputPath}");
    }
}

