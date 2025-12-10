using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfMergeTool : IAsposeTool
{
    public string Description => "Merge multiple PDF documents into one";

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

        using var mergedDocument = new Document(inputPaths[0]);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            using var doc = new Document(inputPaths[i]);
            mergedDocument.Pages.Add(doc.Pages);
        }

        mergedDocument.Save(outputPath);

        return await Task.FromResult($"Merged {inputPaths.Count} PDF documents into: {outputPath}");
    }
}

