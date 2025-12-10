using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PdfSplitTool : IAsposeTool
{
    public string Description => "Split a PDF document into multiple files";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input file path"
            },
            outputDir = new
            {
                type = "string",
                description = "Output directory for split files"
            },
            pagesPerFile = new
            {
                type = "number",
                description = "Number of pages per file (default: 1)"
            }
        },
        required = new[] { "inputPath", "outputDir" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var pagesPerFile = arguments?["pagesPerFile"]?.GetValue<int>() ?? 1;

        Directory.CreateDirectory(outputDir);

        using var document = new Document(inputPath);
        var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(inputPath));
        int fileCount = 0;

        for (int i = 0; i < document.Pages.Count; i += pagesPerFile)
        {
            using var newDocument = new Document();
            
            for (int j = 0; j < pagesPerFile && (i + j) < document.Pages.Count; j++)
            {
                newDocument.Pages.Add(document.Pages[i + j + 1]);
            }

            var safeFileName = SecurityHelper.SanitizeFileName($"{fileBaseName}_part_{++fileCount}.pdf");
            var outputPath = Path.Combine(outputDir, safeFileName);
            newDocument.Save(outputPath);
        }

        return await Task.FromResult($"PDF split into {fileCount} files in: {outputDir}");
    }
}

