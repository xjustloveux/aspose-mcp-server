using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptMergeTool : IAsposeTool
{
    public string Description => "Merge multiple PowerPoint presentations into one";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            outputPath = new
            {
                type = "string",
                description = "Output presentation file path"
            },
            inputPaths = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Array of input presentation file paths"
            },
            keepSourceFormatting = new
            {
                type = "boolean",
                description = "Keep source formatting (optional, default: true)"
            }
        },
        required = new[] { "outputPath", "inputPaths" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? throw new ArgumentException("outputPath is required");
        var inputPathsArray = arguments?["inputPaths"]?.AsArray() ?? throw new ArgumentException("inputPaths is required");
        var keepSourceFormatting = arguments?["keepSourceFormatting"]?.GetValue<bool?>() ?? true;

        if (inputPathsArray.Count == 0)
        {
            throw new ArgumentException("At least one input path is required");
        }

        var inputPaths = inputPathsArray.Select(p => p?.GetValue<string>()).Where(p => !string.IsNullOrEmpty(p)).ToList();
        if (inputPaths.Count == 0)
        {
            throw new ArgumentException("No valid input paths provided");
        }

        using var masterPresentation = new Presentation(inputPaths[0]!);

        for (int i = 1; i < inputPaths.Count; i++)
        {
            var inputPath = inputPaths[i];
            if (string.IsNullOrEmpty(inputPath) || !File.Exists(inputPath))
            {
                continue;
            }

            using var sourcePresentation = new Presentation(inputPath);
            foreach (var slide in sourcePresentation.Slides)
            {
                if (keepSourceFormatting)
                {
                    masterPresentation.Slides.AddClone(slide);
                }
                else
                {
                    masterPresentation.Slides.AddClone(slide, masterPresentation.Masters[0], true);
                }
            }
        }

        masterPresentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Merged {inputPaths.Count} presentations into: {outputPath} (Total slides: {masterPresentation.Slides.Count})");
    }
}

