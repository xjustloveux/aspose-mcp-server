using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class PptSplitTool : IAsposeTool
{
    public string Description => "Split a PowerPoint presentation into multiple files (one slide per file or by range)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            inputPath = new
            {
                type = "string",
                description = "Input presentation file path"
            },
            outputDirectory = new
            {
                type = "string",
                description = "Output directory path"
            },
            slidesPerFile = new
            {
                type = "number",
                description = "Number of slides per output file (optional, default: 1)"
            },
            startSlideIndex = new
            {
                type = "number",
                description = "Start slide index (0-based, optional)"
            },
            endSlideIndex = new
            {
                type = "number",
                description = "End slide index (0-based, optional)"
            },
            outputFileNamePattern = new
            {
                type = "string",
                description = "Output file name pattern, use {index} for slide number (optional, default: 'slide_{index}.pptx')"
            }
        },
        required = new[] { "inputPath", "outputDirectory" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputDirectory = arguments?["outputDirectory"]?.GetValue<string>() ?? throw new ArgumentException("outputDirectory is required");
        var slidesPerFile = arguments?["slidesPerFile"]?.GetValue<int?>() ?? 1;
        var startSlideIndex = arguments?["startSlideIndex"]?.GetValue<int?>();
        var endSlideIndex = arguments?["endSlideIndex"]?.GetValue<int?>();
        var fileNamePattern = arguments?["outputFileNamePattern"]?.GetValue<string>() ?? "slide_{index}.pptx";

        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var presentation = new Presentation(inputPath);
        var totalSlides = presentation.Slides.Count;

        var start = startSlideIndex ?? 0;
        var end = endSlideIndex ?? (totalSlides - 1);

        if (start < 0 || start >= totalSlides || end < 0 || end >= totalSlides || start > end)
        {
            throw new ArgumentException($"Invalid slide range: start={start}, end={end}, total={totalSlides}");
        }

        var fileCount = 0;
        for (int i = start; i <= end; i += slidesPerFile)
        {
            using var newPresentation = new Presentation();
            newPresentation.Slides.RemoveAt(0);

            for (int j = 0; j < slidesPerFile && (i + j) <= end; j++)
            {
                newPresentation.Slides.AddClone(presentation.Slides[i + j]);
            }

            var outputFileName = fileNamePattern.Replace("{index}", fileCount.ToString());
            outputFileName = SecurityHelper.SanitizeFileName(outputFileName);
            var outputPath = Path.Combine(outputDirectory, outputFileName);
            newPresentation.Save(outputPath, SaveFormat.Pptx);
            fileCount++;
        }

        return await Task.FromResult($"Split presentation into {fileCount} file(s) in: {outputDirectory}");
    }
}

