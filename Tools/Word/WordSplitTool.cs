using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordSplitTool : IAsposeTool
{
    public string Description => "Split a Word document by sections or pages";

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
            splitBy = new
            {
                type = "string",
                description = "Split by 'section' or 'page' (default: section)"
            }
        },
        required = new[] { "inputPath", "outputDir" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var inputPath = arguments?["inputPath"]?.GetValue<string>() ?? throw new ArgumentException("inputPath is required");
        var outputDir = arguments?["outputDir"]?.GetValue<string>() ?? throw new ArgumentException("outputDir is required");
        var splitBy = arguments?["splitBy"]?.GetValue<string>() ?? "section";

        Directory.CreateDirectory(outputDir);

        var doc = new Document(inputPath);
        var fileBaseName = SecurityHelper.SanitizeFileName(Path.GetFileNameWithoutExtension(inputPath));

        if (splitBy.ToLower() == "section")
        {
            for (int i = 0; i < doc.Sections.Count; i++)
            {
                var sectionDoc = new Document();
                sectionDoc.FirstSection.Remove();
                var importedSection = sectionDoc.ImportNode(doc.Sections[i], true);
                sectionDoc.AppendChild(importedSection);

                var outputPath = Path.Combine(outputDir, $"{fileBaseName}_section_{i + 1}.docx");
                sectionDoc.Save(outputPath);
            }

            return await Task.FromResult($"Document split into {doc.Sections.Count} sections in: {outputDir}");
        }
        else
        {
            var pageCount = doc.PageCount;
            for (int i = 0; i < pageCount; i++)
            {
                var pageDoc = doc.ExtractPages(i, 1);
                var outputPath = Path.Combine(outputDir, $"{fileBaseName}_page_{i + 1}.docx");
                pageDoc.Save(outputPath);
            }

            return await Task.FromResult($"Document split into {pageCount} pages in: {outputDir}");
        }
    }
}

