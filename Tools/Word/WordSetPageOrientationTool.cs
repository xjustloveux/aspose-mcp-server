using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetPageOrientationTool : IAsposeTool
{
    public string Description => "Set page orientation (portrait/landscape) for section(s) in Word document";

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
            orientation = new
            {
                type = "string",
                description = "Orientation: 'Portrait' or 'Landscape'",
                @enum = new[] { "Portrait", "Landscape" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided applies to all sections)"
            },
            sectionIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of section indices (0-based, optional, overrides sectionIndex)"
            }
        },
        required = new[] { "path", "orientation" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var orientation = arguments?["orientation"]?.GetValue<string>() ?? throw new ArgumentException("orientation is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var sectionIndicesArray = arguments?["sectionIndices"]?.AsArray();

        var doc = new Document(path);
        var orientationEnum = orientation.ToLower() == "landscape" ? Orientation.Landscape : Orientation.Portrait;

        List<int> sectionsToUpdate;
        if (sectionIndicesArray != null && sectionIndicesArray.Count > 0)
        {
            sectionsToUpdate = sectionIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).ToList();
        }
        else if (sectionIndex.HasValue)
        {
            sectionsToUpdate = new List<int> { sectionIndex.Value };
        }
        else
        {
            sectionsToUpdate = Enumerable.Range(0, doc.Sections.Count).ToList();
        }

        foreach (var idx in sectionsToUpdate)
        {
            if (idx >= 0 && idx < doc.Sections.Count)
            {
                doc.Sections[idx].PageSetup.Orientation = orientationEnum;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Page orientation set to {orientation} for {sectionsToUpdate.Count} section(s): {path}");
    }
}

