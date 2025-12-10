using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteSectionTool : IAsposeTool
{
    public string Description => "Delete section(s) from Word document";

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
            sectionIndex = new
            {
                type = "number",
                description = "Section index to delete (0-based)"
            },
            sectionIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of section indices to delete (0-based, optional, overrides sectionIndex)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var sectionIndicesArray = arguments?["sectionIndices"]?.AsArray();

        var doc = new Document(path);
        if (doc.Sections.Count <= 1)
        {
            throw new ArgumentException("Cannot delete the last section. Document must have at least one section.");
        }

        List<int> sectionsToDelete;
        if (sectionIndicesArray != null && sectionIndicesArray.Count > 0)
        {
            sectionsToDelete = sectionIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue).Select(s => s!.Value).OrderByDescending(s => s).ToList();
        }
        else if (sectionIndex.HasValue)
        {
            sectionsToDelete = new List<int> { sectionIndex.Value };
        }
        else
        {
            throw new ArgumentException("Either sectionIndex or sectionIndices must be provided");
        }

        foreach (var idx in sectionsToDelete)
        {
            if (idx < 0 || idx >= doc.Sections.Count)
            {
                continue;
            }
            if (doc.Sections.Count <= 1)
            {
                break; // Don't delete the last section
            }
            doc.Sections.RemoveAt(idx);
        }

        doc.Save(path);
        return await Task.FromResult($"Deleted {sectionsToDelete.Count} section(s). Remaining sections: {doc.Sections.Count}");
    }
}

