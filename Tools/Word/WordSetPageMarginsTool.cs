using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetPageMarginsTool : IAsposeTool
{
    public string Description => "Set page margins for section(s) in Word document";

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
            top = new
            {
                type = "number",
                description = "Top margin in points (optional)"
            },
            bottom = new
            {
                type = "number",
                description = "Bottom margin in points (optional)"
            },
            left = new
            {
                type = "number",
                description = "Left margin in points (optional)"
            },
            right = new
            {
                type = "number",
                description = "Right margin in points (optional)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided applies to all sections)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var top = arguments?["top"]?.GetValue<double?>();
        var bottom = arguments?["bottom"]?.GetValue<double?>();
        var left = arguments?["left"]?.GetValue<double?>();
        var right = arguments?["right"]?.GetValue<double?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        List<int> sectionsToUpdate;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            sectionsToUpdate = new List<int> { sectionIndex.Value };
        }
        else
        {
            sectionsToUpdate = Enumerable.Range(0, doc.Sections.Count).ToList();
        }

        foreach (var idx in sectionsToUpdate)
        {
            var pageSetup = doc.Sections[idx].PageSetup;
            if (top.HasValue) pageSetup.TopMargin = top.Value;
            if (bottom.HasValue) pageSetup.BottomMargin = bottom.Value;
            if (left.HasValue) pageSetup.LeftMargin = left.Value;
            if (right.HasValue) pageSetup.RightMargin = right.Value;
        }

        doc.Save(path);
        return await Task.FromResult($"Page margins updated for {sectionsToUpdate.Count} section(s): {path}");
    }
}

