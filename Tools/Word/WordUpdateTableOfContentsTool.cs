using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordUpdateTableOfContentsTool : IAsposeTool
{
    public string Description => "Update table of contents field in Word document";

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
            tocIndex = new
            {
                type = "number",
                description = "TOC field index (0-based, optional, if not provided updates all TOC fields)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var tocIndex = arguments?["tocIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var tocFields = doc.Range.Fields
            .Cast<Field>()
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        if (tocFields.Count == 0)
        {
            return await Task.FromResult("No table of contents fields found in document");
        }

        if (tocIndex.HasValue)
        {
            if (tocIndex.Value < 0 || tocIndex.Value >= tocFields.Count)
            {
                throw new ArgumentException($"tocIndex must be between 0 and {tocFields.Count - 1}");
            }
            tocFields[tocIndex.Value].Update();
        }
        else
        {
            foreach (var tocField in tocFields)
            {
                tocField.Update();
            }
        }

        doc.UpdateFields();
        doc.Save(path);
        var updatedCount = tocIndex.HasValue ? 1 : tocFields.Count;
        return await Task.FromResult($"Updated {updatedCount} table of contents field(s): {path}");
    }
}

