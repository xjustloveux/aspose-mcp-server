using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordMoveTableTool : IAsposeTool
{
    public string Description => "Move table to different position in Word document";

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
            tableIndex = new
            {
                type = "number",
                description = "Table index to move (0-based)"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index to move after (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            }
        },
        required = new[] { "path", "tableIndex", "targetParagraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var targetParagraphIndex = arguments?["targetParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetParagraphIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var section = doc.Sections[sectionIdx];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
        }
        if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"targetParagraphIndex must be between 0 and {paragraphs.Count - 1}");
        }

        var table = tables[tableIndex];
        var targetPara = paragraphs[targetParagraphIndex];

        section.Body.InsertAfter(table, targetPara);

        doc.Save(path);
        return await Task.FromResult($"Table moved: {path}");
    }
}

