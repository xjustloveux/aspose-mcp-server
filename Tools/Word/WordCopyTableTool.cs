using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCopyTableTool : IAsposeTool
{
    public string Description => "Copy table from one location to another in Word document";

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
            sourceTableIndex = new
            {
                type = "number",
                description = "Source table index (0-based)"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index to insert after (0-based)"
            },
            sourceSectionIndex = new
            {
                type = "number",
                description = "Source section index (0-based, optional, default: 0)"
            },
            targetSectionIndex = new
            {
                type = "number",
                description = "Target section index (0-based, optional, default: 0)"
            }
        },
        required = new[] { "path", "sourceTableIndex", "targetParagraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sourceTableIndex = arguments?["sourceTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("sourceTableIndex is required");
        var targetParagraphIndex = arguments?["targetParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("targetParagraphIndex is required");
        var sourceSectionIndex = arguments?["sourceSectionIndex"]?.GetValue<int?>();
        var targetSectionIndex = arguments?["targetSectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var sourceSectionIdx = sourceSectionIndex ?? 0;
        var targetSectionIdx = targetSectionIndex ?? 0;
        
        if (sourceSectionIdx < 0 || sourceSectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sourceSectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }
        if (targetSectionIdx < 0 || targetSectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"targetSectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var sourceSection = doc.Sections[sourceSectionIdx];
        var sourceTables = sourceSection.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        
        if (sourceTableIndex < 0 || sourceTableIndex >= sourceTables.Count)
        {
            throw new ArgumentException($"sourceTableIndex must be between 0 and {sourceTables.Count - 1}");
        }

        var sourceTable = sourceTables[sourceTableIndex];
        var targetSection = doc.Sections[targetSectionIdx];
        var targetParagraphs = targetSection.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (targetParagraphIndex < 0 || targetParagraphIndex >= targetParagraphs.Count)
        {
            throw new ArgumentException($"targetParagraphIndex must be between 0 and {targetParagraphs.Count - 1}");
        }

        var targetPara = targetParagraphs[targetParagraphIndex];
        var clonedTable = (Aspose.Words.Tables.Table)sourceTable.Clone(true);
        targetSection.Body.InsertAfter(clonedTable, targetPara);

        doc.Save(path);
        return await Task.FromResult($"Table copied: {path}");
    }
}

