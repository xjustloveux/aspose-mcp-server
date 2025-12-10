using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordApplyStyleTool : IAsposeTool
{
    public string Description => "Apply style to paragraph(s), run(s), or table in Word document";

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
            styleName = new
            {
                type = "string",
                description = "Style name to apply"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based, optional)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            },
            paragraphIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of paragraph indices to apply style (optional, overrides paragraphIndex)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based, optional, applies style to table)"
            },
            applyToAllParagraphs = new
            {
                type = "boolean",
                description = "Apply to all paragraphs in document (optional, default: false)"
            }
        },
        required = new[] { "path", "styleName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var styleName = arguments?["styleName"]?.GetValue<string>() ?? throw new ArgumentException("styleName is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var paragraphIndicesArray = arguments?["paragraphIndices"]?.AsArray();
        var tableIndex = arguments?["tableIndex"]?.GetValue<int?>();
        var applyToAllParagraphs = arguments?["applyToAllParagraphs"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var style = doc.Styles[styleName];
        if (style == null)
        {
            throw new ArgumentException($"Style '{styleName}' not found");
        }

        int appliedCount = 0;

        if (tableIndex.HasValue)
        {
            var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
            if (tableIndex.Value < 0 || tableIndex.Value >= tables.Count)
            {
                throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
            }
            tables[tableIndex.Value].Style = style;
            appliedCount = 1;
        }
        else if (applyToAllParagraphs)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            foreach (var para in paragraphs)
            {
                para.ParagraphFormat.Style = style;
                appliedCount++;
            }
        }
        else if (paragraphIndicesArray != null && paragraphIndicesArray.Count > 0)
        {
            var sectionIdx = sectionIndex.HasValue ? sectionIndex.Value : 0;
            if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }

            var section = doc.Sections[sectionIdx];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            
            foreach (var idxObj in paragraphIndicesArray)
            {
                var idx = idxObj?.GetValue<int>();
                if (idx.HasValue && idx.Value >= 0 && idx.Value < paragraphs.Count)
                {
                    paragraphs[idx.Value].ParagraphFormat.Style = style;
                    appliedCount++;
                }
            }
        }
        else if (paragraphIndex.HasValue)
        {
            var sectionIdx = sectionIndex.HasValue ? sectionIndex.Value : 0;
            if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }

            var section = doc.Sections[sectionIdx];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
            }

            paragraphs[paragraphIndex.Value].ParagraphFormat.Style = style;
            appliedCount = 1;
        }
        else
        {
            throw new ArgumentException("Either paragraphIndex, paragraphIndices, tableIndex, or applyToAllParagraphs must be provided");
        }

        doc.Save(path);
        return await Task.FromResult($"Applied style '{styleName}' to {appliedCount} element(s): {path}");
    }
}

