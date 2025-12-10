using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Notes;
using Aspose.Words.Replacing;

namespace AsposeMcpServer.Tools;

public class WordAddFootnoteTool : IAsposeTool
{
    public string Description => "Add footnote to Word document";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            },
            referenceText = new
            {
                type = "string",
                description = "Reference text in document (optional, if not provided inserts at paragraph end)"
            },
            footnoteText = new
            {
                type = "string",
                description = "Footnote text"
            },
            customMark = new
            {
                type = "string",
                description = "Custom footnote mark (optional, if not provided uses auto-numbering)"
            }
        },
        required = new[] { "path", "footnoteText" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>() ?? 0;
        var referenceText = arguments?["referenceText"]?.GetValue<string>();
        var footnoteText = arguments?["footnoteText"]?.GetValue<string>() ?? throw new ArgumentException("footnoteText is required");
        var customMark = arguments?["customMark"]?.GetValue<string>();

        var doc = new Document(path);
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var builder = new DocumentBuilder(doc);

        if (!string.IsNullOrEmpty(referenceText))
        {
            // Find the reference text and insert footnote there
            var finder = new FindReplaceOptions { MatchCase = false };
            var found = doc.Range.Replace(referenceText, referenceText, finder);
            if (found > 0)
            {
                builder.MoveToDocumentEnd();
                var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
                if (!string.IsNullOrEmpty(customMark))
                {
                    footnote.ReferenceMark = customMark;
                }
            }
            else
            {
                throw new ArgumentException($"Reference text '{referenceText}' not found");
            }
        }
        else if (paragraphIndex.HasValue)
        {
            var section = doc.Sections[sectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
            }

            var para = paragraphs[paragraphIndex.Value];
            builder.MoveTo(para);
            var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                footnote.ReferenceMark = customMark;
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
            var footnote = builder.InsertFootnote(FootnoteType.Footnote, footnoteText);
            if (!string.IsNullOrEmpty(customMark))
            {
                footnote.ReferenceMark = customMark;
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Footnote added: {path}");
    }
}

