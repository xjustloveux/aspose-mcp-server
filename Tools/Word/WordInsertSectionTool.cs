using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordInsertSectionTool : IAsposeTool
{
    public string Description => "Insert new section break in Word document";

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
            sectionBreakType = new
            {
                type = "string",
                description = "Section break type: 'NextPage', 'Continuous', 'EvenPage', 'OddPage'",
                @enum = new[] { "NextPage", "Continuous", "EvenPage", "OddPage" }
            },
            insertAtParagraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert section break after (0-based, optional, default: end of document)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: last section)"
            }
        },
        required = new[] { "path", "sectionBreakType" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionBreakType = arguments?["sectionBreakType"]?.GetValue<string>() ?? throw new ArgumentException("sectionBreakType is required");
        var insertAtParagraphIndex = arguments?["insertAtParagraphIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        var breakType = sectionBreakType switch
        {
            "NextPage" => SectionStart.NewPage,
            "Continuous" => SectionStart.Continuous,
            "EvenPage" => SectionStart.EvenPage,
            "OddPage" => SectionStart.OddPage,
            _ => SectionStart.NewPage
        };

        if (insertAtParagraphIndex.HasValue)
        {
            var actualSectionIndex = sectionIndex ?? 0;
            if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
            {
                actualSectionIndex = 0;
            }

            var section = doc.Sections[actualSectionIndex];
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            
            if (insertAtParagraphIndex.Value < 0 || insertAtParagraphIndex.Value >= paragraphs.Count)
            {
                throw new ArgumentException($"insertAtParagraphIndex must be between 0 and {paragraphs.Count - 1}");
            }

            var para = paragraphs[insertAtParagraphIndex.Value];
            builder.MoveTo(para);
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.CurrentSection.PageSetup.SectionStart = breakType;
        }
        else
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.CurrentSection.PageSetup.SectionStart = breakType;
        }

        doc.Save(path);
        return await Task.FromResult($"Section break inserted ({sectionBreakType}): {path}");
    }
}

