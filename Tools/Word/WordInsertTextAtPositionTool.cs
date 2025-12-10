using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordInsertTextAtPositionTool : IAsposeTool
{
    public string Description => "Insert text at specific position in Word document";

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
            charIndex = new
            {
                type = "number",
                description = "Character index within paragraph (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            },
            text = new
            {
                type = "string",
                description = "Text to insert"
            },
            insertBefore = new
            {
                type = "boolean",
                description = "Insert before position (optional, default: false, inserts after)"
            }
        },
        required = new[] { "path", "paragraphIndex", "charIndex", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var charIndex = arguments?["charIndex"]?.GetValue<int>() ?? throw new ArgumentException("charIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var insertBefore = arguments?["insertBefore"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
        }

        var para = paragraphs[paragraphIndex];
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var totalChars = 0;
        int targetRunIndex = -1;
        int targetRunCharIndex = 0;

        for (int i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;
            if (totalChars + runLength >= charIndex)
            {
                targetRunIndex = i;
                targetRunCharIndex = charIndex - totalChars;
                break;
            }
            totalChars += runLength;
        }

        if (targetRunIndex == -1)
        {
            // Insert at end of paragraph
            var builder = new DocumentBuilder(doc);
            builder.MoveTo(para);
            builder.Write(text);
        }
        else
        {
            var targetRun = runs[targetRunIndex];
            if (insertBefore)
            {
                targetRun.Text = targetRun.Text.Insert(targetRunCharIndex, text);
            }
            else
            {
                targetRun.Text = targetRun.Text.Insert(targetRunCharIndex, text);
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Text inserted at position: {path}");
    }
}

