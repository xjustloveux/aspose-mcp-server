using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordDeleteTextRangeTool : IAsposeTool
{
    public string Description => "Delete text range from Word document";

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
            startParagraphIndex = new
            {
                type = "number",
                description = "Start paragraph index (0-based)"
            },
            startCharIndex = new
            {
                type = "number",
                description = "Start character index within paragraph (0-based)"
            },
            endParagraphIndex = new
            {
                type = "number",
                description = "End paragraph index (0-based)"
            },
            endCharIndex = new
            {
                type = "number",
                description = "End character index within paragraph (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            }
        },
        required = new[] { "path", "startParagraphIndex", "startCharIndex", "endParagraphIndex", "endCharIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var startParagraphIndex = arguments?["startParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("startParagraphIndex is required");
        var startCharIndex = arguments?["startCharIndex"]?.GetValue<int>() ?? throw new ArgumentException("startCharIndex is required");
        var endParagraphIndex = arguments?["endParagraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("endParagraphIndex is required");
        var endCharIndex = arguments?["endCharIndex"]?.GetValue<int>() ?? throw new ArgumentException("endCharIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
        {
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        }

        var section = doc.Sections[sectionIdx];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (startParagraphIndex < 0 || startParagraphIndex >= paragraphs.Count ||
            endParagraphIndex < 0 || endParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException("Paragraph indices out of range");
        }

        var startPara = paragraphs[startParagraphIndex];
        var endPara = paragraphs[endParagraphIndex];

        if (startParagraphIndex == endParagraphIndex)
        {
            // Delete within same paragraph
            var runs = startPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var totalChars = 0;
            int startRunIndex = -1, endRunIndex = -1;
            int startRunCharIndex = 0, endRunCharIndex = 0;

            for (int i = 0; i < runs.Count; i++)
            {
                var runLength = runs[i].Text.Length;
                if (startRunIndex == -1 && totalChars + runLength > startCharIndex)
                {
                    startRunIndex = i;
                    startRunCharIndex = startCharIndex - totalChars;
                }
                if (totalChars + runLength > endCharIndex)
                {
                    endRunIndex = i;
                    endRunCharIndex = endCharIndex - totalChars;
                    break;
                }
                totalChars += runLength;
            }

            if (startRunIndex >= 0 && endRunIndex >= 0)
            {
                if (startRunIndex == endRunIndex)
                {
                    var run = runs[startRunIndex];
                    run.Text = run.Text.Remove(startRunCharIndex, endRunCharIndex - startRunCharIndex);
                }
                else
                {
                    // Delete across multiple runs
                    var startRun = runs[startRunIndex];
                    startRun.Text = startRun.Text.Substring(0, startRunCharIndex);

                    for (int i = startRunIndex + 1; i < endRunIndex; i++)
                    {
                        runs[i].Remove();
                    }

                    if (endRunIndex < runs.Count)
                    {
                        var endRun = runs[endRunIndex];
                        endRun.Text = endRun.Text.Substring(endRunCharIndex);
                    }
                }
            }
        }
        else
        {
            // Delete across multiple paragraphs
            var startParaRuns = startPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            var startRun = startParaRuns.LastOrDefault();
            if (startRun != null && startRun.Text.Length > startCharIndex)
            {
                startRun.Text = startRun.Text.Substring(0, startCharIndex);
            }

            // Remove paragraphs in between
            for (int i = startParagraphIndex + 1; i < endParagraphIndex; i++)
            {
                paragraphs[i].Remove();
            }

            var endParaRuns = endPara.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
            if (endParaRuns.Count > 0 && endCharIndex < endParaRuns[0].Text.Length)
            {
                endParaRuns[0].Text = endParaRuns[0].Text.Substring(endCharIndex);
                // Remove other runs in end paragraph
                for (int i = 1; i < endParaRuns.Count; i++)
                {
                    endParaRuns[i].Remove();
                }
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Text range deleted: {path}");
    }
}

