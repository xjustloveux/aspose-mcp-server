using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordGetRunFormatTool : IAsposeTool
{
    public string Description => "Get run (text) format information from Word document";

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
            runIndex = new
            {
                type = "number",
                description = "Run index within paragraph (0-based, optional, if not provided returns all runs)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var runIndex = arguments?["runIndex"]?.GetValue<int?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();

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
        var sb = new StringBuilder();

        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"runIndex must be between 0 and {runs.Count - 1}");
            }

            var run = runs[runIndex.Value];
            sb.AppendLine($"=== Run {runIndex.Value} Format ===");
            sb.AppendLine($"  Text: {run.Text}");
            sb.AppendLine($"  Font Name: {run.Font.Name}");
            sb.AppendLine($"  Font Name (ASCII): {run.Font.NameAscii}");
            sb.AppendLine($"  Font Name (Far East): {run.Font.NameFarEast}");
            sb.AppendLine($"  Font Size: {run.Font.Size} pt");
            sb.AppendLine($"  Bold: {run.Font.Bold}");
            sb.AppendLine($"  Italic: {run.Font.Italic}");
            sb.AppendLine($"  Underline: {run.Font.Underline}");
            sb.AppendLine($"  Color: #{run.Font.Color.R:X2}{run.Font.Color.G:X2}{run.Font.Color.B:X2}");
        }
        else
        {
            sb.AppendLine($"=== Runs in Paragraph {paragraphIndex} ({runs.Count}) ===");
            for (int i = 0; i < runs.Count; i++)
            {
                var run = runs[i];
                sb.AppendLine($"\n[{i}] Text: {run.Text}");
                sb.AppendLine($"    Font: {run.Font.NameAscii}/{run.Font.NameFarEast}, Size: {run.Font.Size}pt");
                sb.AppendLine($"    Bold: {run.Font.Bold}, Italic: {run.Font.Italic}");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

