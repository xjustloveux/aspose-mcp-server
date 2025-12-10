using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetRunFormatTool : IAsposeTool
{
    public string Description => "Set run (text) format in Word document";

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
                description = "Run index within paragraph (0-based, optional, if not provided applies to all runs)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, default: 0)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (optional)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic (optional)"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline (optional)"
            },
            color = new
            {
                type = "string",
                description = "Font color hex (e.g., '000000', optional)"
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
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var underline = arguments?["underline"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();

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

        List<Run> runsToFormat;
        if (runIndex.HasValue)
        {
            if (runIndex.Value < 0 || runIndex.Value >= runs.Count)
            {
                throw new ArgumentException($"runIndex must be between 0 and {runs.Count - 1}");
            }
            runsToFormat = new List<Run> { runs[runIndex.Value] };
        }
        else
        {
            runsToFormat = runs;
        }

        foreach (var run in runsToFormat)
        {
            if (!string.IsNullOrEmpty(fontName)) run.Font.Name = fontName;
            if (!string.IsNullOrEmpty(fontNameAscii)) run.Font.NameAscii = fontNameAscii;
            if (!string.IsNullOrEmpty(fontNameFarEast)) run.Font.NameFarEast = fontNameFarEast;
            if (fontSize.HasValue) run.Font.Size = fontSize.Value;
            if (bold.HasValue) run.Font.Bold = bold.Value;
            if (italic.HasValue) run.Font.Italic = italic.Value;
            if (underline.HasValue) run.Font.Underline = underline.Value ? Underline.Single : Underline.None;
            if (!string.IsNullOrEmpty(color))
            {
                var colorStr = color.TrimStart('#');
                if (colorStr.Length == 6)
                {
                    var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
                    var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
                    var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
                    run.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
            }
        }

        doc.Save(path);
        return await Task.FromResult($"Run format updated: {path}");
    }
}

