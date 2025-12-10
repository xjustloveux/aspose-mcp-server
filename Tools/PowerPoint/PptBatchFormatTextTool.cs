using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class PptBatchFormatTextTool : IAsposeTool
{
    public string Description => "Batch format text (font, size, bold, italic, color) across slides";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices to apply (optional; default all)"
            },
            fontName = new { type = "string", description = "Font name (optional)" },
            fontSize = new { type = "number", description = "Font size (optional)" },
            bold = new { type = "boolean", description = "Bold (optional)" },
            italic = new { type = "boolean", description = "Italic (optional)" },
            color = new { type = "string", description = "Hex color, e.g. #FF5500 (optional)" }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var colorHex = arguments?["color"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        var targets = slideIndices?.Length > 0
            ? slideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        foreach (var idx in targets)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

        Color? color = null;
        if (!string.IsNullOrWhiteSpace(colorHex))
        {
            color = ColorTranslator.FromHtml(colorHex);
        }

        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];
            foreach (var shape in slide.Shapes)
            {
                if (shape is IAutoShape auto && auto.TextFrame != null)
                {
                    foreach (var para in auto.TextFrame.Paragraphs)
                    {
                        foreach (var portion in para.Portions)
                        {
                            if (!string.IsNullOrWhiteSpace(fontName)) portion.PortionFormat.LatinFont = new FontData(fontName);
                            if (fontSize.HasValue) portion.PortionFormat.FontHeight = (float)fontSize.Value;
                            if (bold.HasValue) portion.PortionFormat.FontBold = bold.Value ? NullableBool.True : NullableBool.False;
                            if (italic.HasValue) portion.PortionFormat.FontItalic = italic.Value ? NullableBool.True : NullableBool.False;
                            if (color.HasValue)
                            {
                                portion.PortionFormat.FillFormat.FillType = FillType.Solid;
                                portion.PortionFormat.FillFormat.SolidFillColor.Color = color.Value;
                            }
                        }
                    }
                }
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已批次格式化文字，套用投影片數：{targets.Length}");
    }
}

