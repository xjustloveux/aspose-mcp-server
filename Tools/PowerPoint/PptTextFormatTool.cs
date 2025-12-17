using System.Drawing;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint text formatting (batch format text)
///     Merges: PptBatchFormatTextTool
/// </summary>
public class PptTextFormatTool : IAsposeTool
{
    public string Description => @"Batch format PowerPoint text. Formats font, size, bold, italic, color across slides.

Usage examples:
- Format all slides: ppt_text_format(path='presentation.pptx', fontName='Arial', fontSize=14, bold=true)
- Format specific slides: ppt_text_format(path='presentation.pptx', slideIndices=[0,1,2], fontName='Times New Roman', fontSize=12)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path (required)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices to apply (optional; default all)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
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
            color = new
            {
                type = "string",
                description = "Hex color, e.g. #FF5500 (optional)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
        var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
        var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
        var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
        var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
        var colorHex = ArgumentHelper.GetStringNullable(arguments, "color");

        using var presentation = new Presentation(path);
        var targets = slideIndices?.Length > 0
            ? slideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        Color? color = null;
        if (!string.IsNullOrWhiteSpace(colorHex)) color = ColorHelper.ParseColor(colorHex);

        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];
            foreach (var shape in slide.Shapes)
                if (shape is IAutoShape { TextFrame: not null } auto)
                    foreach (var para in auto.TextFrame.Paragraphs)
                    foreach (var portion in para.Portions)
                    {
                        // Apply font settings using FontHelper
                        var colorStr = color.HasValue
                            ? $"#{color.Value.R:X2}{color.Value.G:X2}{color.Value.B:X2}"
                            : null;
                        FontHelper.Ppt.ApplyFontSettings(
                            portion.PortionFormat,
                            fontName,
                            fontSize,
                            bold,
                            italic,
                            colorStr
                        );
                    }
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Batch formatted text, applied to {targets.Length} slides");
    }
}