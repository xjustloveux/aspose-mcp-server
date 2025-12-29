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
Applies to text in AutoShapes and Table cells.

Color format: Hex color code (e.g., #FF5500, #RGB, #RRGGBB) or named colors (e.g., Red, Blue, DarkGreen).

Usage examples:
- Format all slides: ppt_text_format(path='presentation.pptx', fontName='Arial', fontSize=14, bold=true)
- Format specific slides: ppt_text_format(path='presentation.pptx', slideIndices=[0,1,2], fontName='Times New Roman', fontSize=12)
- Format with color: ppt_text_format(path='presentation.pptx', color='#FF0000') or ppt_text_format(path='presentation.pptx', color='Red')";

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
                description = "Text color: Hex (#FF5500, #RGB) or named color (Red, Blue, DarkGreen) (optional)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when slide index is out of range.</exception>
    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var path = ArgumentHelper.GetAndValidatePath(arguments);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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

            var colorStr = color.HasValue
                ? $"#{color.Value.R:X2}{color.Value.G:X2}{color.Value.B:X2}"
                : null;

            foreach (var idx in targets)
            {
                var slide = presentation.Slides[idx];
                foreach (var shape in slide.Shapes)
                    if (shape is IAutoShape { TextFrame: not null } auto)
                        ApplyFontToTextFrame(auto.TextFrame, fontName, fontSize, bold, italic, colorStr);
                    else if (shape is ITable table)
                        ApplyFontToTable(table, fontName, fontSize, bold, italic, colorStr);
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Batch formatted text applied to {targets.Length} slides. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies font settings to all portions in a text frame.
    /// </summary>
    /// <param name="textFrame">The text frame to format.</param>
    /// <param name="fontName">Font name (null to skip).</param>
    /// <param name="fontSize">Font size (null to skip).</param>
    /// <param name="bold">Bold setting (null to skip).</param>
    /// <param name="italic">Italic setting (null to skip).</param>
    /// <param name="colorStr">Color string in hex format (null to skip).</param>
    private static void ApplyFontToTextFrame(ITextFrame textFrame, string? fontName, double? fontSize, bool? bold,
        bool? italic, string? colorStr)
    {
        foreach (var para in textFrame.Paragraphs)
        foreach (var portion in para.Portions)
            FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, fontName, fontSize, bold, italic, colorStr);
    }

    /// <summary>
    ///     Applies font settings to all cells in a table.
    /// </summary>
    /// <param name="table">The table to format.</param>
    /// <param name="fontName">Font name (null to skip).</param>
    /// <param name="fontSize">Font size (null to skip).</param>
    /// <param name="bold">Bold setting (null to skip).</param>
    /// <param name="italic">Italic setting (null to skip).</param>
    /// <param name="colorStr">Color string in hex format (null to skip).</param>
    private static void ApplyFontToTable(ITable table, string? fontName, double? fontSize, bool? bold, bool? italic,
        string? colorStr)
    {
        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                ApplyFontToTextFrame(cell.TextFrame, fontName, fontSize, bold, italic, colorStr);
        }
    }
}