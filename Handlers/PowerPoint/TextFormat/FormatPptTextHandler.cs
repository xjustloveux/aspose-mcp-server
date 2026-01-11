using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.TextFormat;

/// <summary>
///     Handler for batch formatting text in PowerPoint presentations.
/// </summary>
public class FormatPptTextHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "format";

    /// <summary>
    ///     Batch formats text across slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndices, fontName, fontSize, bold, italic, color
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndicesJson = parameters.GetOptional<string?>("slideIndices");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var color = parameters.GetOptional<string?>("color");

        var presentation = context.Document;

        int[] targets;
        if (!string.IsNullOrWhiteSpace(slideIndicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(slideIndicesJson);
            targets = indices ?? Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }
        else
        {
            targets = Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        Color? parsedColor = null;
        if (!string.IsNullOrWhiteSpace(color)) parsedColor = ColorHelper.ParseColor(color);

        var colorStr = parsedColor.HasValue
            ? $"#{parsedColor.Value.R:X2}{parsedColor.Value.G:X2}{parsedColor.Value.B:X2}"
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

        MarkModified(context);

        return Success($"Batch formatted text applied to {targets.Length} slides.");
    }

    /// <summary>
    ///     Applies font settings to all portions in a text frame.
    /// </summary>
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
