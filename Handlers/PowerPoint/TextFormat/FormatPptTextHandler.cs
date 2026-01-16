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
    ///     Optional: slideIndices, fontName, fontSize, bold, italic, color.
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var formatParams = ExtractFormatParameters(parameters);

        var presentation = context.Document;

        var targets = ParseSlideIndices(formatParams.SlideIndicesJson, presentation.Slides.Count);

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        Color? parsedColor = null;
        if (!string.IsNullOrWhiteSpace(formatParams.Color)) parsedColor = ColorHelper.ParseColor(formatParams.Color);

        var colorStr = parsedColor.HasValue
            ? $"#{parsedColor.Value.R:X2}{parsedColor.Value.G:X2}{parsedColor.Value.B:X2}"
            : null;

        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];
            foreach (var shape in slide.Shapes)
                if (shape is IAutoShape { TextFrame: not null } auto)
                    ApplyFontToTextFrame(auto.TextFrame, formatParams.FontName, formatParams.FontSize,
                        formatParams.Bold, formatParams.Italic, colorStr);
                else if (shape is ITable table)
                    ApplyFontToTable(table, formatParams.FontName, formatParams.FontSize, formatParams.Bold,
                        formatParams.Italic, colorStr);
        }

        MarkModified(context);

        return Success($"Batch formatted text applied to {targets.Length} slides.");
    }

    /// <summary>
    ///     Applies font settings to all portions in a text frame.
    /// </summary>
    /// <param name="textFrame">The text frame.</param>
    /// <param name="fontName">The optional font name.</param>
    /// <param name="fontSize">The optional font size.</param>
    /// <param name="bold">The optional bold setting.</param>
    /// <param name="italic">The optional italic setting.</param>
    /// <param name="colorStr">The optional color string.</param>
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
    /// <param name="table">The table.</param>
    /// <param name="fontName">The optional font name.</param>
    /// <param name="fontSize">The optional font size.</param>
    /// <param name="bold">The optional bold setting.</param>
    /// <param name="italic">The optional italic setting.</param>
    /// <param name="colorStr">The optional color string.</param>
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

    /// <summary>
    ///     Parses slide indices from JSON string or returns all slide indices.
    /// </summary>
    /// <param name="slideIndicesJson">The optional JSON string containing slide indices.</param>
    /// <param name="slideCount">The total number of slides.</param>
    /// <returns>Array of slide indices.</returns>
    private static int[] ParseSlideIndices(string? slideIndicesJson, int slideCount)
    {
        if (!string.IsNullOrWhiteSpace(slideIndicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(slideIndicesJson);
            return indices ?? Enumerable.Range(0, slideCount).ToArray();
        }

        return Enumerable.Range(0, slideCount).ToArray();
    }

    /// <summary>
    ///     Extracts format parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted format parameters.</returns>
    private static FormatParameters ExtractFormatParameters(OperationParameters parameters)
    {
        return new FormatParameters(
            parameters.GetOptional<string?>("slideIndices"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("color")
        );
    }

    /// <summary>
    ///     Record for holding format text parameters.
    /// </summary>
    /// <param name="SlideIndicesJson">The optional JSON string containing slide indices.</param>
    /// <param name="FontName">The optional font name.</param>
    /// <param name="FontSize">The optional font size.</param>
    /// <param name="Bold">The optional bold setting.</param>
    /// <param name="Italic">The optional italic setting.</param>
    /// <param name="Color">The optional color string.</param>
    private sealed record FormatParameters(
        string? SlideIndicesJson,
        string? FontName,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Color);
}
