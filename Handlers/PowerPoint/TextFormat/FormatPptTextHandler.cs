using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.TextFormat;

/// <summary>
///     Handler for batch formatting text in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class FormatPptTextHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "format";

    /// <summary>
    ///     Batch formats text across slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndices, shapeIndices, fontName, fontSize, bold, italic, color, alignment.
    /// </param>
    /// <returns>Success message with formatting details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var formatParams = ExtractFormatParameters(parameters);

        var presentation = context.Document;

        var targets = ParseIndices(formatParams.SlideIndicesJson, presentation.Slides.Count);
        var shapeFilter = ParseOptionalIndices(formatParams.ShapeIndicesJson);

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        Color? parsedColor = null;
        if (!string.IsNullOrWhiteSpace(formatParams.Color)) parsedColor = ColorHelper.ParseColor(formatParams.Color);

        var colorStr = parsedColor.HasValue
            ? $"#{parsedColor.Value.R:X2}{parsedColor.Value.G:X2}{parsedColor.Value.B:X2}"
            : null;

        TextAlignment? parsedAlignment = null;
        if (!string.IsNullOrWhiteSpace(formatParams.Alignment))
            parsedAlignment = ParseTextAlignment(formatParams.Alignment);

        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];

            if (shapeFilter != null)
                foreach (var si in shapeFilter)
                {
                    if (si < 0 || si >= slide.Shapes.Count)
                        throw new ArgumentException(
                            $"shape index {si} out of range on slide {idx} (total: {slide.Shapes.Count})");

                    var shape = slide.Shapes[si];
                    ApplyFormatToShape(shape, formatParams, colorStr, parsedAlignment);
                }
            else
                foreach (var shape in slide.Shapes)
                    ApplyFormatToShape(shape, formatParams, colorStr, parsedAlignment);
        }

        MarkModified(context);

        var shapeInfo = shapeFilter != null ? $" ({shapeFilter.Length} shapes targeted)" : "";
        return new SuccessResult
        {
            Message = $"Batch formatted text applied to {targets.Length} slide(s){shapeInfo}."
        };
    }

    /// <summary>
    ///     Applies font and paragraph formatting to a single shape.
    /// </summary>
    /// <param name="shape">The shape to format.</param>
    /// <param name="formatParams">The format parameters.</param>
    /// <param name="colorStr">The parsed color string.</param>
    /// <param name="alignment">The parsed text alignment.</param>
    private static void ApplyFormatToShape(IShape shape, FormatParameters formatParams, string? colorStr,
        TextAlignment? alignment)
    {
        if (shape is IAutoShape { TextFrame: not null } auto)
            ApplyFontToTextFrame(auto.TextFrame, formatParams.FontName, formatParams.FontSize,
                formatParams.Bold, formatParams.Italic, colorStr, alignment);
        else if (shape is ITable table)
            ApplyFontToTable(table, formatParams.FontName, formatParams.FontSize, formatParams.Bold,
                formatParams.Italic, colorStr, alignment);
    }

    /// <summary>
    ///     Applies font and paragraph settings to all portions in a text frame.
    /// </summary>
    /// <param name="textFrame">The text frame.</param>
    /// <param name="fontName">The optional font name.</param>
    /// <param name="fontSize">The optional font size.</param>
    /// <param name="bold">The optional bold setting.</param>
    /// <param name="italic">The optional italic setting.</param>
    /// <param name="colorStr">The optional color string.</param>
    /// <param name="alignment">The optional text alignment.</param>
    private static void ApplyFontToTextFrame(ITextFrame textFrame, string? fontName, double? fontSize, bool? bold,
        bool? italic, string? colorStr, TextAlignment? alignment)
    {
        foreach (var para in textFrame.Paragraphs)
        {
            if (alignment.HasValue)
                para.ParagraphFormat.Alignment = alignment.Value;

            foreach (var portion in para.Portions)
                FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, fontName, fontSize, bold, italic, colorStr);
        }
    }

    /// <summary>
    ///     Applies font and paragraph settings to all cells in a table.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="fontName">The optional font name.</param>
    /// <param name="fontSize">The optional font size.</param>
    /// <param name="bold">The optional bold setting.</param>
    /// <param name="italic">The optional italic setting.</param>
    /// <param name="colorStr">The optional color string.</param>
    /// <param name="alignment">The optional text alignment.</param>
    private static void ApplyFontToTable(ITable table, string? fontName, double? fontSize, bool? bold, bool? italic,
        string? colorStr, TextAlignment? alignment)
    {
        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                ApplyFontToTextFrame(cell.TextFrame, fontName, fontSize, bold, italic, colorStr, alignment);
        }
    }

    /// <summary>
    ///     Parses indices from JSON string or returns all indices up to the given count.
    /// </summary>
    /// <param name="indicesJson">The optional JSON string containing indices.</param>
    /// <param name="totalCount">The total count for generating all indices as fallback.</param>
    /// <returns>Array of indices.</returns>
    private static int[] ParseIndices(string? indicesJson, int totalCount)
    {
        if (!string.IsNullOrWhiteSpace(indicesJson))
        {
            var indices = JsonSerializer.Deserialize<int[]>(indicesJson);
            return indices ?? Enumerable.Range(0, totalCount).ToArray();
        }

        return Enumerable.Range(0, totalCount).ToArray();
    }

    /// <summary>
    ///     Parses optional indices from JSON string. Returns null if not provided.
    /// </summary>
    /// <param name="indicesJson">The optional JSON string containing indices.</param>
    /// <returns>Array of indices, or null if not provided.</returns>
    private static int[]? ParseOptionalIndices(string? indicesJson)
    {
        if (!string.IsNullOrWhiteSpace(indicesJson))
            return JsonSerializer.Deserialize<int[]>(indicesJson);

        return null;
    }

    /// <summary>
    ///     Parses a text alignment string to a <see cref="TextAlignment" /> enum value.
    /// </summary>
    /// <param name="alignment">The alignment string (case-insensitive).</param>
    /// <returns>The parsed TextAlignment value.</returns>
    /// <exception cref="ArgumentException">Thrown when the alignment string is not recognized.</exception>
    private static TextAlignment ParseTextAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "left" => TextAlignment.Left,
            "center" => TextAlignment.Center,
            "right" => TextAlignment.Right,
            "justify" => TextAlignment.Justify,
            "distributed" => TextAlignment.Distributed,
            _ => throw new ArgumentException(
                $"Unsupported alignment: {alignment}. " +
                "Supported values: Left, Center, Right, Justify, Distributed.")
        };
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
            parameters.GetOptional<string?>("shapeIndices"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("color"),
            parameters.GetOptional<string?>("alignment")
        );
    }

    /// <summary>
    ///     Record for holding format text parameters.
    /// </summary>
    /// <param name="SlideIndicesJson">The optional JSON string containing slide indices.</param>
    /// <param name="ShapeIndicesJson">The optional JSON string containing shape indices.</param>
    /// <param name="FontName">The optional font name.</param>
    /// <param name="FontSize">The optional font size.</param>
    /// <param name="Bold">The optional bold setting.</param>
    /// <param name="Italic">The optional italic setting.</param>
    /// <param name="Color">The optional color string.</param>
    /// <param name="Alignment">The optional text alignment (Left, Center, Right, Justify, Distributed).</param>
    private sealed record FormatParameters(
        string? SlideIndicesJson,
        string? ShapeIndicesJson,
        string? FontName,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Color,
        string? Alignment);
}
