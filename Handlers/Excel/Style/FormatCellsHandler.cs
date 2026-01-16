using System.Drawing;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Style;

/// <summary>
///     Handler for formatting cells in Excel workbooks.
/// </summary>
public class FormatCellsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "format";

    /// <summary>
    ///     Formats cells in the specified range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range or ranges
    ///     Optional: sheetIndex, fontName, fontSize, bold, italic, fontColor,
    ///     backgroundColor, patternType, patternColor, numberFormat,
    ///     borderStyle, borderColor, horizontalAlignment, verticalAlignment
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var formatParams = ExtractFormatParameters(parameters);
        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, formatParams.SheetIndex);
        var style = workbook.CreateStyle();

        ApplyFontSettings(style, formatParams);
        ApplyBackgroundIfNeeded(style, formatParams);
        ApplyNumberFormat(style, formatParams.NumberFormat);
        ApplyAlignmentSettings(style, formatParams);
        ApplyBorderIfNeeded(style, formatParams);

        var styleFlag = new StyleFlag { All = true, Borders = !string.IsNullOrEmpty(formatParams.BorderStyle) };
        ApplyStyleToRanges(worksheet, formatParams.Range, formatParams.RangesJson, style, styleFlag);

        MarkModified(context);
        return Success($"Cells formatted in sheet {formatParams.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts format parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A FormatParameters record containing all extracted values.</returns>
    private static FormatParameters ExtractFormatParameters(OperationParameters parameters)
    {
        return new FormatParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("range"),
            parameters.GetOptional<string?>("ranges"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<int?>("fontSize"),
            parameters.GetOptional<bool?>("bold"),
            parameters.GetOptional<bool?>("italic"),
            parameters.GetOptional<string?>("fontColor"),
            parameters.GetOptional<string?>("backgroundColor"),
            parameters.GetOptional<string?>("patternType"),
            parameters.GetOptional<string?>("patternColor"),
            parameters.GetOptional<string?>("numberFormat"),
            parameters.GetOptional<string?>("borderStyle"),
            parameters.GetOptional<string?>("borderColor"),
            parameters.GetOptional<string?>("horizontalAlignment"),
            parameters.GetOptional<string?>("verticalAlignment")
        );
    }

    /// <summary>
    ///     Applies font settings to the style.
    /// </summary>
    /// <param name="style">The style to apply font settings to.</param>
    /// <param name="p">The format parameters containing font settings.</param>
    /// <exception cref="ArgumentException">Thrown when the font color format is invalid.</exception>
    private static void ApplyFontSettings(Aspose.Cells.Style style, FormatParameters p)
    {
        try
        {
            FontHelper.Excel.ApplyFontSettings(style, p.FontName, p.FontSize, p.Bold, p.Italic, p.FontColor);
        }
        catch (ArgumentException colorEx) when (!string.IsNullOrWhiteSpace(p.FontColor))
        {
            throw new ArgumentException(
                $"Unable to parse font color '{p.FontColor}': {colorEx.Message}. Please use a valid color format (e.g., #FF0000, 255,0,0, or red)");
        }
    }

    /// <summary>
    ///     Applies background settings if background color or pattern type is specified.
    /// </summary>
    /// <param name="style">The style to apply background settings to.</param>
    /// <param name="p">The format parameters containing background settings.</param>
    private static void ApplyBackgroundIfNeeded(Aspose.Cells.Style style, FormatParameters p)
    {
        if (!string.IsNullOrWhiteSpace(p.BackgroundColor) || !string.IsNullOrWhiteSpace(p.PatternType))
            ApplyBackgroundSettings(style, p.BackgroundColor, p.PatternType, p.PatternColor);
    }

    /// <summary>
    ///     Applies number format to the style.
    /// </summary>
    /// <param name="style">The style to apply number format to.</param>
    /// <param name="numberFormat">The number format string or built-in format number.</param>
    private static void ApplyNumberFormat(Aspose.Cells.Style style, string? numberFormat)
    {
        if (string.IsNullOrEmpty(numberFormat)) return;
        if (int.TryParse(numberFormat, out var formatNumber))
            style.Number = formatNumber;
        else
            style.Custom = numberFormat;
    }

    /// <summary>
    ///     Applies alignment settings to the style.
    /// </summary>
    /// <param name="style">The style to apply alignment to.</param>
    /// <param name="p">The format parameters containing alignment settings.</param>
    private static void ApplyAlignmentSettings(Aspose.Cells.Style style, FormatParameters p)
    {
        if (!string.IsNullOrEmpty(p.HorizontalAlignment))
            style.HorizontalAlignment = ParseHorizontalAlignment(p.HorizontalAlignment);
        if (!string.IsNullOrEmpty(p.VerticalAlignment))
            style.VerticalAlignment = ParseVerticalAlignment(p.VerticalAlignment);
    }

    /// <summary>
    ///     Applies border settings if border style is specified.
    /// </summary>
    /// <param name="style">The style to apply border to.</param>
    /// <param name="p">The format parameters containing border settings.</param>
    private static void ApplyBorderIfNeeded(Aspose.Cells.Style style, FormatParameters p)
    {
        if (!string.IsNullOrEmpty(p.BorderStyle))
            ApplyBorderSettings(style, p.BorderStyle, p.BorderColor);
    }

    /// <summary>
    ///     Applies the style to the specified ranges.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the ranges.</param>
    /// <param name="range">A single range string.</param>
    /// <param name="rangesJson">A JSON array of range strings.</param>
    /// <param name="style">The style to apply.</param>
    /// <param name="styleFlag">The style flag specifying which properties to apply.</param>
    /// <exception cref="ArgumentException">Thrown when neither range nor ranges is provided.</exception>
    private static void ApplyStyleToRanges(Worksheet worksheet, string? range, string? rangesJson,
        Aspose.Cells.Style style, StyleFlag styleFlag)
    {
        if (!string.IsNullOrEmpty(rangesJson))
        {
            var rangesList = JsonSerializer.Deserialize<List<string>>(rangesJson);
            if (rangesList != null)
                foreach (var rangeStr in rangesList.Where(r => !string.IsNullOrEmpty(r)))
                    ExcelHelper.CreateRange(worksheet.Cells, rangeStr).ApplyStyle(style, styleFlag);
            return;
        }

        if (!string.IsNullOrEmpty(range))
        {
            ExcelHelper.CreateRange(worksheet.Cells, range).ApplyStyle(style, styleFlag);
            return;
        }

        throw new ArgumentException("Either range or ranges must be provided for format operation");
    }

    /// <summary>
    ///     Applies background settings to the style including pattern and colors.
    /// </summary>
    /// <param name="style">The style to apply background settings to.</param>
    /// <param name="backgroundColor">The background color string.</param>
    /// <param name="patternType">The pattern type string.</param>
    /// <param name="patternColor">The pattern color string.</param>
    private static void ApplyBackgroundSettings(Aspose.Cells.Style style, string? backgroundColor, string? patternType,
        string? patternColor)
    {
        var bgPattern = BackgroundType.Solid;
        if (!string.IsNullOrWhiteSpace(patternType))
            bgPattern = patternType.ToLower() switch
            {
                "solid" => BackgroundType.Solid,
                "gray50" => BackgroundType.Gray50,
                "gray75" => BackgroundType.Gray75,
                "gray25" => BackgroundType.Gray25,
                "horizontalstripe" => BackgroundType.HorizontalStripe,
                "verticalstripe" => BackgroundType.VerticalStripe,
                "diagonalstripe" => BackgroundType.DiagonalStripe,
                "reversediagonalstripe" => BackgroundType.ReverseDiagonalStripe,
                "diagonalcrosshatch" => BackgroundType.DiagonalCrosshatch,
                "thickdiagonalcrosshatch" => BackgroundType.ThickDiagonalCrosshatch,
                "thinhorizontalstripe" => BackgroundType.ThinHorizontalStripe,
                "thinverticalstripe" => BackgroundType.ThinVerticalStripe,
                "thinreversediagonalstripe" => BackgroundType.ThinReverseDiagonalStripe,
                "thindiagonalstripe" => BackgroundType.ThinDiagonalStripe,
                "thinhorizontalcrosshatch" => BackgroundType.ThinHorizontalCrosshatch,
                "thindiagonalcrosshatch" => BackgroundType.ThinDiagonalCrosshatch,
                _ => BackgroundType.Solid
            };

        style.Pattern = bgPattern;

        if (!string.IsNullOrWhiteSpace(backgroundColor))
            style.ForegroundColor = ColorHelper.ParseColor(backgroundColor, true);

        if (!string.IsNullOrWhiteSpace(patternColor))
            style.BackgroundColor = ColorHelper.ParseColor(patternColor, true);
    }

    /// <summary>
    ///     Parses a horizontal alignment string to TextAlignmentType.
    /// </summary>
    /// <param name="alignment">The alignment string (left, center, right).</param>
    /// <returns>The corresponding TextAlignmentType value.</returns>
    private static TextAlignmentType ParseHorizontalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "left" => TextAlignmentType.Left,
            "center" => TextAlignmentType.Center,
            "right" => TextAlignmentType.Right,
            _ => TextAlignmentType.Left
        };
    }

    /// <summary>
    ///     Parses a vertical alignment string to TextAlignmentType.
    /// </summary>
    /// <param name="alignment">The alignment string (top, center, bottom).</param>
    /// <returns>The corresponding TextAlignmentType value.</returns>
    private static TextAlignmentType ParseVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => TextAlignmentType.Top,
            "center" => TextAlignmentType.Center,
            "bottom" => TextAlignmentType.Bottom,
            _ => TextAlignmentType.Center
        };
    }

    /// <summary>
    ///     Applies border settings to all four sides of the style.
    /// </summary>
    /// <param name="style">The style to apply border to.</param>
    /// <param name="borderStyle">The border style string (none, thin, medium, thick, dotted, dashed, double).</param>
    /// <param name="borderColor">The border color string.</param>
    private static void ApplyBorderSettings(Aspose.Cells.Style style, string borderStyle, string? borderColor)
    {
        var borderType = borderStyle.ToLower() switch
        {
            "none" => CellBorderType.None,
            "thin" => CellBorderType.Thin,
            "medium" => CellBorderType.Medium,
            "thick" => CellBorderType.Thick,
            "dotted" => CellBorderType.Dotted,
            "dashed" => CellBorderType.Dashed,
            "double" => CellBorderType.Double,
            _ => CellBorderType.Thin
        };

        var borderColorValue = Color.Black;
        if (!string.IsNullOrWhiteSpace(borderColor))
            borderColorValue = ColorHelper.ParseColor(borderColor, true);

        style.SetBorder(BorderType.TopBorder, borderType, borderColorValue);
        style.SetBorder(BorderType.BottomBorder, borderType, borderColorValue);
        style.SetBorder(BorderType.LeftBorder, borderType, borderColorValue);
        style.SetBorder(BorderType.RightBorder, borderType, borderColorValue);
    }

    /// <summary>
    ///     Record containing all format parameters for cell formatting operations.
    /// </summary>
    private sealed record FormatParameters(
        int SheetIndex,
        string? Range,
        string? RangesJson,
        string? FontName,
        int? FontSize,
        bool? Bold,
        bool? Italic,
        string? FontColor,
        string? BackgroundColor,
        string? PatternType,
        string? PatternColor,
        string? NumberFormat,
        string? BorderStyle,
        string? BorderColor,
        string? HorizontalAlignment,
        string? VerticalAlignment);
}
