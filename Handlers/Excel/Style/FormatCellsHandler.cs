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
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetOptional<string?>("range");
        var rangesJson = parameters.GetOptional<string?>("ranges");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontSize = parameters.GetOptional<int?>("fontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var fontColor = parameters.GetOptional<string?>("fontColor");
        var backgroundColor = parameters.GetOptional<string?>("backgroundColor");
        var patternType = parameters.GetOptional<string?>("patternType");
        var patternColor = parameters.GetOptional<string?>("patternColor");
        var numberFormat = parameters.GetOptional<string?>("numberFormat");
        var borderStyle = parameters.GetOptional<string?>("borderStyle");
        var borderColor = parameters.GetOptional<string?>("borderColor");
        var horizontalAlignment = parameters.GetOptional<string?>("horizontalAlignment");
        var verticalAlignment = parameters.GetOptional<string?>("verticalAlignment");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var style = workbook.CreateStyle();

        try
        {
            FontHelper.Excel.ApplyFontSettings(
                style,
                fontName,
                fontSize,
                bold,
                italic,
                fontColor
            );
        }
        catch (Exception colorEx) when (colorEx is ArgumentException && !string.IsNullOrWhiteSpace(fontColor))
        {
            throw new ArgumentException(
                $"Unable to parse font color '{fontColor}': {colorEx.Message}. Please use a valid color format (e.g., #FF0000, 255,0,0, or red)");
        }

        if (!string.IsNullOrWhiteSpace(backgroundColor) || !string.IsNullOrWhiteSpace(patternType))
            ApplyBackgroundSettings(style, backgroundColor, patternType, patternColor);

        if (!string.IsNullOrEmpty(numberFormat))
        {
            if (int.TryParse(numberFormat, out var formatNumber))
                style.Number = formatNumber;
            else
                style.Custom = numberFormat;
        }

        if (!string.IsNullOrEmpty(horizontalAlignment))
            style.HorizontalAlignment = ParseHorizontalAlignment(horizontalAlignment);

        if (!string.IsNullOrEmpty(verticalAlignment))
            style.VerticalAlignment = ParseVerticalAlignment(verticalAlignment);

        if (!string.IsNullOrEmpty(borderStyle))
            ApplyBorderSettings(style, borderStyle, borderColor);

        var styleFlag = new StyleFlag
        {
            All = true,
            Borders = !string.IsNullOrEmpty(borderStyle)
        };

        if (!string.IsNullOrEmpty(rangesJson))
        {
            var rangesList = JsonSerializer.Deserialize<List<string>>(rangesJson);
            if (rangesList != null)
                foreach (var rangeStr in rangesList)
                    if (!string.IsNullOrEmpty(rangeStr))
                    {
                        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, rangeStr);
                        cellRange.ApplyStyle(style, styleFlag);
                    }
        }
        else if (!string.IsNullOrEmpty(range))
        {
            var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
            cellRange.ApplyStyle(style, styleFlag);
        }
        else
        {
            throw new ArgumentException("Either range or ranges must be provided for format operation");
        }

        MarkModified(context);

        return Success($"Cells formatted in sheet {sheetIndex}.");
    }

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
}
