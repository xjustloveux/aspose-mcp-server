using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Style;

/// <summary>
///     Handler for getting cell format information from Excel workbooks.
/// </summary>
public class GetCellFormatHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_format";

    /// <summary>
    ///     Gets format information from cells in the specified range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell or range
    ///     Optional: sheetIndex, fields
    /// </param>
    /// <returns>JSON string containing format information.</returns>
    /// <exception cref="ArgumentException">Thrown when neither cell nor range is provided, or when the cell range is invalid.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetCellFormatParameters(parameters);

        if (string.IsNullOrEmpty(p.Cell) && string.IsNullOrEmpty(p.Range))
            throw new ArgumentException("Either cell or range is required for get_format operation");

        var cellOrRange = p.Cell ?? p.Range!;
        var requestedFields = ParseFields(p.Fields);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        try
        {
            var cellList = CollectCellFormats(worksheet, cellOrRange, requestedFields);
            var result = new
            {
                count = cellList.Count,
                worksheetName = worksheet.Name,
                range = cellOrRange,
                fields = p.Fields ?? "all",
                items = cellList
            };
            return JsonResult(result);
        }
        catch (Exception ex)
        {
            throw new ArgumentException(
                $"Invalid cell range: '{cellOrRange}'. Expected format: single cell (e.g., 'A1') or range (e.g., 'A1:C5'). Error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Collects format information from cells in the specified range.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the cells.</param>
    /// <param name="cellOrRange">The cell reference or range string.</param>
    /// <param name="requestedFields">The set of fields to include in the output.</param>
    /// <returns>A list of dictionaries containing cell format data.</returns>
    private static List<Dictionary<string, object?>> CollectCellFormats(Worksheet worksheet, string cellOrRange,
        HashSet<string> requestedFields)
    {
        var cells = worksheet.Cells;
        var cellRange = ExcelHelper.CreateRange(cells, cellOrRange);
        var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
        var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

        List<Dictionary<string, object?>> cellList = [];
        for (var row = cellRange.FirstRow; row <= endRow; row++)
        for (var col = cellRange.FirstColumn; col <= endCol; col++)
            cellList.Add(BuildCellData(cells[row, col], row, col, requestedFields));

        return cellList;
    }

    /// <summary>
    ///     Builds cell data dictionary for a single cell.
    /// </summary>
    /// <param name="cellObj">The cell object.</param>
    /// <param name="row">The row index.</param>
    /// <param name="col">The column index.</param>
    /// <param name="requestedFields">The set of fields to include.</param>
    /// <returns>A dictionary containing the cell data.</returns>
    private static Dictionary<string, object?> BuildCellData(Aspose.Cells.Cell cellObj, int row, int col,
        HashSet<string> requestedFields)
    {
        var style = cellObj.GetStyle();
        var cellData = new Dictionary<string, object?> { ["cell"] = CellsHelper.CellIndexToName(row, col) };

        if (ShouldInclude(requestedFields, "value"))
        {
            cellData["value"] = cellObj.Value?.ToString() ?? "(empty)";
            cellData["formula"] = cellObj.Formula;
            cellData["dataType"] = cellObj.Type.ToString();
        }

        var formatData = BuildFormatData(style, requestedFields);
        if (formatData.Count > 0)
            cellData["format"] = formatData;

        return cellData;
    }

    /// <summary>
    ///     Builds format data dictionary based on requested fields.
    /// </summary>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of fields to include.</param>
    /// <returns>A dictionary containing the format data.</returns>
    private static Dictionary<string, object?> BuildFormatData(Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        var formatData = new Dictionary<string, object?>();
        AddFontData(formatData, style, requestedFields);
        AddColorData(formatData, style, requestedFields);
        AddAlignmentData(formatData, style, requestedFields);
        AddNumberData(formatData, style, requestedFields);
        AddBorderData(formatData, style, requestedFields);
        return formatData;
    }

    /// <summary>
    ///     Determines whether a field should be included in the output.
    /// </summary>
    /// <param name="requestedFields">The set of requested fields.</param>
    /// <param name="field">The field name to check.</param>
    /// <returns>True if the field should be included; otherwise, false.</returns>
    private static bool ShouldInclude(HashSet<string> requestedFields, string field)
    {
        return requestedFields.Contains(field) || requestedFields.Contains("all");
    }

    /// <summary>
    ///     Adds font data to the format data dictionary.
    /// </summary>
    /// <param name="formatData">The format data dictionary to add to.</param>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of requested fields.</param>
    private static void AddFontData(Dictionary<string, object?> formatData, Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        if (!ShouldInclude(requestedFields, "font")) return;
        formatData["fontName"] = style.Font.Name;
        formatData["fontSize"] = style.Font.Size;
        formatData["bold"] = style.Font.IsBold;
        formatData["italic"] = style.Font.IsItalic;
        formatData["underline"] = style.Font.Underline.ToString();
        formatData["strikethrough"] = style.Font.IsStrikeout;
    }

    /// <summary>
    ///     Adds color data to the format data dictionary.
    /// </summary>
    /// <param name="formatData">The format data dictionary to add to.</param>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of requested fields.</param>
    private static void AddColorData(Dictionary<string, object?> formatData, Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        if (!ShouldInclude(requestedFields, "color")) return;
        formatData["fontColor"] = style.Font.Color.ToString();
        formatData["foregroundColor"] = style.ForegroundColor.ToString();
        formatData["backgroundColor"] = style.BackgroundColor.ToString();
        formatData["patternType"] = style.Pattern.ToString();
    }

    /// <summary>
    ///     Adds alignment data to the format data dictionary.
    /// </summary>
    /// <param name="formatData">The format data dictionary to add to.</param>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of requested fields.</param>
    private static void AddAlignmentData(Dictionary<string, object?> formatData, Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        if (!ShouldInclude(requestedFields, "alignment")) return;
        formatData["horizontalAlignment"] = style.HorizontalAlignment.ToString();
        formatData["verticalAlignment"] = style.VerticalAlignment.ToString();
    }

    /// <summary>
    ///     Adds number format data to the format data dictionary.
    /// </summary>
    /// <param name="formatData">The format data dictionary to add to.</param>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of requested fields.</param>
    private static void AddNumberData(Dictionary<string, object?> formatData, Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        if (!ShouldInclude(requestedFields, "number")) return;
        formatData["numberFormat"] = style.Number;
        formatData["customFormat"] = style.Custom;
    }

    /// <summary>
    ///     Adds border data to the format data dictionary.
    /// </summary>
    /// <param name="formatData">The format data dictionary to add to.</param>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of requested fields.</param>
    private static void AddBorderData(Dictionary<string, object?> formatData, Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        if (!ShouldInclude(requestedFields, "border")) return;
        formatData["borders"] = new
        {
            top = BuildBorderInfo(style.Borders[BorderType.TopBorder]),
            bottom = BuildBorderInfo(style.Borders[BorderType.BottomBorder]),
            left = BuildBorderInfo(style.Borders[BorderType.LeftBorder]),
            right = BuildBorderInfo(style.Borders[BorderType.RightBorder])
        };
    }

    /// <summary>
    ///     Builds border information object from a border.
    /// </summary>
    /// <param name="border">The border to get information from.</param>
    /// <returns>An anonymous object containing border line style and color.</returns>
    private static object BuildBorderInfo(Border border)
    {
        return new { lineStyle = border.LineStyle.ToString(), color = border.Color.ToString() };
    }

    /// <summary>
    ///     Parses comma-separated field names into a set.
    /// </summary>
    /// <param name="fieldsParam">Comma-separated list of field names.</param>
    /// <returns>A set of field names, or a set containing "all" if no fields are specified.</returns>
    private static HashSet<string> ParseFields(string? fieldsParam)
    {
        if (string.IsNullOrWhiteSpace(fieldsParam))
            return ["all"];

        var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fieldsParam.Split(',',
                     StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            fields.Add(field.ToLower());

        return fields.Count == 0 ? ["all"] : fields;
    }

    /// <summary>
    ///     Extracts get cell format parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A GetCellFormatParameters record containing all extracted values.</returns>
    private static GetCellFormatParameters ExtractGetCellFormatParameters(OperationParameters parameters)
    {
        return new GetCellFormatParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("cell"),
            parameters.GetOptional<string?>("range"),
            parameters.GetOptional<string?>("fields")
        );
    }

    /// <summary>
    ///     Record containing parameters for get cell format operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Cell">The cell reference.</param>
    /// <param name="Range">The range reference.</param>
    /// <param name="Fields">The fields to include in the output.</param>
    private sealed record GetCellFormatParameters(
        int SheetIndex,
        string? Cell,
        string? Range,
        string? Fields);
}
