using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Style;

namespace AsposeMcpServer.Handlers.Excel.Style;

/// <summary>
///     Handler for getting cell format information from Excel workbooks.
/// </summary>
[ResultType(typeof(GetCellFormatResult))]
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
    /// <returns>A GetCellFormatResult containing format information.</returns>
    /// <exception cref="ArgumentException">Thrown when neither cell nor range is provided, or when the cell range is invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
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
            return new GetCellFormatResult
            {
                Count = cellList.Count,
                WorksheetName = worksheet.Name,
                Range = cellOrRange,
                Fields = p.Fields ?? "all",
                Items = cellList
            };
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
    /// <returns>A list of CellFormatInfo objects containing cell format data.</returns>
    private static List<CellFormatInfo> CollectCellFormats(Worksheet worksheet, string cellOrRange,
        HashSet<string> requestedFields)
    {
        var cells = worksheet.Cells;
        var cellRange = ExcelHelper.CreateRange(cells, cellOrRange);
        var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
        var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

        List<CellFormatInfo> cellList = [];
        for (var row = cellRange.FirstRow; row <= endRow; row++)
        for (var col = cellRange.FirstColumn; col <= endCol; col++)
            cellList.Add(BuildCellData(cells[row, col], row, col, requestedFields));

        return cellList;
    }

    /// <summary>
    ///     Builds cell data for a single cell.
    /// </summary>
    /// <param name="cellObj">The cell object.</param>
    /// <param name="row">The row index.</param>
    /// <param name="col">The column index.</param>
    /// <param name="requestedFields">The set of fields to include.</param>
    /// <returns>A CellFormatInfo object containing the cell data.</returns>
    private static CellFormatInfo BuildCellData(Aspose.Cells.Cell cellObj, int row, int col,
        HashSet<string> requestedFields)
    {
        var style = cellObj.GetStyle();
        var cellRef = CellsHelper.CellIndexToName(row, col);

        string? value = null;
        string? formula = null;
        string? dataType = null;

        if (ShouldInclude(requestedFields, "value"))
        {
            value = cellObj.Value?.ToString() ?? "(empty)";
            formula = cellObj.Formula;
            dataType = cellObj.Type.ToString();
        }

        var formatDetails = BuildFormatData(style, requestedFields);

        return new CellFormatInfo
        {
            Cell = cellRef,
            Value = value,
            Formula = formula,
            DataType = dataType,
            Format = formatDetails
        };
    }

    /// <summary>
    ///     Builds format data based on requested fields.
    /// </summary>
    /// <param name="style">The cell style.</param>
    /// <param name="requestedFields">The set of fields to include.</param>
    /// <returns>A CellFormatDetails object containing the format data, or null if no format fields requested.</returns>
    private static CellFormatDetails? BuildFormatData(Aspose.Cells.Style style,
        HashSet<string> requestedFields)
    {
        var hasFont = ShouldInclude(requestedFields, "font");
        var hasColor = ShouldInclude(requestedFields, "color");
        var hasAlignment = ShouldInclude(requestedFields, "alignment");
        var hasNumber = ShouldInclude(requestedFields, "number");
        var hasBorder = ShouldInclude(requestedFields, "border");

        if (!hasFont && !hasColor && !hasAlignment && !hasNumber && !hasBorder)
            return null;

        return new CellFormatDetails
        {
            FontName = hasFont ? style.Font.Name : null,
            FontSize = hasFont ? style.Font.Size : null,
            Bold = hasFont ? style.Font.IsBold : null,
            Italic = hasFont ? style.Font.IsItalic : null,
            Underline = hasFont ? style.Font.Underline.ToString() : null,
            Strikethrough = hasFont ? style.Font.IsStrikeout : null,
            FontColor = hasColor ? style.Font.Color.ToString() : null,
            ForegroundColor = hasColor ? style.ForegroundColor.ToString() : null,
            BackgroundColor = hasColor ? style.BackgroundColor.ToString() : null,
            PatternType = hasColor ? style.Pattern.ToString() : null,
            HorizontalAlignment = hasAlignment ? style.HorizontalAlignment.ToString() : null,
            VerticalAlignment = hasAlignment ? style.VerticalAlignment.ToString() : null,
            NumberFormat = hasNumber ? style.Number : null,
            CustomFormat = hasNumber ? style.Custom : null,
            Borders = hasBorder ? BuildBordersInfo(style) : null
        };
    }

    /// <summary>
    ///     Builds borders information from the cell style.
    /// </summary>
    /// <param name="style">The cell style.</param>
    /// <returns>A BordersInfo object containing all border information.</returns>
    private static BordersInfo BuildBordersInfo(Aspose.Cells.Style style)
    {
        return new BordersInfo
        {
            Top = BuildBorderInfo(style.Borders[BorderType.TopBorder]),
            Bottom = BuildBorderInfo(style.Borders[BorderType.BottomBorder]),
            Left = BuildBorderInfo(style.Borders[BorderType.LeftBorder]),
            Right = BuildBorderInfo(style.Borders[BorderType.RightBorder])
        };
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
    ///     Builds border information from a border.
    /// </summary>
    /// <param name="border">The border to get information from.</param>
    /// <returns>A BorderInfo object containing border line style and color.</returns>
    private static BorderInfo BuildBorderInfo(Border border)
    {
        return new BorderInfo
        {
            LineStyle = border.LineStyle.ToString(),
            Color = border.Color.ToString()
        };
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
