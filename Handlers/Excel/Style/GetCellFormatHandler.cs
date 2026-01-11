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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var range = parameters.GetOptional<string?>("range");
        var fieldsParam = parameters.GetOptional<string?>("fields");

        if (string.IsNullOrEmpty(cell) && string.IsNullOrEmpty(range))
            throw new ArgumentException("Either cell or range is required for get_format operation");

        var cellOrRange = cell ?? range!;
        var requestedFields = ParseFields(fieldsParam);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        try
        {
            var cellRange = ExcelHelper.CreateRange(cells, cellOrRange);
            var startRow = cellRange.FirstRow;
            var endRow = cellRange.FirstRow + cellRange.RowCount - 1;
            var startCol = cellRange.FirstColumn;
            var endCol = cellRange.FirstColumn + cellRange.ColumnCount - 1;

            List<Dictionary<string, object?>> cellList = [];
            for (var row = startRow; row <= endRow; row++)
            for (var col = startCol; col <= endCol; col++)
            {
                var cellObj = cells[row, col];
                var style = cellObj.GetStyle();

                var cellData = new Dictionary<string, object?>
                {
                    ["cell"] = CellsHelper.CellIndexToName(row, col)
                };

                if (requestedFields.Contains("value") || requestedFields.Contains("all"))
                {
                    cellData["value"] = cellObj.Value?.ToString() ?? "(empty)";
                    cellData["formula"] = cellObj.Formula;
                    cellData["dataType"] = cellObj.Type.ToString();
                }

                var formatData = new Dictionary<string, object?>();

                if (requestedFields.Contains("font") || requestedFields.Contains("all"))
                {
                    formatData["fontName"] = style.Font.Name;
                    formatData["fontSize"] = style.Font.Size;
                    formatData["bold"] = style.Font.IsBold;
                    formatData["italic"] = style.Font.IsItalic;
                    formatData["underline"] = style.Font.Underline.ToString();
                    formatData["strikethrough"] = style.Font.IsStrikeout;
                }

                if (requestedFields.Contains("color") || requestedFields.Contains("all"))
                {
                    formatData["fontColor"] = style.Font.Color.ToString();
                    formatData["foregroundColor"] = style.ForegroundColor.ToString();
                    formatData["backgroundColor"] = style.BackgroundColor.ToString();
                    formatData["patternType"] = style.Pattern.ToString();
                }

                if (requestedFields.Contains("alignment") || requestedFields.Contains("all"))
                {
                    formatData["horizontalAlignment"] = style.HorizontalAlignment.ToString();
                    formatData["verticalAlignment"] = style.VerticalAlignment.ToString();
                }

                if (requestedFields.Contains("number") || requestedFields.Contains("all"))
                {
                    formatData["numberFormat"] = style.Number;
                    formatData["customFormat"] = style.Custom;
                }

                if (requestedFields.Contains("border") || requestedFields.Contains("all"))
                {
                    var topBorder = style.Borders[BorderType.TopBorder];
                    var bottomBorder = style.Borders[BorderType.BottomBorder];
                    var leftBorder = style.Borders[BorderType.LeftBorder];
                    var rightBorder = style.Borders[BorderType.RightBorder];

                    formatData["borders"] = new
                    {
                        top = new { lineStyle = topBorder.LineStyle.ToString(), color = topBorder.Color.ToString() },
                        bottom = new
                            { lineStyle = bottomBorder.LineStyle.ToString(), color = bottomBorder.Color.ToString() },
                        left = new { lineStyle = leftBorder.LineStyle.ToString(), color = leftBorder.Color.ToString() },
                        right = new
                        {
                            lineStyle = rightBorder.LineStyle.ToString(), color = rightBorder.Color.ToString()
                        }
                    };
                }

                if (formatData.Count > 0)
                    cellData["format"] = formatData;

                cellList.Add(cellData);
            }

            var result = new
            {
                count = cellList.Count,
                worksheetName = worksheet.Name,
                range = cellOrRange,
                fields = fieldsParam ?? "all",
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
}
