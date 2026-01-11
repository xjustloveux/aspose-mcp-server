using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for getting cell values in Excel workbooks.
/// </summary>
public class GetExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets the value and properties of the specified cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell (cell reference like "A1")
    ///     Optional: sheetIndex (default: 0), calculateFormula (default: false),
    ///     includeFormula (default: true), includeFormat (default: false)
    /// </param>
    /// <returns>JSON result with cell information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var cell = parameters.GetRequired<string>("cell");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var calculateFormula = parameters.GetOptional("calculateFormula", false);
        var includeFormula = parameters.GetOptional("includeFormula", true);
        var includeFormat = parameters.GetOptional("includeFormat", false);

        ExcelCellHelper.ValidateCellAddress(cell);

        var workbook = context.Document;

        if (calculateFormula)
            workbook.CalculateFormula();

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];

        object resultObj;

        if (includeFormat)
        {
            var style = cellObj.GetStyle();
            resultObj = new
            {
                cell,
                value = cellObj.Value?.ToString() ?? "(empty)",
                valueType = cellObj.Type.ToString(),
                formula = includeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null,
                format = new
                {
                    fontName = style.Font.Name,
                    fontSize = style.Font.Size,
                    bold = style.Font.IsBold,
                    italic = style.Font.IsItalic,
                    backgroundColor = style.ForegroundColor.ToString(),
                    numberFormat = style.Number
                }
            };
        }
        else
        {
            resultObj = new
            {
                cell,
                value = cellObj.Value?.ToString() ?? "(empty)",
                valueType = cellObj.Type.ToString(),
                formula = includeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null
            };
        }

        return JsonResult(resultObj);
    }
}
