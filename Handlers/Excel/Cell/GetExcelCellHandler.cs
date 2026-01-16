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
    ///     Required: cell
    ///     Optional: sheetIndex, calculateFormula, includeFormula, includeFormat
    /// </param>
    /// <returns>JSON result with cell information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        ExcelCellHelper.ValidateCellAddress(getParams.Cell);

        var workbook = context.Document;

        if (getParams.CalculateFormula)
            workbook.CalculateFormula();

        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var cellObj = worksheet.Cells[getParams.Cell];

        object resultObj;

        if (getParams.IncludeFormat)
        {
            var style = cellObj.GetStyle();
            resultObj = new
            {
                cell = getParams.Cell,
                value = cellObj.Value?.ToString() ?? "(empty)",
                valueType = cellObj.Type.ToString(),
                formula = getParams.IncludeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null,
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
                cell = getParams.Cell,
                value = cellObj.Value?.ToString() ?? "(empty)",
                valueType = cellObj.Type.ToString(),
                formula = getParams.IncludeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null
            };
        }

        return JsonResult(resultObj);
    }

    /// <summary>
    ///     Extracts get parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("calculateFormula", false),
            parameters.GetOptional("includeFormula", true),
            parameters.GetOptional("includeFormat", false)
        );
    }

    /// <summary>
    ///     Record to hold get cell parameters.
    /// </summary>
    private record GetParameters(
        string Cell,
        int SheetIndex,
        bool CalculateFormula,
        bool IncludeFormula,
        bool IncludeFormat);
}
