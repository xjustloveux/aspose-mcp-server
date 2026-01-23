using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Cell;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for getting cell values in Excel workbooks.
/// </summary>
[ResultType(typeof(GetCellResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        ExcelCellHelper.ValidateCellAddress(getParams.Cell);

        var workbook = context.Document;

        if (getParams.CalculateFormula)
            workbook.CalculateFormula();

        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var cellObj = worksheet.Cells[getParams.Cell];

        var formula = getParams.IncludeFormula && !string.IsNullOrEmpty(cellObj.Formula) ? cellObj.Formula : null;
        GetCellFormatInfo? formatInfo = null;

        if (getParams.IncludeFormat)
        {
            var style = cellObj.GetStyle();
            formatInfo = new GetCellFormatInfo
            {
                FontName = style.Font.Name,
                FontSize = style.Font.Size,
                Bold = style.Font.IsBold,
                Italic = style.Font.IsItalic,
                BackgroundColor = style.ForegroundColor.ToString(),
                NumberFormat = style.Number
            };
        }

        return new GetCellResult
        {
            Cell = getParams.Cell,
            Value = cellObj.Value?.ToString() ?? "(empty)",
            ValueType = cellObj.Type.ToString(),
            Formula = formula,
            Format = formatInfo
        };
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
    private sealed record GetParameters(
        string Cell,
        int SheetIndex,
        bool CalculateFormula,
        bool IncludeFormula,
        bool IncludeFormat);
}
