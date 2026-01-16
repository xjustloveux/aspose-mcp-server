using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Handler for adding formulas to Excel cells.
/// </summary>
public class AddFormulaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a formula to a cell in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell, formula
    ///     Optional: sheetIndex (default: 0), autoCalculate (default: true)
    /// </param>
    /// <returns>Success message with formula details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);

        var cellObj = worksheet.Cells[addParams.Cell];
        cellObj.Formula = addParams.Formula;

        string? warningMessage = null;
        if (addParams.AutoCalculate)
        {
            workbook.CalculateFormula();
            _ = cellObj.Value;

            if (cellObj.Type == CellValueType.IsError)
            {
                var errorValue = cellObj.DisplayStringValue;
                if (!string.IsNullOrEmpty(errorValue) && errorValue.StartsWith("#"))
                {
                    warningMessage = $" Warning: {errorValue}";
                    warningMessage += errorValue switch
                    {
                        "#NAME?" => " (invalid function name)",
                        "#VALUE?" => " (incorrect argument type)",
                        "#REF!" => " (invalid cell reference)",
                        _ => ""
                    };
                }
            }
        }

        MarkModified(context);

        var result = $"Formula added to {addParams.Cell}: {addParams.Formula}";
        if (!string.IsNullOrEmpty(warningMessage)) result += $".{warningMessage}";
        return result;
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetRequired<string>("formula"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("autoCalculate", true)
        );
    }

    /// <summary>
    ///     Record to hold add formula parameters.
    /// </summary>
    private record AddParameters(string Cell, string Formula, int SheetIndex, bool AutoCalculate);
}
