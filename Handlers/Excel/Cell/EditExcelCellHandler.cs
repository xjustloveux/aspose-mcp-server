using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for editing Excel cells with value, formula, or clearing.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits the specified cell with a new value, formula, or clears it.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell
    ///     Optional: sheetIndex, value, formula, clearValue
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when no edit operation is specified.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var editParams = ExtractEditParameters(parameters);

        ExcelCellHelper.ValidateCellAddress(editParams.Cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
        var cellObj = worksheet.Cells[editParams.Cell];

        if (editParams.ClearValue)
            cellObj.PutValue("");
        else if (!string.IsNullOrEmpty(editParams.Formula))
            cellObj.Formula = editParams.Formula;
        else if (!string.IsNullOrEmpty(editParams.Value))
            ExcelHelper.SetCellValue(cellObj, editParams.Value);
        else
            throw new ArgumentException("Either value, formula, or clearValue must be provided");

        MarkModified(context);

        return new SuccessResult { Message = $"Cell {editParams.Cell} edited in sheet {editParams.SheetIndex}." };
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("value"),
            parameters.GetOptional<string?>("formula"),
            parameters.GetOptional("clearValue", false)
        );
    }

    /// <summary>
    ///     Record to hold edit cell parameters.
    /// </summary>
    private sealed record EditParameters(string Cell, int SheetIndex, string? Value, string? Formula, bool ClearValue);
}
