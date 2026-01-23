using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Handler for writing values to Excel cells.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class WriteExcelCellHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "write";

    /// <summary>
    ///     Writes a value to the specified cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell, value
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when value is empty.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var writeParams = ExtractWriteParameters(parameters);

        if (string.IsNullOrEmpty(writeParams.Value))
            throw new ArgumentException("value is required for write operation");

        ExcelCellHelper.ValidateCellAddress(writeParams.Cell);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, writeParams.SheetIndex);
        var cellObj = worksheet.Cells[writeParams.Cell];

        ExcelHelper.SetCellValue(cellObj, writeParams.Value);

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Cell {writeParams.Cell} written with value '{writeParams.Value}' in sheet {writeParams.SheetIndex}."
        };
    }

    /// <summary>
    ///     Extracts write parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted write parameters.</returns>
    private static WriteParameters ExtractWriteParameters(OperationParameters parameters)
    {
        return new WriteParameters(
            parameters.GetRequired<string>("cell"),
            parameters.GetRequired<string>("value"),
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Record to hold write cell parameters.
    /// </summary>
    private sealed record WriteParameters(string Cell, string Value, int SheetIndex);
}
