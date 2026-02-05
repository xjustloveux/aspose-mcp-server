using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PageBreak;

/// <summary>
///     Handler for adding a horizontal page break to an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddHorizontalExcelPageBreakHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add_horizontal";

    /// <summary>
    ///     Adds a horizontal page break at the specified row.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: row
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var row = parameters.GetOptional<int?>("row");

        if (!row.HasValue)
            throw new ArgumentException("row is required for add_horizontal operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            worksheet.HorizontalPageBreaks.Add(row.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message = $"Horizontal page break added at row {row.Value} in sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add horizontal page break: {ex.Message}");
        }
    }
}
