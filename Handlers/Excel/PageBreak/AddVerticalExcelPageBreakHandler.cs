using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PageBreak;

/// <summary>
///     Handler for adding a vertical page break to an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddVerticalExcelPageBreakHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add_vertical";

    /// <summary>
    ///     Adds a vertical page break at the specified column.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: column
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var column = parameters.GetOptional<int?>("column");

        if (!column.HasValue)
            throw new ArgumentException("column is required for add_vertical operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            worksheet.VerticalPageBreaks.Add(column.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message = $"Vertical page break added at column {column.Value} in sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add vertical page break: {ex.Message}");
        }
    }
}
