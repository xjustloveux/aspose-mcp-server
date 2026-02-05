using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for converting a table (ListObject) to a normal range in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ConvertToRangeExcelTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "convert_to_range";

    /// <summary>
    ///     Converts a table to a normal cell range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: tableIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");

        if (!tableIndex.HasValue)
            throw new ArgumentException("tableIndex is required for convert_to_range operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            GetExcelTablesHandler.ValidateTableIndex(worksheet, tableIndex.Value);

            var listObject = worksheet.ListObjects[tableIndex.Value];
            var tableName = listObject.DisplayName ?? $"Table{tableIndex.Value + 1}";

            listObject.ConvertToRange();

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Table '{tableName}' (index {tableIndex.Value}) converted to a normal range in sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to convert table to range: {ex.Message}");
        }
    }
}
