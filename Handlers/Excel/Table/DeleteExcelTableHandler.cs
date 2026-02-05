using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for deleting a table (ListObject) from an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteExcelTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a table from the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: tableIndex
    ///     Optional: sheetIndex (default: 0), keepData (default: true)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            GetExcelTablesHandler.ValidateTableIndex(worksheet, p.TableIndex);

            var listObject = worksheet.ListObjects[p.TableIndex];
            var tableName = listObject.DisplayName ?? $"Table{p.TableIndex + 1}";

            if (p.KeepData)
                listObject.ConvertToRange();

            worksheet.ListObjects.RemoveAt(p.TableIndex);

            MarkModified(context);

            var dataMessage = p.KeepData ? " Data preserved as regular cells." : " Table data removed.";
            return new SuccessResult
            {
                Message = $"Table '{tableName}' (index {p.TableIndex}) deleted from sheet {p.SheetIndex}.{dataMessage}"
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to delete table at index {p.TableIndex}: {ex.Message}");
        }
    }

    private static DeleteParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");
        var keepData = parameters.GetOptional("keepData", true);

        if (!tableIndex.HasValue)
            throw new ArgumentException("tableIndex is required for delete operation");

        return new DeleteParameters(sheetIndex, tableIndex.Value, keepData);
    }

    private sealed record DeleteParameters(int SheetIndex, int TableIndex, bool KeepData);
}
