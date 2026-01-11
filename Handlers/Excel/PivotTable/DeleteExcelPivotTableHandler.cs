using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for deleting a pivot table from the worksheet.
/// </summary>
public class DeleteExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a pivot table from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: pivotTableIndex
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>Success message with delete details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var pivotTableIndex = parameters.GetRequired<int>("pivotTableIndex");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex];
        var pivotTableName = pivotTable.Name ?? $"PivotTable {pivotTableIndex}";

        pivotTables.RemoveAt(pivotTableIndex);

        MarkModified(context);

        return Success($"Pivot table #{pivotTableIndex} ({pivotTableName}) deleted.");
    }
}
