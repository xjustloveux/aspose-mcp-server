using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for refreshing pivot table data.
/// </summary>
public class RefreshExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "refresh";

    /// <summary>
    ///     Refreshes pivot table data (one or all tables).
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, pivotTableIndex (if not provided, refreshes all)
    /// </param>
    /// <returns>Success message with refresh details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var pivotTableIndex = parameters.GetOptional<int?>("pivotTableIndex");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTables.Count == 0)
            throw new InvalidOperationException($"No pivot tables found in worksheet '{worksheet.Name}'");

        var refreshedCount = 0;

        if (pivotTableIndex.HasValue)
        {
            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            pivotTables[pivotTableIndex.Value].CalculateData();
            refreshedCount = 1;
        }
        else
        {
            foreach (var pivotTable in pivotTables)
            {
                pivotTable.CalculateData();
                refreshedCount++;
            }
        }

        MarkModified(context);

        return Success($"Refreshed {refreshedCount} pivot table(s) in worksheet '{worksheet.Name}'.");
    }
}
