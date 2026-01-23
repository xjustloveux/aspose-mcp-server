using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for refreshing pivot table data.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractRefreshPivotTableParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTables.Count == 0)
            throw new InvalidOperationException($"No pivot tables found in worksheet '{worksheet.Name}'");

        var refreshedCount = 0;

        if (p.PivotTableIndex.HasValue)
        {
            if (p.PivotTableIndex.Value < 0 || p.PivotTableIndex.Value >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {p.PivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            pivotTables[p.PivotTableIndex.Value].CalculateData();
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

        return new SuccessResult
            { Message = $"Refreshed {refreshedCount} pivot table(s) in worksheet '{worksheet.Name}'." };
    }

    private static RefreshPivotTableParameters ExtractRefreshPivotTableParameters(OperationParameters parameters)
    {
        return new RefreshPivotTableParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<int?>("pivotTableIndex")
        );
    }

    private sealed record RefreshPivotTableParameters(int SheetIndex, int? PivotTableIndex);
}
