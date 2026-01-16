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
        var p = ExtractDeletePivotTableParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (p.PivotTableIndex < 0 || p.PivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {p.PivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[p.PivotTableIndex];
        var pivotTableName = pivotTable.Name ?? $"PivotTable {p.PivotTableIndex}";

        pivotTables.RemoveAt(p.PivotTableIndex);

        MarkModified(context);

        return Success($"Pivot table #{p.PivotTableIndex} ({pivotTableName}) deleted.");
    }

    private static DeletePivotTableParameters ExtractDeletePivotTableParameters(OperationParameters parameters)
    {
        return new DeletePivotTableParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex")
        );
    }

    private record DeletePivotTableParameters(int SheetIndex, int PivotTableIndex);
}
