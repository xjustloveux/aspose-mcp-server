using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for adding a new pivot table to the worksheet.
/// </summary>
public class AddExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new pivot table to the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sourceRange, destCell
    ///     Optional: sheetIndex, name
    /// </param>
    /// <returns>Success message with add details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractAddPivotTableParameters(parameters);

        var pivotName = p.Name ?? "PivotTable1";
        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

        var pivotTables = worksheet.PivotTables;
        var pivotIndex = pivotTables.Add($"={worksheet.Name}!{p.SourceRange}", p.DestCell, pivotName);
        var pivotTable = pivotTables[pivotIndex];

        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        pivotTable.CalculateData();

        MarkModified(context);

        return Success($"Pivot table '{pivotName}' added to worksheet.");
    }

    private static AddPivotTableParameters ExtractAddPivotTableParameters(OperationParameters parameters)
    {
        return new AddPivotTableParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("sourceRange"),
            parameters.GetRequired<string>("destCell"),
            parameters.GetOptional<string?>("name")
        );
    }

    private sealed record AddPivotTableParameters(int SheetIndex, string SourceRange, string DestCell, string? Name);
}
