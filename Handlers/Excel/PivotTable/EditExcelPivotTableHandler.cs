using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for editing an existing pivot table.
/// </summary>
public class EditExcelPivotTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing pivot table (name, style, layout, refresh data).
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: pivotTableIndex
    ///     Optional: sheetIndex, name, style, showRowGrand, showColumnGrand, autoFitColumns, refreshData
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var pivotTableIndex = parameters.GetRequired<int>("pivotTableIndex");
        var name = parameters.GetOptional<string?>("name");
        var style = parameters.GetOptional<string?>("style");
        var showRowGrand = parameters.GetOptional<bool?>("showRowGrand");
        var showColumnGrand = parameters.GetOptional<bool?>("showColumnGrand");
        var autoFitColumns = parameters.GetOptional("autoFitColumns", false);
        var refreshData = parameters.GetOptional("refreshData", false);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(name))
        {
            pivotTable.Name = name;
            changes.Add($"name={name}");
        }

        if (!string.IsNullOrEmpty(style))
        {
            var styleType = ExcelPivotTableHelper.ParsePivotTableStyle(style);
            pivotTable.PivotTableStyleType = styleType;
            changes.Add($"style={style}");
        }

        if (showRowGrand.HasValue)
        {
            pivotTable.RowGrand = showRowGrand.Value;
            changes.Add($"showRowGrand={showRowGrand.Value}");
        }

        if (showColumnGrand.HasValue)
        {
            pivotTable.ColumnGrand = showColumnGrand.Value;
            changes.Add($"showColumnGrand={showColumnGrand.Value}");
        }

        if (refreshData)
        {
            pivotTable.CalculateData();
            changes.Add("refreshed");
        }

        if (autoFitColumns)
        {
            worksheet.AutoFitColumns();
            changes.Add("autoFitColumns");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Pivot table #{pivotTableIndex} edited ({changesStr}).");
    }
}
