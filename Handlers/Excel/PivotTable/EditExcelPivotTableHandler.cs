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
        var p = ExtractEditPivotTableParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (p.PivotTableIndex < 0 || p.PivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {p.PivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[p.PivotTableIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.Name))
        {
            pivotTable.Name = p.Name;
            changes.Add($"name={p.Name}");
        }

        if (!string.IsNullOrEmpty(p.Style))
        {
            var styleType = ExcelPivotTableHelper.ParsePivotTableStyle(p.Style);
            pivotTable.PivotTableStyleType = styleType;
            changes.Add($"style={p.Style}");
        }

        if (p.ShowRowGrand.HasValue)
        {
            pivotTable.RowGrand = p.ShowRowGrand.Value;
            changes.Add($"showRowGrand={p.ShowRowGrand.Value}");
        }

        if (p.ShowColumnGrand.HasValue)
        {
            pivotTable.ColumnGrand = p.ShowColumnGrand.Value;
            changes.Add($"showColumnGrand={p.ShowColumnGrand.Value}");
        }

        if (p.RefreshData)
        {
            pivotTable.CalculateData();
            changes.Add("refreshed");
        }

        if (p.AutoFitColumns)
        {
            worksheet.AutoFitColumns();
            changes.Add("autoFitColumns");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return Success($"Pivot table #{p.PivotTableIndex} edited ({changesStr}).");
    }

    private static EditPivotTableParameters ExtractEditPivotTableParameters(OperationParameters parameters)
    {
        return new EditPivotTableParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("pivotTableIndex"),
            parameters.GetOptional<string?>("name"),
            parameters.GetOptional<string?>("style"),
            parameters.GetOptional<bool?>("showRowGrand"),
            parameters.GetOptional<bool?>("showColumnGrand"),
            parameters.GetOptional("autoFitColumns", false),
            parameters.GetOptional("refreshData", false)
        );
    }

    private record EditPivotTableParameters(
        int SheetIndex,
        int PivotTableIndex,
        string? Name,
        string? Style,
        bool? ShowRowGrand,
        bool? ShowColumnGrand,
        bool AutoFitColumns,
        bool RefreshData);
}
