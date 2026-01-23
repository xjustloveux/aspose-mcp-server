using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.PivotTable;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for getting information about all pivot tables.
/// </summary>
[ResultType(typeof(GetPivotTablesResult))]
public class GetExcelPivotTablesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets information about all pivot tables in the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>JSON result with pivot table information.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetPivotTablesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTables.Count == 0)
            return new GetPivotTablesResult
            {
                Count = 0,
                WorksheetName = worksheet.Name,
                Items = Array.Empty<PivotTableInfo>(),
                Message = "No pivot tables found"
            };

        var pivotTableList = BuildPivotTableList(pivotTables);

        return new GetPivotTablesResult
        {
            Count = pivotTables.Count,
            WorksheetName = worksheet.Name,
            Items = pivotTableList
        };
    }

    /// <summary>
    ///     Builds a list of pivot table information objects.
    /// </summary>
    /// <param name="pivotTables">The pivot tables collection.</param>
    /// <returns>List of pivot table information objects.</returns>
    private static List<PivotTableInfo> BuildPivotTableList(PivotTableCollection pivotTables)
    {
        List<PivotTableInfo> pivotTableList = [];
        for (var i = 0; i < pivotTables.Count; i++)
            pivotTableList.Add(BuildPivotTableInfo(pivotTables[i], i));
        return pivotTableList;
    }

    /// <summary>
    ///     Builds information object for a single pivot table.
    /// </summary>
    /// <param name="pivotTable">The pivot table.</param>
    /// <param name="index">The pivot table index.</param>
    /// <returns>Pivot table information object.</returns>
    private static PivotTableInfo BuildPivotTableInfo(Aspose.Cells.Pivot.PivotTable pivotTable, int index)
    {
        return new PivotTableInfo
        {
            Index = index,
            Name = pivotTable.Name ?? "(no name)",
            DataSource = GetDataSourceInfo(pivotTable),
            Location = GetLocationInfo(pivotTable.DataBodyRange),
            RowFields = GetRowFieldsList(pivotTable.RowFields),
            ColumnFields = GetColumnFieldsList(pivotTable.ColumnFields),
            DataFields = GetDataFieldsList(pivotTable.DataFields)
        };
    }

    /// <summary>
    ///     Gets the data source information string.
    /// </summary>
    /// <param name="pivotTable">The pivot table.</param>
    /// <returns>Data source information string.</returns>
    private static string GetDataSourceInfo(Aspose.Cells.Pivot.PivotTable pivotTable)
    {
        if (pivotTable.DataSource is Array { Length: > 0 } dataSourceArray)
            return string.Join(", ", dataSourceArray.Cast<object?>()
                .Where(item => item != null)
                .Select(item => item!.ToString() ?? ""));

        return pivotTable.DataSource?.ToString() ?? "Unknown";
    }

    /// <summary>
    ///     Gets the location information object.
    /// </summary>
    /// <param name="dataBodyRange">The data body range.</param>
    /// <returns>Location information object.</returns>
    private static PivotTableLocation GetLocationInfo(CellArea dataBodyRange)
    {
        if (dataBodyRange.StartRow < 0)
            return new PivotTableLocation
            {
                Range = "Not calculated",
                StartRow = -1,
                EndRow = -1,
                StartColumn = -1,
                EndColumn = -1
            };

        var startCell = CellsHelper.CellIndexToName(dataBodyRange.StartRow, dataBodyRange.StartColumn);
        var endCell = CellsHelper.CellIndexToName(dataBodyRange.EndRow, dataBodyRange.EndColumn);
        return new PivotTableLocation
        {
            Range = $"{startCell}:{endCell}",
            StartRow = dataBodyRange.StartRow,
            EndRow = dataBodyRange.EndRow,
            StartColumn = dataBodyRange.StartColumn,
            EndColumn = dataBodyRange.EndColumn
        };
    }

    /// <summary>
    ///     Gets the row fields list.
    /// </summary>
    /// <param name="rowFields">The row fields collection.</param>
    /// <returns>List of row field objects.</returns>
    private static List<PivotFieldInfo> GetRowFieldsList(PivotFieldCollection? rowFields)
    {
        List<PivotFieldInfo> list = [];
        if (rowFields is not { Count: > 0 }) return list;

        foreach (PivotField field in rowFields)
            list.Add(new PivotFieldInfo { Name = field.Name ?? $"Field {field.Position}", Position = field.Position });
        return list;
    }

    /// <summary>
    ///     Gets the column fields list.
    /// </summary>
    /// <param name="columnFields">The column fields collection.</param>
    /// <returns>List of column field objects.</returns>
    private static List<PivotFieldInfo> GetColumnFieldsList(PivotFieldCollection? columnFields)
    {
        List<PivotFieldInfo> list = [];
        if (columnFields is not { Count: > 0 }) return list;

        foreach (PivotField field in columnFields)
            list.Add(new PivotFieldInfo { Name = field.Name ?? $"Field {field.Position}", Position = field.Position });
        return list;
    }

    /// <summary>
    ///     Gets the data fields list.
    /// </summary>
    /// <param name="dataFields">The data fields collection.</param>
    /// <returns>List of data field objects.</returns>
    private static List<PivotDataFieldInfo> GetDataFieldsList(PivotFieldCollection? dataFields)
    {
        List<PivotDataFieldInfo> list = [];
        if (dataFields is not { Count: > 0 }) return list;

        foreach (PivotField field in dataFields)
            list.Add(new PivotDataFieldInfo
            {
                Name = field.Name ?? $"Field {field.Position}",
                Position = field.Position,
                Function = field.Function.ToString()
            });
        return list;
    }

    private static GetPivotTablesParameters ExtractGetPivotTablesParameters(OperationParameters parameters)
    {
        return new GetPivotTablesParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    private sealed record GetPivotTablesParameters(int SheetIndex);
}
