using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Handler for getting information about all pivot tables.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractGetPivotTablesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTables.Count == 0)
            return JsonResult(new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No pivot tables found"
            });

        List<object> pivotTableList = [];
        for (var i = 0; i < pivotTables.Count; i++)
        {
            var pivotTable = pivotTables[i];

            string dataSourceInfo;
            if (pivotTable.DataSource is Array { Length: > 0 } dataSourceArray)
            {
                List<string> sourceParts = [];
                foreach (var item in dataSourceArray)
                    if (item != null)
                        sourceParts.Add(item.ToString() ?? "");
                dataSourceInfo = string.Join(", ", sourceParts);
            }
            else if (pivotTable.DataSource != null)
            {
                dataSourceInfo = pivotTable.DataSource.ToString() ?? "Unknown";
            }
            else
            {
                dataSourceInfo = "Unknown";
            }

            object locationInfo;
            var dataBodyRange = pivotTable.DataBodyRange;
            if (dataBodyRange.StartRow >= 0)
            {
                var startCell = CellsHelper.CellIndexToName(dataBodyRange.StartRow, dataBodyRange.StartColumn);
                var endCell = CellsHelper.CellIndexToName(dataBodyRange.EndRow, dataBodyRange.EndColumn);
                locationInfo = new
                {
                    range = $"{startCell}:{endCell}",
                    startRow = dataBodyRange.StartRow,
                    endRow = dataBodyRange.EndRow,
                    startColumn = dataBodyRange.StartColumn,
                    endColumn = dataBodyRange.EndColumn
                };
            }
            else
            {
                locationInfo = new
                {
                    range = "Not calculated",
                    startRow = -1,
                    endRow = -1,
                    startColumn = -1,
                    endColumn = -1
                };
            }

            List<object> rowFieldsList = [];
            if (pivotTable.RowFields is { Count: > 0 } rowFields)
                foreach (PivotField field in rowFields)
                    rowFieldsList.Add(new
                    {
                        name = field.Name ?? $"Field {field.Position}",
                        position = field.Position
                    });

            List<object> columnFieldsList = [];
            if (pivotTable.ColumnFields is { Count: > 0 } columnFields)
                foreach (PivotField field in columnFields)
                    columnFieldsList.Add(new
                    {
                        name = field.Name ?? $"Field {field.Position}",
                        position = field.Position
                    });

            List<object> dataFieldsList = [];
            if (pivotTable.DataFields is { Count: > 0 } dataFields)
                foreach (PivotField field in dataFields)
                    dataFieldsList.Add(new
                    {
                        name = field.Name ?? $"Field {field.Position}",
                        position = field.Position,
                        function = field.Function.ToString()
                    });

            pivotTableList.Add(new
            {
                index = i,
                name = pivotTable.Name ?? "(no name)",
                dataSource = dataSourceInfo,
                location = locationInfo,
                rowFields = rowFieldsList,
                columnFields = columnFieldsList,
                dataFields = dataFieldsList
            });
        }

        return JsonResult(new
        {
            count = pivotTables.Count,
            worksheetName = worksheet.Name,
            items = pivotTableList
        });
    }

    private static GetPivotTablesParameters ExtractGetPivotTablesParameters(OperationParameters parameters)
    {
        return new GetPivotTablesParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    private record GetPivotTablesParameters(int SheetIndex);
}
