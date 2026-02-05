using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Table;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for getting tables (ListObjects) from an Excel worksheet.
/// </summary>
[ResultType(typeof(GetTablesExcelResult))]
public class GetExcelTablesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all tables from a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), tableIndex (specific table)
    /// </param>
    /// <returns>Table information result.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var tables = new List<ExcelTableInfo>();

            if (tableIndex.HasValue)
            {
                ValidateTableIndex(worksheet, tableIndex.Value);
                tables.Add(BuildTableInfo(worksheet.ListObjects[tableIndex.Value], tableIndex.Value));
            }
            else
            {
                for (var i = 0; i < worksheet.ListObjects.Count; i++)
                    tables.Add(BuildTableInfo(worksheet.ListObjects[i], i));
            }

            return new GetTablesExcelResult
            {
                Count = tables.Count,
                SheetIndex = sheetIndex,
                Items = tables,
                Message = tables.Count == 0 ? "No tables found in the worksheet." : null
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to get tables from sheet {sheetIndex}: {ex.Message}");
        }
    }

    /// <summary>
    ///     Builds table information from a ListObject.
    /// </summary>
    /// <param name="listObject">The list object.</param>
    /// <param name="index">The table index.</param>
    /// <returns>Table information record.</returns>
    internal static ExcelTableInfo BuildTableInfo(ListObject listObject, int index)
    {
        return new ExcelTableInfo
        {
            Index = index,
            Name = listObject.DisplayName ?? $"Table{index + 1}",
            Range = listObject.DataRange?.RefersTo
                    ?? $"{listObject.StartRow}:{listObject.EndRow}",
            ShowHeaderRow = listObject.ShowHeaderRow,
            ShowTotals = listObject.ShowTotals,
            StyleName = listObject.TableStyleType.ToString(),
            DataRowCount = listObject.DataRange?.RowCount ?? 0,
            ColumnCount = listObject.ListColumns.Count
        };
    }

    /// <summary>
    ///     Validates a table index.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of range.</exception>
    internal static void ValidateTableIndex(Worksheet worksheet, int tableIndex)
    {
        if (tableIndex < 0 || tableIndex >= worksheet.ListObjects.Count)
            throw new ArgumentException(
                $"Table index {tableIndex} is out of range (worksheet has {worksheet.ListObjects.Count} tables)");
    }
}
