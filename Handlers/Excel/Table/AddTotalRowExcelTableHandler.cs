using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for adding or configuring a total row on a table (ListObject) in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTotalRowExcelTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add_total_row";

    /// <summary>
    ///     Adds or configures a total row on a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: tableIndex
    ///     Optional: sheetIndex (default: 0), columnIndex, totalFunction (sum, count, average, max, min, none)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            GetExcelTablesHandler.ValidateTableIndex(worksheet, p.TableIndex);

            var listObject = worksheet.ListObjects[p.TableIndex];
            listObject.ShowTotals = true;

            if (p.ColumnIndex.HasValue && !string.IsNullOrEmpty(p.TotalFunction))
            {
                if (p.ColumnIndex.Value < 0 || p.ColumnIndex.Value >= listObject.ListColumns.Count)
                    throw new ArgumentException(
                        $"Column index {p.ColumnIndex.Value} is out of range (table has {listObject.ListColumns.Count} columns)");

                var function = ResolveTotalsCalculation(p.TotalFunction);
                listObject.ListColumns[p.ColumnIndex.Value].TotalsCalculation = function;
            }

            MarkModified(context);

            var tableName = listObject.DisplayName ?? $"Table{p.TableIndex + 1}";
            var message = $"Total row enabled for table '{tableName}'.";
            if (p.ColumnIndex.HasValue && !string.IsNullOrEmpty(p.TotalFunction))
                message += $" Column {p.ColumnIndex.Value} set to {p.TotalFunction}.";

            return new SuccessResult { Message = message };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add total row: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves a total function string to a TotalsCalculation enum value.
    /// </summary>
    /// <param name="function">The function name (e.g., "sum", "count", "average").</param>
    /// <returns>The corresponding TotalsCalculation value.</returns>
    /// <exception cref="ArgumentException">Thrown when the function name is unknown.</exception>
    internal static TotalsCalculation ResolveTotalsCalculation(string function)
    {
        return function.ToLowerInvariant() switch
        {
            "sum" => TotalsCalculation.Sum,
            "count" => TotalsCalculation.Count,
            "average" => TotalsCalculation.Average,
            "max" => TotalsCalculation.Max,
            "min" => TotalsCalculation.Min,
            "none" => TotalsCalculation.None,
            "countnums" => TotalsCalculation.CountNums,
            "stddev" => TotalsCalculation.StdDev,
            "var" => TotalsCalculation.Var,
            _ => throw new ArgumentException(
                $"Unknown total function: '{function}'. Supported: sum, count, average, max, min, none, countnums, stddev, var")
        };
    }

    private static AddTotalRowParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var tableIndex = parameters.GetOptional<int?>("tableIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var totalFunction = parameters.GetOptional<string?>("totalFunction");

        if (!tableIndex.HasValue)
            throw new ArgumentException("tableIndex is required for add_total_row operation");

        return new AddTotalRowParameters(sheetIndex, tableIndex.Value, columnIndex, totalFunction);
    }

    private sealed record AddTotalRowParameters(
        int SheetIndex,
        int TableIndex,
        int? ColumnIndex,
        string? TotalFunction);
}
