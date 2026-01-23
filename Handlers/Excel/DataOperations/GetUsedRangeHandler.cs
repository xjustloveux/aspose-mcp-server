using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataOperations;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for getting the used range information from Excel worksheets.
/// </summary>
[ResultType(typeof(GetUsedRangeResult))]
public class GetUsedRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_used_range";

    /// <summary>
    ///     Gets the used range information for the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>JSON string containing the used range information.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var usedRangeParams = ExtractGetUsedRangeParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, usedRangeParams.SheetIndex);
            var cells = worksheet.Cells;

            string? rangeAddress = null;
            if (cells.MaxDataRow >= cells.MinDataRow && cells.MaxDataColumn >= cells.MinDataColumn)
            {
                var firstCell = CellsHelper.CellIndexToName(cells.MinDataRow, cells.MinDataColumn);
                var lastCell = CellsHelper.CellIndexToName(cells.MaxDataRow, cells.MaxDataColumn);
                rangeAddress = $"{firstCell}:{lastCell}";
            }

            return new GetUsedRangeResult
            {
                WorksheetName = worksheet.Name,
                SheetIndex = usedRangeParams.SheetIndex,
                FirstRow = cells.MinDataRow,
                LastRow = cells.MaxDataRow,
                FirstColumn = cells.MinDataColumn,
                LastColumn = cells.MaxDataColumn,
                Range = rangeAddress
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Extracts get used range parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get used range parameters.</returns>
    private static GetUsedRangeParameters ExtractGetUsedRangeParameters(OperationParameters parameters)
    {
        return new GetUsedRangeParameters(parameters.GetOptional("sheetIndex", 0));
    }

    /// <summary>
    ///     Parameters for get used range operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    private sealed record GetUsedRangeParameters(int SheetIndex);
}
