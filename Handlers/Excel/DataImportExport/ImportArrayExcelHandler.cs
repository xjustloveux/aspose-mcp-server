using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataImportExport;

namespace AsposeMcpServer.Handlers.Excel.DataImportExport;

/// <summary>
///     Handler for importing array data into an Excel worksheet.
/// </summary>
[ResultType(typeof(ImportExcelResult))]
public class ImportArrayExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "import_array";

    /// <summary>
    ///     Imports string array data into a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: arrayData (comma-separated values, rows separated by semicolons)
    ///     Optional: sheetIndex (default: 0), startCell (default: "A1"), isVertical (default: false)
    /// </param>
    /// <returns>Import result with row/column counts.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var arrayData = parameters.GetOptional<string?>("arrayData");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startCell = parameters.GetOptional("startCell", "A1");
        var isVertical = parameters.GetOptional("isVertical", false);

        if (string.IsNullOrEmpty(arrayData))
            throw new ArgumentException("arrayData is required for import_array operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var cell = worksheet.Cells[startCell];
            var startRow = cell.Row;
            var startColumn = cell.Column;

            var rows = arrayData.Split(';', StringSplitOptions.RemoveEmptyEntries);
            var rowCount = rows.Length;
            var maxCols = 0;

            if (rows.Length == 1 && !rows[0].Contains(','))
            {
                var singleValues = new[] { rows[0].Trim() };
                worksheet.Cells.ImportArray(singleValues, startRow, startColumn, isVertical);
                maxCols = 1;
            }
            else if (rows.Length == 1)
            {
                var values = rows[0].Split(',', StringSplitOptions.TrimEntries);
                worksheet.Cells.ImportArray(values, startRow, startColumn, isVertical);
                maxCols = values.Length;
                rowCount = isVertical ? values.Length : 1;
            }
            else
            {
                for (var i = 0; i < rows.Length; i++)
                {
                    var values = rows[i].Split(',', StringSplitOptions.TrimEntries);
                    if (values.Length > maxCols) maxCols = values.Length;
                    worksheet.Cells.ImportArray(values, startRow + i, startColumn, false);
                }
            }

            MarkModified(context);

            return new ImportExcelResult
            {
                RowCount = rowCount,
                ColumnCount = maxCols,
                StartCell = startCell,
                Message =
                    $"Array data ({rowCount} rows, {maxCols} columns) imported starting at {startCell} in sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to import array data: {ex.Message}");
        }
    }
}
