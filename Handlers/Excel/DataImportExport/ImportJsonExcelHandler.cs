using Aspose.Cells;
using Aspose.Cells.Utility;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataImportExport;

namespace AsposeMcpServer.Handlers.Excel.DataImportExport;

/// <summary>
///     Handler for importing JSON data into an Excel worksheet.
/// </summary>
[ResultType(typeof(ImportExcelResult))]
public class ImportJsonExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "import_json";

    /// <summary>
    ///     Imports JSON data into a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: jsonData
    ///     Optional: sheetIndex (default: 0), startCell (default: "A1")
    /// </param>
    /// <returns>Import result with row/column counts.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var jsonData = parameters.GetOptional<string?>("jsonData");
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var startCell = parameters.GetOptional("startCell", "A1");

        if (string.IsNullOrEmpty(jsonData))
            throw new ArgumentException("jsonData is required for import_json operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var cell = worksheet.Cells[startCell];
            var startRow = cell.Row;
            var startColumn = cell.Column;

            var layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            JsonUtility.ImportData(jsonData, worksheet.Cells, startRow, startColumn, layoutOptions);

            MarkModified(context);

            var usedRange = worksheet.Cells.MaxDataRow - startRow + 1;
            var usedCols = worksheet.Cells.MaxDataColumn - startColumn + 1;

            return new ImportExcelResult
            {
                RowCount = usedRange > 0 ? usedRange : 0,
                ColumnCount = usedCols > 0 ? usedCols : 0,
                StartCell = startCell,
                Message = $"JSON data imported starting at {startCell} in sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to import JSON data: {ex.Message}");
        }
    }
}
