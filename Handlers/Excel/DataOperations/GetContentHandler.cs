using System.Data;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataOperations;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for getting content from Excel worksheets.
/// </summary>
[ResultType(typeof(GetContentResult))]
public class GetContentHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_content";

    /// <summary>
    ///     Gets content from a range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, range
    /// </param>
    /// <returns>JSON string containing the range content.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var contentParams = ExtractGetContentParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, contentParams.SheetIndex);
            var cells = worksheet.Cells;

            List<Dictionary<string, object?>> rows;
            if (!string.IsNullOrEmpty(contentParams.Range))
            {
                var cellRange = ExcelHelper.CreateRange(cells, contentParams.Range);
                var options = new ExportTableOptions { ExportColumnName = false };
                var dataTable = cells.ExportDataTable(cellRange.FirstRow, cellRange.FirstColumn,
                    cellRange.RowCount, cellRange.ColumnCount, options);

                rows = ConvertDataTableToList(dataTable);
            }
            else
            {
                var maxRow = cells.MaxDataRow;
                var maxCol = cells.MaxDataColumn;
                var options = new ExportTableOptions { ExportColumnName = false };
                var dataTable = cells.ExportDataTable(0, 0, maxRow + 1, maxCol + 1, options);

                rows = ConvertDataTableToList(dataTable);
            }

            return new GetContentResult { Rows = rows };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Extracts get content parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get content parameters.</returns>
    private static GetContentParameters ExtractGetContentParameters(OperationParameters parameters)
    {
        return new GetContentParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("range")
        );
    }

    /// <summary>
    ///     Converts a DataTable to a list of dictionaries for JSON serialization.
    /// </summary>
    /// <param name="dataTable">The DataTable to convert.</param>
    /// <returns>A list of dictionaries representing the table rows.</returns>
    private static List<Dictionary<string, object?>> ConvertDataTableToList(DataTable dataTable)
    {
        List<Dictionary<string, object?>> rows = [];
        foreach (DataRow row in dataTable.Rows)
        {
            var rowDict = new Dictionary<string, object?>();
            foreach (DataColumn column in dataTable.Columns)
            {
                var value = row[column];
                rowDict[column.ColumnName] = value == DBNull.Value ? null : value;
            }

            rows.Add(rowDict);
        }

        return rows;
    }

    /// <summary>
    ///     Parameters for get content operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="Range">The cell range to get content from, or null for entire used range.</param>
    private sealed record GetContentParameters(int SheetIndex, string? Range);
}
