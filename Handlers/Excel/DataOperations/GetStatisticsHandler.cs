using System.Globalization;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for getting statistics from Excel worksheets.
/// </summary>
public class GetStatisticsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_statistics";

    /// <summary>
    ///     Gets statistics for a range (count, sum, average, min, max).
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, range
    /// </param>
    /// <returns>JSON string containing the statistics.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional<int?>("sheetIndex");
        var range = parameters.GetOptional<string?>("range");

        try
        {
            var workbook = context.Document;
            List<object> worksheets = [];

            if (sheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);
                worksheets.Add(GetSheetStatistics(worksheet, sheetIndex.Value, range));
            }
            else
            {
                for (var i = 0; i < workbook.Worksheets.Count; i++)
                    worksheets.Add(GetSheetStatistics(workbook.Worksheets[i], i, range));
            }

            var result = new
            {
                totalWorksheets = workbook.Worksheets.Count,
                fileFormat = workbook.FileFormat.ToString(),
                worksheets
            };

            return JsonSerializer.Serialize(result, JsonDefaults.Indented);
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    /// <summary>
    ///     Gets statistics for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get statistics for.</param>
    /// <param name="index">The worksheet index.</param>
    /// <param name="range">The cell range to calculate statistics for, or null for basic sheet info.</param>
    /// <returns>An object containing the worksheet statistics.</returns>
    private static object GetSheetStatistics(Worksheet worksheet, int index, string? range)
    {
        var baseStats = new Dictionary<string, object>
        {
            ["index"] = index,
            ["name"] = worksheet.Name,
            ["maxDataRow"] = worksheet.Cells.MaxDataRow + 1,
            ["maxDataColumn"] = worksheet.Cells.MaxDataColumn + 1,
            ["chartsCount"] = worksheet.Charts.Count,
            ["picturesCount"] = worksheet.Pictures.Count,
            ["hyperlinksCount"] = worksheet.Hyperlinks.Count,
            ["commentsCount"] = worksheet.Comments.Count
        };

        if (!string.IsNullOrEmpty(range))
            try
            {
                var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
                List<double> numericValues = [];
                var nonNumericCount = 0;
                var emptyCount = 0;

                for (var row = cellRange.FirstRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
                for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var value = cell.Value;

                    if (value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
                        emptyCount++;
                    else if (value is double || value is int || value is float || value is decimal)
                        numericValues.Add(Convert.ToDouble(value));
                    else if (double.TryParse(value.ToString(), NumberStyles.Any,
                                 CultureInfo.InvariantCulture, out var numValue))
                        numericValues.Add(numValue);
                    else
                        nonNumericCount++;
                }

                var rangeStats = new Dictionary<string, object>
                {
                    ["range"] = range,
                    ["totalCells"] = cellRange.RowCount * cellRange.ColumnCount,
                    ["numericCells"] = numericValues.Count,
                    ["nonNumericCells"] = nonNumericCount,
                    ["emptyCells"] = emptyCount
                };

                if (numericValues.Count > 0)
                {
                    numericValues.Sort();
                    rangeStats["sum"] = Math.Round(numericValues.Sum(), 2);
                    rangeStats["average"] = Math.Round(numericValues.Sum() / numericValues.Count, 2);
                    rangeStats["min"] = Math.Round(numericValues[0], 2);
                    rangeStats["max"] = Math.Round(numericValues[^1], 2);
                    rangeStats["count"] = numericValues.Count;
                }

                baseStats["rangeStatistics"] = rangeStats;
            }
            catch (Exception ex)
            {
                baseStats["rangeStatisticsError"] = ex.Message;
            }

        return baseStats;
    }
}
