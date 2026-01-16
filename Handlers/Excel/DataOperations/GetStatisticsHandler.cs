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
        var statisticsParams = ExtractGetStatisticsParameters(parameters);

        try
        {
            var workbook = context.Document;
            List<object> worksheets = [];

            if (statisticsParams.SheetIndex.HasValue)
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, statisticsParams.SheetIndex.Value);
                worksheets.Add(GetSheetStatistics(worksheet, statisticsParams.SheetIndex.Value,
                    statisticsParams.Range));
            }
            else
            {
                for (var i = 0; i < workbook.Worksheets.Count; i++)
                    worksheets.Add(GetSheetStatistics(workbook.Worksheets[i], i, statisticsParams.Range));
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
    ///     Extracts get statistics parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get statistics parameters.</returns>
    private static GetStatisticsParameters ExtractGetStatisticsParameters(OperationParameters parameters)
    {
        return new GetStatisticsParameters(
            parameters.GetOptional<int?>("sheetIndex"),
            parameters.GetOptional<string?>("range")
        );
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
        var baseStats = BuildBaseStats(worksheet, index);

        if (!string.IsNullOrEmpty(range))
            TryAddRangeStatistics(worksheet, range, baseStats);

        return baseStats;
    }

    /// <summary>
    ///     Builds base statistics for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get statistics for.</param>
    /// <param name="index">The worksheet index.</param>
    /// <returns>A dictionary containing basic worksheet statistics.</returns>
    private static Dictionary<string, object> BuildBaseStats(Worksheet worksheet, int index)
    {
        return new Dictionary<string, object>
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
    }

    /// <summary>
    ///     Tries to add range statistics to the base statistics dictionary.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="range">The range to calculate statistics for.</param>
    /// <param name="baseStats">The base statistics dictionary to add to.</param>
    private static void TryAddRangeStatistics(Worksheet worksheet, string range, Dictionary<string, object> baseStats)
    {
        try
        {
            var rangeStats = CalculateRangeStatistics(worksheet, range);
            baseStats["rangeStatistics"] = rangeStats;
        }
        catch (Exception ex)
        {
            baseStats["rangeStatisticsError"] = ex.Message;
        }
    }

    /// <summary>
    ///     Calculates statistics for a specific range.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="range">The range to calculate statistics for.</param>
    /// <returns>A dictionary containing range statistics.</returns>
    private static Dictionary<string, object> CalculateRangeStatistics(Worksheet worksheet, string range)
    {
        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
        var (numericValues, nonNumericCount, emptyCount) = CollectCellValues(worksheet, cellRange);

        var rangeStats = new Dictionary<string, object>
        {
            ["range"] = range,
            ["totalCells"] = cellRange.RowCount * cellRange.ColumnCount,
            ["numericCells"] = numericValues.Count,
            ["nonNumericCells"] = nonNumericCount,
            ["emptyCells"] = emptyCount
        };

        if (numericValues.Count > 0)
            AddNumericStatistics(numericValues, rangeStats);

        return rangeStats;
    }

    /// <summary>
    ///     Collects and classifies cell values from a range.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="cellRange">The range to collect values from.</param>
    /// <returns>A tuple containing numeric values, non-numeric count, and empty count.</returns>
    private static (List<double> numericValues, int nonNumericCount, int emptyCount) CollectCellValues(
        Worksheet worksheet, Aspose.Cells.Range cellRange)
    {
        List<double> numericValues = [];
        var nonNumericCount = 0;
        var emptyCount = 0;

        for (var row = cellRange.FirstRow; row < cellRange.FirstRow + cellRange.RowCount; row++)
        for (var col = cellRange.FirstColumn; col < cellRange.FirstColumn + cellRange.ColumnCount; col++)
        {
            var value = worksheet.Cells[row, col].Value;
            ClassifyCellValue(value, numericValues, ref nonNumericCount, ref emptyCount);
        }

        return (numericValues, nonNumericCount, emptyCount);
    }

    /// <summary>
    ///     Classifies a cell value as numeric, non-numeric, or empty.
    /// </summary>
    /// <param name="value">The cell value to classify.</param>
    /// <param name="numericValues">The list to add numeric values to.</param>
    /// <param name="nonNumericCount">The count of non-numeric values.</param>
    /// <param name="emptyCount">The count of empty values.</param>
    private static void ClassifyCellValue(object? value, List<double> numericValues, ref int nonNumericCount,
        ref int emptyCount)
    {
        if (value == null || (value is string str && string.IsNullOrWhiteSpace(str)))
        {
            emptyCount++;
            return;
        }

        if (value is double or int or float or decimal)
        {
            numericValues.Add(Convert.ToDouble(value));
            return;
        }

        if (double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out var numValue))
            numericValues.Add(numValue);
        else
            nonNumericCount++;
    }

    /// <summary>
    ///     Adds numeric statistics to the range statistics dictionary.
    /// </summary>
    /// <param name="numericValues">The list of numeric values.</param>
    /// <param name="rangeStats">The range statistics dictionary to add to.</param>
    private static void AddNumericStatistics(List<double> numericValues, Dictionary<string, object> rangeStats)
    {
        numericValues.Sort();
        rangeStats["sum"] = Math.Round(numericValues.Sum(), 2);
        rangeStats["average"] = Math.Round(numericValues.Sum() / numericValues.Count, 2);
        rangeStats["min"] = Math.Round(numericValues[0], 2);
        rangeStats["max"] = Math.Round(numericValues[^1], 2);
        rangeStats["count"] = numericValues.Count;
    }

    /// <summary>
    ///     Parameters for get statistics operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based), or null to get statistics for all sheets.</param>
    /// <param name="Range">The cell range to calculate statistics for, or null for basic sheet info.</param>
    private sealed record GetStatisticsParameters(int? SheetIndex, string? Range);
}
