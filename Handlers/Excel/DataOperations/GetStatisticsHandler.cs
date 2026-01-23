using System.Globalization;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataOperations;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Handler for getting statistics from Excel worksheets.
/// </summary>
[ResultType(typeof(GetStatisticsResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var statisticsParams = ExtractGetStatisticsParameters(parameters);

        try
        {
            var workbook = context.Document;
            List<WorksheetStatistics> worksheets = [];

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

            return new GetStatisticsResult
            {
                TotalWorksheets = workbook.Worksheets.Count,
                FileFormat = workbook.FileFormat.ToString(),
                Worksheets = worksheets
            };
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
    /// <returns>A WorksheetStatistics object containing the worksheet statistics.</returns>
    private static WorksheetStatistics GetSheetStatistics(Worksheet worksheet, int index, string? range)
    {
        var baseStats = BuildBaseStats(worksheet, index);

        if (!string.IsNullOrEmpty(range))
            baseStats = TryAddRangeStatistics(worksheet, range, baseStats);

        return baseStats;
    }

    /// <summary>
    ///     Builds base statistics for a worksheet.
    /// </summary>
    /// <param name="worksheet">The worksheet to get statistics for.</param>
    /// <param name="index">The worksheet index.</param>
    /// <returns>A WorksheetStatistics object containing basic worksheet statistics.</returns>
    private static WorksheetStatistics BuildBaseStats(Worksheet worksheet, int index)
    {
        return new WorksheetStatistics
        {
            Index = index,
            Name = worksheet.Name,
            MaxDataRow = worksheet.Cells.MaxDataRow + 1,
            MaxDataColumn = worksheet.Cells.MaxDataColumn + 1,
            ChartsCount = worksheet.Charts.Count,
            PicturesCount = worksheet.Pictures.Count,
            HyperlinksCount = worksheet.Hyperlinks.Count,
            CommentsCount = worksheet.Comments.Count
        };
    }

    /// <summary>
    ///     Tries to add range statistics to the worksheet statistics.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="range">The range to calculate statistics for.</param>
    /// <param name="baseStats">The worksheet statistics to update.</param>
    /// <returns>Updated WorksheetStatistics with range statistics or error.</returns>
    private static WorksheetStatistics TryAddRangeStatistics(
        Worksheet worksheet, string range, WorksheetStatistics baseStats)
    {
        try
        {
            var rangeStats = CalculateRangeStatistics(worksheet, range);
            return baseStats with { RangeStatistics = rangeStats };
        }
        catch (Exception ex)
        {
            return baseStats with { RangeStatisticsError = ex.Message };
        }
    }

    /// <summary>
    ///     Calculates statistics for a specific range.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="range">The range to calculate statistics for.</param>
    /// <returns>A RangeStatistics object containing range statistics.</returns>
    private static RangeStatistics CalculateRangeStatistics(Worksheet worksheet, string range)
    {
        var cellRange = ExcelHelper.CreateRange(worksheet.Cells, range);
        var (numericValues, nonNumericCount, emptyCount) = CollectCellValues(worksheet, cellRange);

        var rangeStats = new RangeStatistics
        {
            Range = range,
            TotalCells = cellRange.RowCount * cellRange.ColumnCount,
            NumericCells = numericValues.Count,
            NonNumericCells = nonNumericCount,
            EmptyCells = emptyCount
        };

        if (numericValues.Count > 0)
            rangeStats = AddNumericStatistics(numericValues, rangeStats);

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
    ///     Adds numeric statistics to the range statistics.
    /// </summary>
    /// <param name="numericValues">The list of numeric values.</param>
    /// <param name="rangeStats">The range statistics to update.</param>
    /// <returns>Updated RangeStatistics with numeric statistics.</returns>
    private static RangeStatistics AddNumericStatistics(List<double> numericValues, RangeStatistics rangeStats)
    {
        numericValues.Sort();
        return rangeStats with
        {
            Sum = Math.Round(numericValues.Sum(), 2),
            Average = Math.Round(numericValues.Sum() / numericValues.Count, 2),
            Min = Math.Round(numericValues[0], 2),
            Max = Math.Round(numericValues[^1], 2),
            Count = numericValues.Count
        };
    }

    /// <summary>
    ///     Parameters for get statistics operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based), or null to get statistics for all sheets.</param>
    /// <param name="Range">The cell range to calculate statistics for, or null for basic sheet info.</param>
    private sealed record GetStatisticsParameters(int? SheetIndex, string? Range);
}
