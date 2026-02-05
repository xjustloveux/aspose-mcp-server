using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sparkline;

/// <summary>
///     Handler for adding a sparkline group to an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddExcelSparklineHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a sparkline group to a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: dataRange, locationRange
    ///     Optional: sheetIndex (default: 0), type (line/column/stacked, default: line),
    ///     isVertical (auto-detected from data range shape if not specified)
    /// </param>
    /// <returns>Success message with sparkline group details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            var sparklineType = ResolveSparklineType(p.Type);
            var locationArea = CellArea.CreateCellArea(p.LocationRange, p.LocationRange);
            var isVertical = ResolveIsVertical(p.IsVertical, p.DataRange);

            worksheet.SparklineGroups.Add(sparklineType, p.DataRange, isVertical, locationArea);

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Sparkline group ({p.Type}) added with data range '{p.DataRange}' at '{p.LocationRange}' in sheet {p.SheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add sparkline: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves a sparkline type string to a SparklineType enum value.
    /// </summary>
    /// <param name="type">The type string (line, column, stacked).</param>
    /// <returns>The corresponding SparklineType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the type is unknown.</exception>
    internal static SparklineType ResolveSparklineType(string type)
    {
        return type.ToLowerInvariant() switch
        {
            "line" => SparklineType.Line,
            "column" => SparklineType.Column,
            "stacked" or "win_loss" => SparklineType.Stacked,
            _ => throw new ArgumentException(
                $"Unknown sparkline type: '{type}'. Supported: line, column, stacked (win_loss)")
        };
    }

    /// <summary>
    ///     Resolves the isVertical flag for sparkline data orientation.
    ///     If explicitly provided, uses that value. Otherwise, auto-detects from the data range shape:
    ///     vertical (column) data ranges default to true, horizontal (row) ranges to false.
    /// </summary>
    /// <param name="isVertical">The explicitly provided value, or null for auto-detection.</param>
    /// <param name="dataRange">The data range string (e.g., "Sheet1!A1:A5").</param>
    /// <returns>The resolved isVertical value.</returns>
    internal static bool ResolveIsVertical(bool? isVertical, string dataRange)
    {
        if (isVertical.HasValue)
            return isVertical.Value;

        var range = dataRange;
        var exclamation = range.LastIndexOf('!');
        if (exclamation >= 0)
            range = range[(exclamation + 1)..];

        var parts = range.Split(':');
        if (parts.Length != 2)
            return false;

        CellsHelper.CellNameToIndex(parts[0], out var startRow, out var startCol);
        CellsHelper.CellNameToIndex(parts[1], out var endRow, out var endCol);

        var rowCount = endRow - startRow + 1;
        var colCount = endCol - startCol + 1;

        return rowCount > colCount;
    }

    private static AddParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var dataRange = parameters.GetOptional<string?>("dataRange");
        var locationRange = parameters.GetOptional<string?>("locationRange");
        var type = parameters.GetOptional("type", "line");
        var isVertical = parameters.GetOptional<bool?>("isVertical");

        if (string.IsNullOrEmpty(dataRange))
            throw new ArgumentException("dataRange is required for add operation");
        if (string.IsNullOrEmpty(locationRange))
            throw new ArgumentException("locationRange is required for add operation");

        return new AddParameters(sheetIndex, dataRange, locationRange, type, isVertical);
    }

    private sealed record AddParameters(
        int SheetIndex,
        string DataRange,
        string LocationRange,
        string Type,
        bool? IsVertical);
}
