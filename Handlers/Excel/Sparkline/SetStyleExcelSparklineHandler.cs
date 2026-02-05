using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sparkline;

/// <summary>
///     Handler for setting the style of a sparkline group in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetStyleExcelSparklineHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_style";

    /// <summary>
    ///     Sets the preset style of a sparkline group.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: groupIndex, presetStyle
    ///     Optional: sheetIndex (default: 0), showHighPoint, showLowPoint, showFirstPoint, showLastPoint,
    ///     showNegativePoints, showMarkers
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

            if (p.GroupIndex < 0 || p.GroupIndex >= worksheet.SparklineGroups.Count)
                throw new ArgumentException(
                    $"Sparkline group index {p.GroupIndex} is out of range (worksheet has {worksheet.SparklineGroups.Count} sparkline groups)");

            var group = worksheet.SparklineGroups[p.GroupIndex];

            if (!string.IsNullOrEmpty(p.PresetStyle))
            {
                if (!Enum.TryParse<SparklinePresetStyleType>(p.PresetStyle, true, out var styleType))
                    throw new ArgumentException(
                        $"Unknown sparkline preset style: '{p.PresetStyle}'. Use Style1-Style36.");
                group.PresetStyle = styleType;
            }

            if (p.ShowHighPoint.HasValue) group.ShowHighPoint = p.ShowHighPoint.Value;
            if (p.ShowLowPoint.HasValue) group.ShowLowPoint = p.ShowLowPoint.Value;
            if (p.ShowFirstPoint.HasValue) group.ShowFirstPoint = p.ShowFirstPoint.Value;
            if (p.ShowLastPoint.HasValue) group.ShowLastPoint = p.ShowLastPoint.Value;
            if (p.ShowNegativePoints.HasValue) group.ShowNegativePoints = p.ShowNegativePoints.Value;
            if (p.ShowMarkers.HasValue) group.ShowMarkers = p.ShowMarkers.Value;

            MarkModified(context);

            return new SuccessResult
            {
                Message = $"Sparkline group at index {p.GroupIndex} style updated in sheet {p.SheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to set sparkline style: {ex.Message}");
        }
    }

    private static SetStyleParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var groupIndex = parameters.GetOptional<int?>("groupIndex");
        var presetStyle = parameters.GetOptional<string?>("presetStyle");
        var showHighPoint = parameters.GetOptional<bool?>("showHighPoint");
        var showLowPoint = parameters.GetOptional<bool?>("showLowPoint");
        var showFirstPoint = parameters.GetOptional<bool?>("showFirstPoint");
        var showLastPoint = parameters.GetOptional<bool?>("showLastPoint");
        var showNegativePoints = parameters.GetOptional<bool?>("showNegativePoints");
        var showMarkers = parameters.GetOptional<bool?>("showMarkers");

        if (!groupIndex.HasValue)
            throw new ArgumentException("groupIndex is required for set_style operation");

        return new SetStyleParameters(sheetIndex, groupIndex.Value, presetStyle, showHighPoint, showLowPoint,
            showFirstPoint, showLastPoint, showNegativePoints, showMarkers);
    }

    private sealed record SetStyleParameters(
        int SheetIndex,
        int GroupIndex,
        string? PresetStyle,
        bool? ShowHighPoint,
        bool? ShowLowPoint,
        bool? ShowFirstPoint,
        bool? ShowLastPoint,
        bool? ShowNegativePoints,
        bool? ShowMarkers);
}
