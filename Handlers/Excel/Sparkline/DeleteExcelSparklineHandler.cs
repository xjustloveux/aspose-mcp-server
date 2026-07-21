using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Errors.Excel;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sparkline;

/// <summary>
///     Handler for deleting a sparkline group from an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteExcelSparklineHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a sparkline group from the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: groupIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var groupIndex = parameters.GetOptional<int?>("groupIndex");

        if (!groupIndex.HasValue)
            throw new ArgumentException("groupIndex is required for delete operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (groupIndex.Value < 0 || groupIndex.Value >= worksheet.SparklineGroups.Count)
                throw new ArgumentException(
                    $"Sparkline group index {groupIndex.Value} is out of range (worksheet has {worksheet.SparklineGroups.Count} sparkline groups)");

            var group = worksheet.SparklineGroups[groupIndex.Value];
            var sparklineCount = group.Sparklines.Count;

            // Remove exactly the indexed group. An area-based ClearSparklineGroups would also
            // wipe any other group overlapping this group's bounding box (and an empty group's
            // default area would target A1).
            worksheet.SparklineGroups.RemoveAt(groupIndex.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Sparkline group at index {groupIndex.Value} ({sparklineCount} sparkline(s)) deleted from sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw CellsErrorTranslator.Translate(ex);
        }
    }
}
