using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for applying auto filter to Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ApplyExcelFilterHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "apply";

    /// <summary>
    ///     Applies auto filter dropdown buttons to a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with filter details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var applyFilterParams = ExtractApplyFilterParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, applyFilterParams.SheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, applyFilterParams.Range);
        worksheet.AutoFilter.Range = applyFilterParams.Range;

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Auto filter applied to range {applyFilterParams.Range} in sheet {applyFilterParams.SheetIndex}."
        };
    }

    /// <summary>
    ///     Extracts apply filter parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted apply filter parameters.</returns>
    private static ApplyFilterParameters ExtractApplyFilterParameters(OperationParameters parameters)
    {
        return new ApplyFilterParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range")
        );
    }

    /// <summary>
    ///     Parameters for apply filter operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="Range">The range to apply auto filter to.</param>
    private sealed record ApplyFilterParameters(int SheetIndex, string Range);
}
