using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for removing auto filter from Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveExcelFilterHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes auto filter from the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var removeFilterParams = ExtractRemoveFilterParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, removeFilterParams.SheetIndex);

        worksheet.RemoveAutoFilter();

        MarkModified(context);

        return new SuccessResult { Message = $"Auto filter removed from sheet {removeFilterParams.SheetIndex}." };
    }

    /// <summary>
    ///     Extracts remove filter parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted remove filter parameters.</returns>
    private static RemoveFilterParameters ExtractRemoveFilterParameters(OperationParameters parameters)
    {
        return new RemoveFilterParameters(
            parameters.GetOptional("sheetIndex", 0)
        );
    }

    /// <summary>
    ///     Parameters for remove filter operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    private sealed record RemoveFilterParameters(int SheetIndex);
}
