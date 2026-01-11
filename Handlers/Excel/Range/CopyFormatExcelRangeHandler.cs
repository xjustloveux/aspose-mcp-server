using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Handler for copying format (and optionally values) from source range to destination.
/// </summary>
public class CopyFormatExcelRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "copy_format";

    /// <summary>
    ///     Copies format (and optionally values) from source range to destination.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, destTarget
    ///     Optional: sheetIndex, copyValue
    /// </param>
    /// <returns>Success message with copy format details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var copyValue = parameters.GetOptional("copyValue", false);

        if (!parameters.Has("range"))
            throw new ArgumentException("range or sourceRange is required for copy_format operation");
        var range = parameters.GetRequired<string>("range");

        if (!parameters.Has("destTarget"))
            throw new ArgumentException(
                "Either destRange or destCell is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5' or destCell='E1'");
        var destTarget = parameters.GetRequired<string>("destTarget");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var sourceCellRange = ExcelHelper.CreateRange(cells, range, "source range");
        var destCellRange = ExcelHelper.CreateRange(cells, destTarget, "destination");

        var pasteOptions = new PasteOptions
        {
            PasteType = copyValue ? PasteType.All : PasteType.Formats
        };

        destCellRange.Copy(sourceCellRange, pasteOptions);

        MarkModified(context);

        var result = copyValue ? "Format with values copied" : "Format copied";
        return Success($"{result} from {range} to {destTarget}.");
    }
}
