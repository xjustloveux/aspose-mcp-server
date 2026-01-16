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
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractCopyFormatExcelRangeParameters(parameters);

        if (string.IsNullOrEmpty(p.Range))
            throw new ArgumentException("range or sourceRange is required for copy_format operation");

        if (string.IsNullOrEmpty(p.DestTarget))
            throw new ArgumentException(
                "Either destRange or destCell is required for copy_format operation. Example: range='A1:C5', destRange='E1:G5' or destCell='E1'");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var cells = worksheet.Cells;

        var sourceCellRange = ExcelHelper.CreateRange(cells, p.Range, "source range");
        var destCellRange = ExcelHelper.CreateRange(cells, p.DestTarget, "destination");

        var pasteOptions = new PasteOptions
        {
            PasteType = p.CopyValue ? PasteType.All : PasteType.Formats
        };

        destCellRange.Copy(sourceCellRange, pasteOptions);

        MarkModified(context);

        var result = p.CopyValue ? "Format with values copied" : "Format copied";
        return Success($"{result} from {p.Range} to {p.DestTarget}.");
    }

    /// <summary>
    ///     Extracts parameters for CopyFormatExcelRange operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static CopyFormatExcelRangeParameters ExtractCopyFormatExcelRangeParameters(OperationParameters parameters)
    {
        return new CopyFormatExcelRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("copyValue", false),
            parameters.Has("range") ? parameters.GetRequired<string>("range") : string.Empty,
            parameters.Has("destTarget") ? parameters.GetRequired<string>("destTarget") : string.Empty
        );
    }

    /// <summary>
    ///     Parameters for CopyFormatExcelRange operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index.</param>
    /// <param name="CopyValue">Whether to copy values along with format.</param>
    /// <param name="Range">The source range.</param>
    /// <param name="DestTarget">The destination range or cell.</param>
    private sealed record CopyFormatExcelRangeParameters(
        int SheetIndex,
        bool CopyValue,
        string Range,
        string DestTarget);
}
