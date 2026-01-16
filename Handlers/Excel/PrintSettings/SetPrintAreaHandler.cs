using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Handler for setting print area in Excel worksheets.
/// </summary>
public class SetPrintAreaHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_print_area";

    /// <summary>
    ///     Sets the print area for a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range (unless clearPrintArea is true)
    ///     Optional: sheetIndex (default: 0), clearPrintArea (default: false)
    /// </param>
    /// <returns>Success message with print area details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var setParams = ExtractSetPrintAreaParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);

        if (setParams.ClearPrintArea)
            worksheet.PageSetup.PrintArea = "";
        else if (!string.IsNullOrEmpty(setParams.Range))
            worksheet.PageSetup.PrintArea = setParams.Range;
        else
            throw new ArgumentException("Either range or clearPrintArea must be provided");

        MarkModified(context);

        return setParams.ClearPrintArea
            ? Success($"Print area cleared for sheet {setParams.SheetIndex}.")
            : Success($"Print area set to {setParams.Range} for sheet {setParams.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts set print area parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted set print area parameters.</returns>
    private static SetPrintAreaParameters ExtractSetPrintAreaParameters(OperationParameters parameters)
    {
        return new SetPrintAreaParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional<string?>("range"),
            parameters.GetOptional("clearPrintArea", false)
        );
    }

    /// <summary>
    ///     Record to hold set print area parameters.
    /// </summary>
    private record SetPrintAreaParameters(int SheetIndex, string? Range, bool ClearPrintArea);
}
