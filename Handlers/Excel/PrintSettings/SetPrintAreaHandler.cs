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
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetOptional<string?>("range");
        var clearPrintArea = parameters.GetOptional("clearPrintArea", false);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (clearPrintArea)
            worksheet.PageSetup.PrintArea = "";
        else if (!string.IsNullOrEmpty(range))
            worksheet.PageSetup.PrintArea = range;
        else
            throw new ArgumentException("Either range or clearPrintArea must be provided");

        MarkModified(context);

        return clearPrintArea
            ? Success($"Print area cleared for sheet {sheetIndex}.")
            : Success($"Print area set to {range} for sheet {sheetIndex}.");
    }
}
