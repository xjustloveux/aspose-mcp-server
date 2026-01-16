using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for showing or hiding formulas in Excel worksheets.
/// </summary>
public class ShowFormulasExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "show_formulas";

    /// <summary>
    ///     Sets whether to show formulas or calculated values in the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), visible (default: true)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractShowFormulasParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.ShowFormulas = p.Visible;

        MarkModified(context);
        return Success($"Formulas {(p.Visible ? "shown" : "hidden")} for sheet {p.SheetIndex}.");
    }

    /// <summary>
    ///     Extracts show formulas parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A ShowFormulasParameters record containing all extracted values.</returns>
    private static ShowFormulasParameters ExtractShowFormulasParameters(OperationParameters parameters)
    {
        return new ShowFormulasParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("visible", true)
        );
    }

    /// <summary>
    ///     Record containing parameters for show formulas operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Visible">Whether formulas should be visible.</param>
    private sealed record ShowFormulasParameters(
        int SheetIndex,
        bool Visible);
}
