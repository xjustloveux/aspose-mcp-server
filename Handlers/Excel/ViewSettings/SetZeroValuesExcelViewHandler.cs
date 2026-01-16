using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting zero values visibility in Excel worksheets.
/// </summary>
public class SetZeroValuesExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_zero_values";

    /// <summary>
    ///     Sets the visibility of zero values in the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), visible (default: true)
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetZeroValuesParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        worksheet.DisplayZeros = p.Visible;

        MarkModified(context);
        return Success($"Zero values visibility set to {(p.Visible ? "visible" : "hidden")}.");
    }

    /// <summary>
    ///     Extracts set zero values parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetZeroValuesParameters record containing all extracted values.</returns>
    private static SetZeroValuesParameters ExtractSetZeroValuesParameters(OperationParameters parameters)
    {
        return new SetZeroValuesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetOptional("visible", true)
        );
    }

    /// <summary>
    ///     Record containing parameters for set zero values operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Visible">Whether zero values should be visible.</param>
    private record SetZeroValuesParameters(
        int SheetIndex,
        bool Visible);
}
