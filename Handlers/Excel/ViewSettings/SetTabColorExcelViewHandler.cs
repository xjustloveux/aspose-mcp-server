using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Handler for setting worksheet tab color in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetTabColorExcelViewHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_tab_color";

    /// <summary>
    ///     Sets the tab color of a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: color
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractSetTabColorParameters(parameters);

        var worksheet = ExcelHelper.GetWorksheet(context.Document, p.SheetIndex);
        var parsedColor = ColorHelper.ParseColor(p.Color);
        worksheet.TabColor = parsedColor;

        MarkModified(context);
        return new SuccessResult { Message = $"Sheet tab color set to {p.Color}." };
    }

    /// <summary>
    ///     Extracts set tab color parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A SetTabColorParameters record containing all extracted values.</returns>
    private static SetTabColorParameters ExtractSetTabColorParameters(OperationParameters parameters)
    {
        return new SetTabColorParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("color")
        );
    }

    /// <summary>
    ///     Record containing parameters for set tab color operations.
    /// </summary>
    /// <param name="SheetIndex">The index of the worksheet.</param>
    /// <param name="Color">The color to set for the tab.</param>
    private sealed record SetTabColorParameters(
        int SheetIndex,
        string Color);
}
