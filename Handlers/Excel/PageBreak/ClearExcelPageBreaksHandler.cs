using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PageBreak;

/// <summary>
///     Handler for clearing all page breaks from an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ClearExcelPageBreaksHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "clear";

    /// <summary>
    ///     Clears all page breaks from a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), breakType (horizontal, vertical, or all; default: all)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var breakType = parameters.GetOptional("breakType", "all");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var hCount = 0;
            var vCount = 0;

            switch (breakType.ToLowerInvariant())
            {
                case "horizontal":
                    hCount = worksheet.HorizontalPageBreaks.Count;
                    for (var i = hCount - 1; i >= 0; i--)
                        worksheet.HorizontalPageBreaks.RemoveAt(i);
                    break;

                case "vertical":
                    vCount = worksheet.VerticalPageBreaks.Count;
                    for (var i = vCount - 1; i >= 0; i--)
                        worksheet.VerticalPageBreaks.RemoveAt(i);
                    break;

                case "all":
                    hCount = worksheet.HorizontalPageBreaks.Count;
                    for (var i = hCount - 1; i >= 0; i--)
                        worksheet.HorizontalPageBreaks.RemoveAt(i);
                    vCount = worksheet.VerticalPageBreaks.Count;
                    for (var i = vCount - 1; i >= 0; i--)
                        worksheet.VerticalPageBreaks.RemoveAt(i);
                    break;

                default:
                    throw new ArgumentException(
                        $"Unknown break type: '{breakType}'. Use 'horizontal', 'vertical', or 'all'.");
            }

            MarkModified(context);

            return new SuccessResult
            {
                Message = breakType.ToLowerInvariant() switch
                {
                    "horizontal" => $"Cleared {hCount} horizontal page break(s) from sheet {sheetIndex}.",
                    "vertical" => $"Cleared {vCount} vertical page break(s) from sheet {sheetIndex}.",
                    _ => $"Cleared {hCount} horizontal and {vCount} vertical page break(s) from sheet {sheetIndex}."
                }
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to clear page breaks: {ex.Message}");
        }
    }
}
