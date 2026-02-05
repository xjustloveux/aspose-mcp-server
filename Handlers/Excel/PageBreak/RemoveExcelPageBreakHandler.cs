using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.PageBreak;

/// <summary>
///     Handler for removing a page break from an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveExcelPageBreakHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "remove";

    /// <summary>
    ///     Removes a page break by type and index.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: breakType (horizontal or vertical), breakIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var breakType = parameters.GetOptional<string?>("breakType");
        var breakIndex = parameters.GetOptional<int?>("breakIndex");

        if (string.IsNullOrEmpty(breakType))
            throw new ArgumentException("breakType is required for remove operation (horizontal or vertical)");
        if (!breakIndex.HasValue)
            throw new ArgumentException("breakIndex is required for remove operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            switch (breakType.ToLowerInvariant())
            {
                case "horizontal":
                    if (breakIndex.Value < 0 || breakIndex.Value >= worksheet.HorizontalPageBreaks.Count)
                        throw new ArgumentException(
                            $"Horizontal page break index {breakIndex.Value} is out of range (worksheet has {worksheet.HorizontalPageBreaks.Count} horizontal page breaks)");
                    worksheet.HorizontalPageBreaks.RemoveAt(breakIndex.Value);
                    break;

                case "vertical":
                    if (breakIndex.Value < 0 || breakIndex.Value >= worksheet.VerticalPageBreaks.Count)
                        throw new ArgumentException(
                            $"Vertical page break index {breakIndex.Value} is out of range (worksheet has {worksheet.VerticalPageBreaks.Count} vertical page breaks)");
                    worksheet.VerticalPageBreaks.RemoveAt(breakIndex.Value);
                    break;

                default:
                    throw new ArgumentException(
                        $"Unknown break type: '{breakType}'. Use 'horizontal' or 'vertical'.");
            }

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"{breakType} page break at index {breakIndex.Value} removed from sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to remove page break: {ex.Message}");
        }
    }
}
