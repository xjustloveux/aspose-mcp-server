using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Handler for editing hyperlinks in Excel worksheets.
/// </summary>
public class EditExcelHyperlinkHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell or hyperlinkIndex (at least one)
    ///     Optional: sheetIndex (default: 0), url, displayText
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var url = parameters.GetOptional<string?>("url");
        var displayText = parameters.GetOptional<string?>("displayText");
        var hyperlinkIndex = parameters.GetOptional<int?>("hyperlinkIndex");

        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for edit operation.");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        var index = ExcelHyperlinkHelper.ResolveHyperlinkIndex(hyperlinks, hyperlinkIndex, cell);
        var hyperlink = hyperlinks[index];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(url))
        {
            hyperlink.Address = url;
            changes.Add($"url={url}");
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            hyperlink.TextToDisplay = displayText;
            changes.Add($"displayText={displayText}");
        }

        var cellRef = CellsHelper.CellIndexToName(hyperlink.Area.StartRow, hyperlink.Area.StartColumn);

        if (changes.Count > 0)
        {
            MarkModified(context);
            return Success($"Hyperlink at {cellRef} edited: {string.Join(", ", changes)}.");
        }

        return Success($"Hyperlink at {cellRef} unchanged.");
    }
}
