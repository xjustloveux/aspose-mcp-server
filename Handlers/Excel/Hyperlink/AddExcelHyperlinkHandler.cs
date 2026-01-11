using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Handler for adding hyperlinks to Excel worksheets.
/// </summary>
public class AddExcelHyperlinkHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a hyperlink to a cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell, url
    ///     Optional: sheetIndex (default: 0), displayText
    /// </param>
    /// <returns>Success message with hyperlink details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetRequired<string>("cell");
        var url = parameters.GetRequired<string>("url");
        var displayText = parameters.GetOptional<string?>("displayText");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var existingIndex = ExcelHyperlinkHelper.FindHyperlinkIndexByCell(worksheet.Hyperlinks, cell);
        if (existingIndex.HasValue)
            throw new ArgumentException($"Cell {cell} already has a hyperlink. Use 'edit' operation to modify it.");

        var hyperlinkIdx = worksheet.Hyperlinks.Add(cell, 1, 1, url);
        var hyperlink = worksheet.Hyperlinks[hyperlinkIdx];

        if (!string.IsNullOrEmpty(displayText))
            hyperlink.TextToDisplay = displayText;

        MarkModified(context);

        return Success($"Hyperlink added to {cell}: {url}.");
    }
}
