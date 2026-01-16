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
        var addParams = ExtractAddParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);

        var existingIndex = ExcelHyperlinkHelper.FindHyperlinkIndexByCell(worksheet.Hyperlinks, addParams.Cell);
        if (existingIndex.HasValue)
            throw new ArgumentException(
                $"Cell {addParams.Cell} already has a hyperlink. Use 'edit' operation to modify it.");

        var hyperlinkIdx = worksheet.Hyperlinks.Add(addParams.Cell, 1, 1, addParams.Url);
        var hyperlink = worksheet.Hyperlinks[hyperlinkIdx];

        if (!string.IsNullOrEmpty(addParams.DisplayText))
            hyperlink.TextToDisplay = addParams.DisplayText;

        MarkModified(context);

        return Success($"Hyperlink added to {addParams.Cell}: {addParams.Url}.");
    }

    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var url = parameters.GetOptional<string?>("url");
        var displayText = parameters.GetOptional<string?>("displayText");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");
        if (string.IsNullOrEmpty(url))
            throw new ArgumentException("url is required for add operation");

        return new AddParameters(sheetIndex, cell, url, displayText);
    }

    private record AddParameters(int SheetIndex, string Cell, string Url, string? DisplayText);
}
