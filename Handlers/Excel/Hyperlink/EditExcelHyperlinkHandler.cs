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
        var editParams = ExtractEditParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        var index = ExcelHyperlinkHelper.ResolveHyperlinkIndex(hyperlinks, editParams.HyperlinkIndex, editParams.Cell);
        var hyperlink = hyperlinks[index];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(editParams.Url))
        {
            hyperlink.Address = editParams.Url;
            changes.Add($"url={editParams.Url}");
        }

        if (!string.IsNullOrEmpty(editParams.DisplayText))
        {
            hyperlink.TextToDisplay = editParams.DisplayText;
            changes.Add($"displayText={editParams.DisplayText}");
        }

        var cellRef = CellsHelper.CellIndexToName(hyperlink.Area.StartRow, hyperlink.Area.StartColumn);

        if (changes.Count > 0)
        {
            MarkModified(context);
            return Success($"Hyperlink at {cellRef} edited: {string.Join(", ", changes)}.");
        }

        return Success($"Hyperlink at {cellRef} unchanged.");
    }

    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var url = parameters.GetOptional<string?>("url");
        var displayText = parameters.GetOptional<string?>("displayText");
        var hyperlinkIndex = parameters.GetOptional<int?>("hyperlinkIndex");

        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for edit operation.");

        return new EditParameters(sheetIndex, cell, url, displayText, hyperlinkIndex);
    }

    private sealed record EditParameters(
        int SheetIndex,
        string? Cell,
        string? Url,
        string? DisplayText,
        int? HyperlinkIndex);
}
