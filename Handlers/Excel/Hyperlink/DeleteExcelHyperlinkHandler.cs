using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Handler for deleting hyperlinks from Excel worksheets.
/// </summary>
public class DeleteExcelHyperlinkHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a hyperlink from a cell.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: cell or hyperlinkIndex (at least one)
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        var index = ExcelHyperlinkHelper.ResolveHyperlinkIndex(hyperlinks, deleteParams.HyperlinkIndex,
            deleteParams.Cell);
        var cellRef = CellsHelper.CellIndexToName(hyperlinks[index].Area.StartRow, hyperlinks[index].Area.StartColumn);

        hyperlinks.RemoveAt(index);

        MarkModified(context);

        return Success($"Hyperlink at {cellRef} deleted. {hyperlinks.Count} hyperlinks remaining.");
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var hyperlinkIndex = parameters.GetOptional<int?>("hyperlinkIndex");

        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for delete operation.");

        return new DeleteParameters(sheetIndex, cell, hyperlinkIndex);
    }

    private record DeleteParameters(int SheetIndex, string? Cell, int? HyperlinkIndex);
}
