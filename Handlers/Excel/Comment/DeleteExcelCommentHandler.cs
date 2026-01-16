using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Handler for deleting comments from Excel cells.
/// </summary>
public class DeleteExcelCommentHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a comment from a cell.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: cell
    ///     Optional: sheetIndex
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        ExcelCommentHelper.ValidateCellAddress(deleteParams.Cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
            var comment = worksheet.Comments[deleteParams.Cell];

            if (comment != null)
                worksheet.Comments.RemoveAt(deleteParams.Cell);

            MarkModified(context);

            return Success($"Comment deleted from cell {deleteParams.Cell} in sheet {deleteParams.SheetIndex}.");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{deleteParams.Cell}': {ex.Message}");
        }
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for delete operation");

        return new DeleteParameters(sheetIndex, cell);
    }

    private sealed record DeleteParameters(int SheetIndex, string Cell);
}
