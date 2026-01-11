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
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for delete operation");

        ExcelCommentHelper.ValidateCellAddress(cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var comment = worksheet.Comments[cell];

            if (comment != null)
                worksheet.Comments.RemoveAt(cell);

            MarkModified(context);

            return Success($"Comment deleted from cell {cell} in sheet {sheetIndex}.");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
        }
    }
}
