using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Handler for editing existing comments in Excel cells.
/// </summary>
public class EditExcelCommentHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing cell comment.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: cell, comment
    ///     Optional: sheetIndex, author
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var comment = parameters.GetOptional<string?>("comment");
        var author = parameters.GetOptional<string?>("author");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for edit operation");
        if (string.IsNullOrEmpty(comment))
            throw new ArgumentException("comment is required for edit operation");

        ExcelCommentHelper.ValidateCellAddress(cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];
            var commentObj = worksheet.Comments[cellObj.Name];

            if (commentObj == null)
                throw new ArgumentException($"No comment found on cell {cell}");

            commentObj.Note = comment;
            if (!string.IsNullOrEmpty(author))
                commentObj.Author = author;

            MarkModified(context);

            return Success($"Comment edited on cell {cell} in sheet {sheetIndex}.");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
        }
    }
}
