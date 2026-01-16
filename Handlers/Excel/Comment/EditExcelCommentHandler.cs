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
        var editParams = ExtractEditParameters(parameters);

        ExcelCommentHelper.ValidateCellAddress(editParams.Cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
            var cellObj = worksheet.Cells[editParams.Cell];
            var commentObj = worksheet.Comments[cellObj.Name];

            if (commentObj == null)
                throw new ArgumentException($"No comment found on cell {editParams.Cell}");

            commentObj.Note = editParams.Comment;
            if (!string.IsNullOrEmpty(editParams.Author))
                commentObj.Author = editParams.Author;

            MarkModified(context);

            return Success($"Comment edited on cell {editParams.Cell} in sheet {editParams.SheetIndex}.");
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{editParams.Cell}': {ex.Message}");
        }
    }

    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var comment = parameters.GetOptional<string?>("comment");
        var author = parameters.GetOptional<string?>("author");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for edit operation");
        if (string.IsNullOrEmpty(comment))
            throw new ArgumentException("comment is required for edit operation");

        return new EditParameters(sheetIndex, cell, comment, author);
    }

    private record EditParameters(int SheetIndex, string Cell, string Comment, string? Author);
}
