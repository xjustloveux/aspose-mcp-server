using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Handler for adding comments to Excel cells.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddExcelCommentHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a comment to a cell.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: cell, comment
    ///     Optional: sheetIndex, author
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        ExcelCommentHelper.ValidateCellAddress(addParams.Cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);
            var cellObj = worksheet.Cells[addParams.Cell];

            var commentObj = worksheet.Comments[cellObj.Name];
            if (commentObj == null)
            {
                var commentIndex = worksheet.Comments.Add(cellObj.Name);
                commentObj = worksheet.Comments[commentIndex];
            }

            commentObj.Note = addParams.Comment;
            commentObj.Author = addParams.Author ?? ExcelCommentHelper.DefaultAuthor;

            MarkModified(context);

            return new SuccessResult
                { Message = $"Comment added to cell {addParams.Cell} in sheet {addParams.SheetIndex}." };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{addParams.Cell}': {ex.Message}");
        }
    }

    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");
        var comment = parameters.GetOptional<string?>("comment");
        var author = parameters.GetOptional<string?>("author");

        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");
        if (string.IsNullOrEmpty(comment))
            throw new ArgumentException("comment is required for add operation");

        return new AddParameters(sheetIndex, cell, comment, author);
    }

    private sealed record AddParameters(int SheetIndex, string Cell, string Comment, string? Author);
}
