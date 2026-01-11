using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Handler for getting comments from Excel worksheets.
/// </summary>
public class GetExcelCommentsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all comments or comments for a specific cell.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, cell
    /// </param>
    /// <returns>JSON string containing the comment information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");

        if (!string.IsNullOrEmpty(cell))
            ExcelCommentHelper.ValidateCellAddress(cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (!string.IsNullOrEmpty(cell))
            {
                var comment = worksheet.Comments[cell];
                if (comment != null)
                {
                    var result = new
                    {
                        count = 1,
                        sheetIndex,
                        cell,
                        items = new[]
                        {
                            new
                            {
                                cell,
                                author = comment.Author,
                                note = comment.Note
                            }
                        }
                    };
                    return JsonResult(result);
                }
                else
                {
                    var result = new
                    {
                        count = 0,
                        sheetIndex,
                        cell,
                        items = Array.Empty<object>(),
                        message = $"No comment found on cell {cell}"
                    };
                    return JsonResult(result);
                }
            }

            if (worksheet.Comments.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    sheetIndex,
                    items = Array.Empty<object>(),
                    message = "No comments found"
                };
                return JsonResult(emptyResult);
            }

            List<object> commentList = [];
            foreach (var comment in worksheet.Comments)
            {
                var cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                commentList.Add(new
                {
                    cell = cellName,
                    author = comment.Author,
                    note = comment.Note
                });
            }

            var allResult = new
            {
                count = worksheet.Comments.Count,
                sheetIndex,
                items = commentList
            };
            return JsonResult(allResult);
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }
}
