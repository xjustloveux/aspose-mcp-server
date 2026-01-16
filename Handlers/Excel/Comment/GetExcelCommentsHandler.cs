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
        var getParams = ExtractGetParameters(parameters);

        if (!string.IsNullOrEmpty(getParams.Cell))
            ExcelCommentHelper.ValidateCellAddress(getParams.Cell);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);

            if (!string.IsNullOrEmpty(getParams.Cell))
            {
                var comment = worksheet.Comments[getParams.Cell];
                if (comment != null)
                {
                    var result = new
                    {
                        count = 1,
                        sheetIndex = getParams.SheetIndex,
                        cell = getParams.Cell,
                        items = new[]
                        {
                            new
                            {
                                cell = getParams.Cell,
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
                        sheetIndex = getParams.SheetIndex,
                        cell = getParams.Cell,
                        items = Array.Empty<object>(),
                        message = $"No comment found on cell {getParams.Cell}"
                    };
                    return JsonResult(result);
                }
            }

            if (worksheet.Comments.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    sheetIndex = getParams.SheetIndex,
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
                sheetIndex = getParams.SheetIndex,
                items = commentList
            };
            return JsonResult(allResult);
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var cell = parameters.GetOptional<string?>("cell");

        return new GetParameters(sheetIndex, cell);
    }

    private record GetParameters(int SheetIndex, string? Cell);
}
