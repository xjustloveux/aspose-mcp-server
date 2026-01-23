using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Comment;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Handler for getting comments from Excel worksheets.
/// </summary>
[ResultType(typeof(GetCommentsExcelResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
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
                    return new GetCommentsExcelResult
                    {
                        Count = 1,
                        SheetIndex = getParams.SheetIndex,
                        Cell = getParams.Cell,
                        Items =
                        [
                            new ExcelCommentInfo
                            {
                                Cell = getParams.Cell,
                                Author = comment.Author,
                                Note = comment.Note
                            }
                        ]
                    };

                return new GetCommentsExcelResult
                {
                    Count = 0,
                    SheetIndex = getParams.SheetIndex,
                    Cell = getParams.Cell,
                    Items = [],
                    Message = $"No comment found on cell {getParams.Cell}"
                };
            }

            if (worksheet.Comments.Count == 0)
                return new GetCommentsExcelResult
                {
                    Count = 0,
                    SheetIndex = getParams.SheetIndex,
                    Items = [],
                    Message = "No comments found"
                };

            List<ExcelCommentInfo> commentList = [];
            foreach (var comment in worksheet.Comments)
            {
                var cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                commentList.Add(new ExcelCommentInfo
                {
                    Cell = cellName,
                    Author = comment.Author,
                    Note = comment.Note
                });
            }

            return new GetCommentsExcelResult
            {
                Count = worksheet.Comments.Count,
                SheetIndex = getParams.SheetIndex,
                Items = commentList
            };
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

    private sealed record GetParameters(int SheetIndex, string? Cell);
}
