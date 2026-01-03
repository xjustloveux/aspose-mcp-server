using System.ComponentModel;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel comments (add, edit, delete, get)
/// </summary>
[McpServerToolType]
public class ExcelCommentTool
{
    /// <summary>
    ///     Default author name for comments.
    /// </summary>
    private const string DefaultAuthor = "AsposeMCP";

    /// <summary>
    ///     Regex pattern for validating Excel cell addresses (e.g., A1, B2, AA100).
    /// </summary>
    private static readonly Regex CellAddressRegex = new(@"^[A-Za-z]{1,3}\d+$", RegexOptions.Compiled);

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelCommentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelCommentTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_comment")]
    [Description(@"Manage Excel comments. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add comment: excel_comment(operation='add', path='book.xlsx', cell='A1', comment='This is a comment')
- Edit comment: excel_comment(operation='edit', path='book.xlsx', cell='A1', comment='Updated comment')
- Delete comment: excel_comment(operation='delete', path='book.xlsx', cell='A1')
- Get comments: excel_comment(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference (e.g., 'A1', required for add/edit/delete, optional for get)")]
        string? cell = null,
        [Description("Comment text (required for add/edit)")]
        string? comment = null,
        [Description("Comment author (optional)")]
        string? author = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddComment(ctx, outputPath, sheetIndex, cell, comment, author),
            "edit" => EditComment(ctx, outputPath, sheetIndex, cell, comment, author),
            "delete" => DeleteComment(ctx, outputPath, sheetIndex, cell),
            "get" => GetComments(ctx, sheetIndex, cell),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates the cell address format.
    /// </summary>
    /// <param name="cell">The cell address to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the cell address format is invalid.</exception>
    private static void ValidateCellAddress(string cell)
    {
        if (!CellAddressRegex.IsMatch(cell))
            throw new ArgumentException(
                $"Invalid cell address format: '{cell}'. Expected format like 'A1', 'B2', 'AA100'");
    }

    /// <summary>
    ///     Adds a comment to a cell.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address to add the comment to.</param>
    /// <param name="comment">The comment text.</param>
    /// <param name="author">The comment author.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when cell or comment is not provided, or the Excel operation fails.</exception>
    private static string AddComment(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? cell, string? comment, string? author)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");
        if (string.IsNullOrEmpty(comment))
            throw new ArgumentException("comment is required for add operation");

        ValidateCellAddress(cell);

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];

            var commentObj = worksheet.Comments[cellObj.Name];
            if (commentObj == null)
            {
                var commentIndex = worksheet.Comments.Add(cellObj.Name);
                commentObj = worksheet.Comments[commentIndex];
            }

            commentObj.Note = comment;
            commentObj.Author = author ?? DefaultAuthor;

            ctx.Save(outputPath);
            return $"Comment added to cell {cell} in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
        }
    }

    /// <summary>
    ///     Edits an existing cell comment.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address containing the comment to edit.</param>
    /// <param name="comment">The new comment text.</param>
    /// <param name="author">The new comment author.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when cell or comment is not provided, no comment exists on the cell, or the
    ///     Excel operation fails.
    /// </exception>
    private static string EditComment(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? cell, string? comment, string? author)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for edit operation");
        if (string.IsNullOrEmpty(comment))
            throw new ArgumentException("comment is required for edit operation");

        ValidateCellAddress(cell);

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];
            var commentObj = worksheet.Comments[cellObj.Name];

            if (commentObj == null) throw new ArgumentException($"No comment found on cell {cell}");

            commentObj.Note = comment;
            if (!string.IsNullOrEmpty(author)) commentObj.Author = author;

            ctx.Save(outputPath);
            return $"Comment edited on cell {cell} in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
        }
    }

    /// <summary>
    ///     Deletes a comment from a cell.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell address containing the comment to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when cell is not provided or the Excel operation fails.</exception>
    private static string DeleteComment(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? cell)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for delete operation");

        ValidateCellAddress(cell);

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var comment = worksheet.Comments[cell];

            if (comment != null) worksheet.Comments.RemoveAt(cell);

            ctx.Save(outputPath);
            return $"Comment deleted from cell {cell} in sheet {sheetIndex}. {ctx.GetOutputMessage(outputPath)}";
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed for cell '{cell}': {ex.Message}");
        }
    }

    /// <summary>
    ///     Gets all comments or comments for a specific cell.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The optional cell address to get comments for.</param>
    /// <returns>A JSON string containing the comment information.</returns>
    /// <exception cref="ArgumentException">Thrown when the Excel operation fails.</exception>
    private static string GetComments(DocumentContext<Workbook> ctx, int sheetIndex, string? cell)
    {
        if (!string.IsNullOrEmpty(cell))
            ValidateCellAddress(cell);

        try
        {
            var workbook = ctx.Document;
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
                    return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
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
                    return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
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
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
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
            return JsonSerializer.Serialize(allResult, new JsonSerializerOptions { WriteIndented = true });
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Excel operation failed: {ex.Message}");
        }
    }
}