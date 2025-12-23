using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel comments (add, edit, delete, get)
///     Merges: ExcelAddCommentTool, ExcelEditCommentTool, ExcelDeleteCommentTool, ExcelGetCommentsTool
/// </summary>
public class ExcelCommentTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel comments. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add comment: excel_comment(operation='add', path='book.xlsx', cell='A1', comment='This is a comment')
- Edit comment: excel_comment(operation='edit', path='book.xlsx', cell='A1', comment='Updated comment')
- Delete comment: excel_comment(operation='delete', path='book.xlsx', cell='A1')
- Get comments: excel_comment(operation='get', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a comment (required params: path, cell, comment)
- 'edit': Edit a comment (required params: path, cell, comment)
- 'delete': Delete a comment (required params: path, cell)
- 'get': Get all comments (required params: path)",
                @enum = new[] { "add", "edit", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1', required for add/edit/delete, optional for get)"
            },
            comment = new
            {
                type = "string",
                description = "Comment text (required for add/edit)"
            },
            author = new
            {
                type = "string",
                description = "Comment author (optional)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddCommentAsync(arguments, path, sheetIndex),
            "edit" => await EditCommentAsync(arguments, path, sheetIndex),
            "delete" => await DeleteCommentAsync(arguments, path, sheetIndex),
            "get" => await GetCommentsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a comment to a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell address and comment text</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with comment details</returns>
    private Task<string> AddCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var comment = ArgumentHelper.GetString(arguments, "comment");
            var author = ArgumentHelper.GetStringNullable(arguments, "author");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];

            var commentObj = worksheet.Comments[cellObj.Name];
            if (commentObj == null)
            {
                var commentIndex = worksheet.Comments.Add(cellObj.Name);
                commentObj = worksheet.Comments[commentIndex];
            }

            commentObj.Note = comment;
            if (!string.IsNullOrEmpty(author)) commentObj.Author = author;

            workbook.Save(outputPath);
            return $"Comment added to cell {cell} in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing cell comment
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell address and new comment text</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> EditCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var comment = ArgumentHelper.GetString(arguments, "comment");
            var author = ArgumentHelper.GetStringNullable(arguments, "author");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var cellObj = worksheet.Cells[cell];
            var commentObj = worksheet.Comments[cellObj.Name];

            if (commentObj == null) throw new ArgumentException($"No comment found on cell {cell}");

            commentObj.Note = comment;
            if (!string.IsNullOrEmpty(author)) commentObj.Author = author;

            workbook.Save(outputPath);
            return $"Comment edited on cell {cell} in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a comment from a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell address</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var cell = ArgumentHelper.GetString(arguments, "cell");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var comment = worksheet.Comments[cell];

            if (comment != null) worksheet.Comments.RemoveAt(cell);

            workbook.Save(outputPath);
            return $"Comment deleted from cell {cell} in sheet {sheetIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all comments or comments for a specific cell
    /// </summary>
    /// <param name="arguments">JSON arguments optionally containing cell address</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with comment information</returns>
    private Task<string> GetCommentsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var sb = new StringBuilder();

            if (!string.IsNullOrEmpty(cell))
            {
                var comment = worksheet.Comments[cell];
                if (comment != null)
                {
                    sb.AppendLine($"Comment on cell {cell}:");
                    sb.AppendLine($"  Author: {comment.Author}");
                    sb.AppendLine($"  Note: {comment.Note}");
                }
                else
                {
                    sb.AppendLine($"No comment found on cell {cell}");
                }
            }
            else
            {
                sb.AppendLine($"Comments in sheet {sheetIndex}:");
                if (worksheet.Comments.Count > 0)
                    foreach (var comment in worksheet.Comments)
                    {
                        var cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                        sb.AppendLine($"  Cell {cellName}:");
                        sb.AppendLine($"    Author: {comment.Author}");
                        sb.AppendLine($"    Note: {comment.Note}");
                        sb.AppendLine();
                    }
                else
                    sb.AppendLine("  No comments found");
            }

            return sb.ToString();
        });
    }
}