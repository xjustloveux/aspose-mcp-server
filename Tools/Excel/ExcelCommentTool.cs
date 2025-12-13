using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel comments (add, edit, delete, get)
/// Merges: ExcelAddCommentTool, ExcelEditCommentTool, ExcelDeleteCommentTool, ExcelGetCommentsTool
/// </summary>
public class ExcelCommentTool : IAsposeTool
{
    public string Description => @"Manage Excel comments. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add comment: excel_comment(operation='add', path='book.xlsx', cell='A1', comment='This is a comment')
- Edit comment: excel_comment(operation='edit', path='book.xlsx', cell='A1', comment='Updated comment')
- Delete comment: excel_comment(operation='delete', path='book.xlsx', cell='A1')
- Get comments: excel_comment(operation='get', path='book.xlsx')";

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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddCommentAsync(arguments, path, sheetIndex),
            "edit" => await EditCommentAsync(arguments, path, sheetIndex),
            "delete" => await DeleteCommentAsync(arguments, path, sheetIndex),
            "get" => await GetCommentsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for add operation");
        var comment = arguments?["comment"]?.GetValue<string>() ?? throw new ArgumentException("comment is required for add operation");
        var author = arguments?["author"]?.GetValue<string>();

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
        if (!string.IsNullOrEmpty(author))
        {
            commentObj.Author = author;
        }

        workbook.Save(path);
        return await Task.FromResult($"Comment added to cell {cell} in sheet {sheetIndex}: {path}");
    }

    private async Task<string> EditCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for edit operation");
        var comment = arguments?["comment"]?.GetValue<string>() ?? throw new ArgumentException("comment is required for edit operation");
        var author = arguments?["author"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cellObj = worksheet.Cells[cell];
        var commentObj = worksheet.Comments[cellObj.Name];

        if (commentObj == null)
        {
            throw new ArgumentException($"No comment found on cell {cell}");
        }

        commentObj.Note = comment;
        if (!string.IsNullOrEmpty(author))
        {
            commentObj.Author = author;
        }

        workbook.Save(path);
        return await Task.FromResult($"Comment edited on cell {cell} in sheet {sheetIndex}: {path}");
    }

    private async Task<string> DeleteCommentAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required for delete operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var comment = worksheet.Comments[cell];

        if (comment != null)
        {
            worksheet.Comments.RemoveAt(cell);
        }

        workbook.Save(path);
        return await Task.FromResult($"Comment deleted from cell {cell} in sheet {sheetIndex}: {path}");
    }

    private async Task<string> GetCommentsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var cell = arguments?["cell"]?.GetValue<string>();

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
            {
                for (int i = 0; i < worksheet.Comments.Count; i++)
                {
                    var comment = worksheet.Comments[i];
                    var cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                    sb.AppendLine($"  Cell {cellName}:");
                    sb.AppendLine($"    Author: {comment.Author}");
                    sb.AppendLine($"    Note: {comment.Note}");
                    sb.AppendLine();
                }
            }
            else
            {
                sb.AppendLine("  No comments found");
            }
        }

        return await Task.FromResult(sb.ToString());
    }
}

