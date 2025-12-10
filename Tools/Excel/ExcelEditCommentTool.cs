using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditCommentTool : IAsposeTool
{
    public string Description => "Edit comment on a cell in Excel";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description = "Cell reference (e.g., 'A1')"
            },
            comment = new
            {
                type = "string",
                description = "New comment text"
            },
            author = new
            {
                type = "string",
                description = "Comment author (optional)"
            }
        },
        required = new[] { "path", "cell", "comment" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>() ?? throw new ArgumentException("cell is required");
        var comment = arguments?["comment"]?.GetValue<string>() ?? throw new ArgumentException("comment is required");
        var author = arguments?["author"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
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
}

