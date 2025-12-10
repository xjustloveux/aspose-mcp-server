using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetCommentsTool : IAsposeTool
{
    public string Description => "Get all comments from Excel worksheet or specific cell";

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
                description = "Cell reference (e.g., 'A1', optional, if not provided returns all comments)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var cell = arguments?["cell"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"sheetIndex must be between 0 and {workbook.Worksheets.Count - 1}");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
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

