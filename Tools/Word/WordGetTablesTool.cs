using System.Text.Json.Nodes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordGetTablesTool : IAsposeTool
{
    public string Description => "Get all tables from Word document with structure information";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, optional, if not provided returns all tables)"
            },
            includeContent = new
            {
                type = "boolean",
                description = "Include table content (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int?>();
        var includeContent = arguments?["includeContent"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var sb = new StringBuilder();

        List<Aspose.Words.Tables.Table> tables;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
            {
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
            }
            tables = doc.Sections[sectionIndex.Value].Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        }
        else
        {
            tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }

        sb.AppendLine($"=== Tables ({tables.Count}) ===");
        sb.AppendLine();

        for (int i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            sb.AppendLine($"[{i}] Rows: {table.Rows.Count}, Columns: {table.FirstRow?.Cells?.Count ?? 0}");
            sb.AppendLine($"    Style: {table.Style?.Name ?? "(none)"}");
            
            if (includeContent)
            {
                sb.AppendLine("    Content:");
                for (int row = 0; row < Math.Min(3, table.Rows.Count); row++)
                {
                    var rowText = string.Join(" | ", table.Rows[row].Cells.Cast<Aspose.Words.Tables.Cell>().Select(c => c.GetText().Trim().Substring(0, Math.Min(20, c.GetText().Trim().Length))));
                    sb.AppendLine($"      {rowText}");
                }
                if (table.Rows.Count > 3)
                {
                    sb.AppendLine($"      ... ({table.Rows.Count - 3} more rows)");
                }
            }
            sb.AppendLine();
        }

        return await Task.FromResult(sb.ToString());
    }
}

