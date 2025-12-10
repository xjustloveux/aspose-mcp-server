using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordDeleteTableTool : IAsposeTool
{
    public string Description => "Delete a specific table from Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
        {
            throw new ArgumentException($"節索引 {sectionIndex} 超出範圍 (文檔共有 {doc.Sections.Count} 個節)");
        }
        
        var section = doc.Sections[sectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        
        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new ArgumentException($"表格索引 {tableIndex} 超出範圍 (文檔共有 {tables.Count} 個表格)");
        }
        
        var tableToDelete = tables[tableIndex];
        
        // Get table info before deletion
        int rowCount = tableToDelete.Rows.Count;
        int colCount = tableToDelete.Rows.Count > 0 ? tableToDelete.Rows[0].Cells.Count : 0;
        
        // Delete the table
        tableToDelete.Remove();
        
        doc.Save(outputPath);
        
        var result = $"成功刪除表格 #{tableIndex}\n";
        result += $"表格大小: {rowCount} 行 x {colCount} 列\n";
        result += $"文檔剩餘表格數: {section.Body.GetChildNodes(NodeType.Table, true).Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

