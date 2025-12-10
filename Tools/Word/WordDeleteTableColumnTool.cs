using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordDeleteTableColumnTool : IAsposeTool
{
    public string Description => "Delete a specific column from a table";

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
            columnIndex = new
            {
                type = "number",
                description = "Column index to delete (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "columnIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required");
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
        
        var table = tables[tableIndex];
        
        if (table.Rows.Count == 0)
        {
            throw new InvalidOperationException($"表格 {tableIndex} 沒有行");
        }
        
        var firstRow = table.Rows[0];
        if (columnIndex < 0 || columnIndex >= firstRow.Cells.Count)
        {
            throw new ArgumentException($"列索引 {columnIndex} 超出範圍 (表格共有 {firstRow.Cells.Count} 列)");
        }
        
        // Get column info before deletion (for confirmation message)
        string columnPreview = "";
        try
        {
            var firstCellText = firstRow.Cells[columnIndex].GetText().Trim();
            if (firstCellText.Length > 30)
            {
                columnPreview = firstCellText.Substring(0, 30) + "...";
            }
            else
            {
                columnPreview = firstCellText;
            }
        }
        catch
        {
            // Ignore errors when getting column preview
        }
        
        // Delete the column (remove cell from each row)
        int deletedCount = 0;
        foreach (Row row in table.Rows)
        {
            if (columnIndex < row.Cells.Count)
            {
                row.Cells[columnIndex].Remove();
                deletedCount++;
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功刪除列 #{columnIndex}\n";
        result += $"表格: {tableIndex}\n";
        if (!string.IsNullOrEmpty(columnPreview))
        {
            result += $"內容預覽: {columnPreview}\n";
        }
        result += $"刪除單元格數: {deletedCount}\n";
        if (table.Rows.Count > 0)
        {
            result += $"表格剩餘列數: {table.Rows[0].Cells.Count}\n";
        }
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

