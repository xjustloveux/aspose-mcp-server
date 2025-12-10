using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordDeleteTableRowTool : IAsposeTool
{
    public string Description => "Delete a specific row from a table";

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
            rowIndex = new
            {
                type = "number",
                description = "Row index to delete (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "rowIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
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
        
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
        {
            throw new ArgumentException($"行索引 {rowIndex} 超出範圍 (表格共有 {table.Rows.Count} 行)");
        }
        
        var rowToDelete = table.Rows[rowIndex];
        
        // Get row info before deletion (for confirmation message)
        int cellCount = rowToDelete.Cells.Count;
        string rowPreview = "";
        try
        {
            var firstCellText = rowToDelete.Cells[0].GetText().Trim();
            if (firstCellText.Length > 30)
            {
                rowPreview = firstCellText.Substring(0, 30) + "...";
            }
            else
            {
                rowPreview = firstCellText;
            }
        }
        catch
        {
            // Ignore errors when getting row preview
        }
        
        // Delete the row
        rowToDelete.Remove();
        
        doc.Save(outputPath);
        
        var result = $"成功刪除行 #{rowIndex}\n";
        result += $"表格: {tableIndex}\n";
        if (!string.IsNullOrEmpty(rowPreview))
        {
            result += $"內容預覽: {rowPreview}\n";
        }
        result += $"該行單元格數: {cellCount}\n";
        result += $"表格剩餘行數: {table.Rows.Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

