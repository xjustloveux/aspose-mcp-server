using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordSetTableColumnWidthTool : IAsposeTool
{
    public string Description => "Set the width of a specific table column";

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
                description = "Column index (0-based)"
            },
            width = new
            {
                type = "number",
                description = "Column width in points"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "columnIndex", "width" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        if (width <= 0)
        {
            throw new ArgumentException($"列寬 {width} 必須大於 0");
        }

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
        
        // Set width for all cells in the column
        int cellsUpdated = 0;
        foreach (Row row in table.Rows)
        {
            if (columnIndex < row.Cells.Count)
            {
                var cell = row.Cells[columnIndex];
                cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(width);
                cellsUpdated++;
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功設定列寬\n";
        result += $"表格: {tableIndex}\n";
        result += $"列索引: {columnIndex}\n";
        result += $"列寬: {width} pt\n";
        result += $"更新單元格數: {cellsUpdated}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

