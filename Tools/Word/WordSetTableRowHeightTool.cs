using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordSetTableRowHeightTool : IAsposeTool
{
    public string Description => "Set the height of a specific table row";

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
                description = "Row index (0-based)"
            },
            height = new
            {
                type = "number",
                description = "Row height in points"
            },
            heightRule = new
            {
                type = "string",
                description = "Height rule: auto, atLeast, exactly (default: atLeast)",
                @enum = new[] { "auto", "atLeast", "exactly" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "rowIndex", "height" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");
        var heightRule = arguments?["heightRule"]?.GetValue<string>() ?? "atLeast";
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        if (height <= 0)
        {
            throw new ArgumentException($"行高 {height} 必須大於 0");
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
        
        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
        {
            throw new ArgumentException($"行索引 {rowIndex} 超出範圍 (表格共有 {table.Rows.Count} 行)");
        }
        
        var row = table.Rows[rowIndex];
        var rowFormat = row.RowFormat;
        
        // Set height rule
        rowFormat.HeightRule = heightRule.ToLower() switch
        {
            "auto" => HeightRule.Auto,
            "atLeast" => HeightRule.AtLeast,
            "exactly" => HeightRule.Exactly,
            _ => HeightRule.AtLeast
        };
        
        // Set height
        rowFormat.Height = height;
        
        doc.Save(outputPath);
        
        var result = $"成功設定行高\n";
        result += $"表格: {tableIndex}\n";
        result += $"行索引: {rowIndex}\n";
        result += $"行高: {height} pt\n";
        result += $"高度規則: {heightRule}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

