using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordInsertTableRowTool : IAsposeTool
{
    public string Description => "Insert a new row into an existing table";

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
                description = "Row index to insert at (0-based). If insertBefore is false, row will be inserted after this index."
            },
            insertBefore = new
            {
                type = "boolean",
                description = "If true, insert before the specified row. If false, insert after. Default: false"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            data = new
            {
                type = "array",
                description = "Optional array of cell data for the new row",
                items = new { type = "string" }
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
        var insertBefore = arguments?["insertBefore"]?.GetValue<bool>() ?? false;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var dataArray = arguments?["data"]?.AsArray();

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
        
        var targetRow = table.Rows[rowIndex];
        int columnCount = targetRow.Cells.Count;
        
        // Create new row
        Row newRow = new Row(doc);
        
        // Add cells to match column count
        for (int i = 0; i < columnCount; i++)
        {
            Cell newCell = new Cell(doc);
            newRow.AppendChild(newCell);
            
            // Add data if provided
            if (dataArray != null && i < dataArray.Count)
            {
                var cellText = dataArray[i]?.GetValue<string>() ?? "";
                if (!string.IsNullOrEmpty(cellText))
                {
                    Paragraph para = new Paragraph(doc);
                    Run run = new Run(doc, cellText);
                    para.AppendChild(run);
                    newCell.AppendChild(para);
                }
            }
        }
        
        // Insert the row
        if (insertBefore)
        {
            table.InsertBefore(newRow, targetRow);
        }
        else
        {
            table.InsertAfter(newRow, targetRow);
        }
        
        doc.Save(outputPath);
        
        var position = insertBefore ? $"行 #{rowIndex} 之前" : $"行 #{rowIndex} 之後";
        var result = $"成功插入新行\n";
        result += $"表格: {tableIndex}\n";
        result += $"插入位置: {position}\n";
        result += $"新行索引: {(insertBefore ? rowIndex : rowIndex + 1)}\n";
        result += $"表格行數: {table.Rows.Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

