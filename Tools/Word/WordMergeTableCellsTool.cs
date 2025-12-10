using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordMergeTableCellsTool : IAsposeTool
{
    public string Description => "Merge table cells in a specified range (independent tool for merging cells in existing tables)";

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
            startRow = new
            {
                type = "number",
                description = "Start row index (0-based)"
            },
            startCol = new
            {
                type = "number",
                description = "Start column index (0-based)"
            },
            endRow = new
            {
                type = "number",
                description = "End row index (0-based, inclusive)"
            },
            endCol = new
            {
                type = "number",
                description = "End column index (0-based, inclusive)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "startRow", "startCol", "endRow", "endCol" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var startRow = arguments?["startRow"]?.GetValue<int>() ?? throw new ArgumentException("startRow is required");
        var startCol = arguments?["startCol"]?.GetValue<int>() ?? throw new ArgumentException("startCol is required");
        var endRow = arguments?["endRow"]?.GetValue<int>() ?? throw new ArgumentException("endRow is required");
        var endCol = arguments?["endCol"]?.GetValue<int>() ?? throw new ArgumentException("endCol is required");
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
        
        // Validate indices
        if (startRow < 0 || startRow >= table.Rows.Count)
        {
            throw new ArgumentException($"起始行索引 {startRow} 超出範圍 (表格共有 {table.Rows.Count} 行)");
        }
        
        if (endRow < 0 || endRow >= table.Rows.Count)
        {
            throw new ArgumentException($"結束行索引 {endRow} 超出範圍 (表格共有 {table.Rows.Count} 行)");
        }
        
        if (startRow > endRow)
        {
            throw new ArgumentException($"起始行索引 {startRow} 不能大於結束行索引 {endRow}");
        }
        
        var firstRow = table.Rows[startRow];
        if (startCol < 0 || startCol >= firstRow.Cells.Count)
        {
            throw new ArgumentException($"起始列索引 {startCol} 超出範圍 (該行共有 {firstRow.Cells.Count} 列)");
        }
        
        if (endCol < 0 || endCol >= firstRow.Cells.Count)
        {
            throw new ArgumentException($"結束列索引 {endCol} 超出範圍 (該行共有 {firstRow.Cells.Count} 列)");
        }
        
        if (startCol > endCol)
        {
            throw new ArgumentException($"起始列索引 {startCol} 不能大於結束列索引 {endCol}");
        }
        
        // Get the start cell (this will be the merged cell)
        var startCell = table.Rows[startRow].Cells[startCol];
        
        // Merge cells: set all cells except the start cell to merge with previous/above
        for (int row = startRow; row <= endRow; row++)
        {
            var currentRow = table.Rows[row];
            
            for (int col = startCol; col <= endCol; col++)
            {
                var cell = currentRow.Cells[col];
                
                if (row == startRow && col == startCol)
                {
                    // This is the start cell - set as first for both horizontal and vertical merge
                    if (startRow != endRow)
                    {
                        cell.CellFormat.VerticalMerge = CellMerge.First;
                    }
                    if (startCol != endCol)
                    {
                        cell.CellFormat.HorizontalMerge = CellMerge.First;
                    }
                }
                else
                {
                    // Set as previous merge
                    if (row == startRow)
                    {
                        // Same row, different column - horizontal merge
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                    }
                    else if (col == startCol)
                    {
                        // Same column, different row - vertical merge
                        cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    else
                    {
                        // Both different - need both horizontal and vertical merge
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                        cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
            }
        }
        
        doc.Save(outputPath);
        
        var result = $"成功合併單元格\n";
        result += $"表格: {tableIndex}\n";
        result += $"合併範圍: [{startRow}, {startCol}] 到 [{endRow}, {endCol}]\n";
        result += $"合併單元格數: {(endRow - startRow + 1) * (endCol - startCol + 1)}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

