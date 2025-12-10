using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordInsertTableColumnTool : IAsposeTool
{
    public string Description => "Insert a new column into an existing table";

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
                description = "Column index to insert at (0-based). If insertBefore is false, column will be inserted after this index."
            },
            insertBefore = new
            {
                type = "boolean",
                description = "If true, insert before the specified column. If false, insert after. Default: false"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            data = new
            {
                type = "array",
                description = "Optional array of cell data for the new column (one value per row)",
                items = new { type = "string" }
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
        
        if (table.Rows.Count == 0)
        {
            throw new InvalidOperationException($"表格 {tableIndex} 沒有行");
        }
        
        var firstRow = table.Rows[0];
        if (columnIndex < 0 || columnIndex >= firstRow.Cells.Count)
        {
            throw new ArgumentException($"列索引 {columnIndex} 超出範圍 (表格共有 {firstRow.Cells.Count} 列)");
        }
        
        // Insert a new cell in each row
        int insertPosition = insertBefore ? columnIndex : columnIndex + 1;
        
        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var row = table.Rows[rowIdx];
            
            // Create new cell
            Cell newCell = new Cell(doc);
            
            // Copy formatting from adjacent cell if possible
            if (columnIndex < row.Cells.Count)
            {
                var sourceCell = row.Cells[columnIndex];
                newCell.CellFormat.Width = sourceCell.CellFormat.Width;
                newCell.CellFormat.VerticalAlignment = sourceCell.CellFormat.VerticalAlignment;
                newCell.CellFormat.Shading.BackgroundPatternColor = sourceCell.CellFormat.Shading.BackgroundPatternColor;
                newCell.CellFormat.SetPaddings(
                    sourceCell.CellFormat.TopPadding,
                    sourceCell.CellFormat.BottomPadding,
                    sourceCell.CellFormat.LeftPadding,
                    sourceCell.CellFormat.RightPadding
                );
            }
            
            // Add data if provided
            if (dataArray != null && rowIdx < dataArray.Count)
            {
                var cellText = dataArray[rowIdx]?.GetValue<string>() ?? "";
                if (!string.IsNullOrEmpty(cellText))
                {
                    Paragraph para = new Paragraph(doc);
                    Run run = new Run(doc, cellText);
                    para.AppendChild(run);
                    newCell.AppendChild(para);
                }
            }
            else
            {
                // Add empty paragraph
                Paragraph para = new Paragraph(doc);
                newCell.AppendChild(para);
            }
            
            // Insert the cell
            if (insertPosition <= row.Cells.Count)
            {
                row.Cells.Insert(insertPosition, newCell);
            }
            else
            {
                row.AppendChild(newCell);
            }
        }
        
        doc.Save(outputPath);
        
        var position = insertBefore ? $"列 #{columnIndex} 之前" : $"列 #{columnIndex} 之後";
        var result = $"成功插入新列\n";
        result += $"表格: {tableIndex}\n";
        result += $"插入位置: {position}\n";
        result += $"新列索引: {(insertBefore ? columnIndex : columnIndex + 1)}\n";
        result += $"表格列數: {table.Rows[0].Cells.Count}\n";
        result += $"輸出: {outputPath}";
        
        return await Task.FromResult(result);
    }
}

