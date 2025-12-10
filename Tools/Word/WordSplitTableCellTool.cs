using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeMcpServer.Tools;

public class WordSplitTableCellTool : IAsposeTool
{
    public string Description => "Split a table cell into multiple rows and columns";

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
            colIndex = new
            {
                type = "number",
                description = "Column index (0-based)"
            },
            rows = new
            {
                type = "number",
                description = "Number of rows to split into (default: 2)"
            },
            cols = new
            {
                type = "number",
                description = "Number of columns to split into (default: 2)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            }
        },
        required = new[] { "path", "tableIndex", "rowIndex", "colIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var colIndex = arguments?["colIndex"]?.GetValue<int>() ?? throw new ArgumentException("colIndex is required");
        var rows = arguments?["rows"]?.GetValue<int>() ?? 2;
        var cols = arguments?["cols"]?.GetValue<int>() ?? 2;
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        if (rows < 1 || rows > 10)
        {
            throw new ArgumentException($"行數 {rows} 超出範圍 (必須在 1-10 之間)");
        }
        
        if (cols < 1 || cols > 10)
        {
            throw new ArgumentException($"列數 {cols} 超出範圍 (必須在 1-10 之間)");
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
        
        if (colIndex < 0 || colIndex >= row.Cells.Count)
        {
            throw new ArgumentException($"列索引 {colIndex} 超出範圍 (該行共有 {row.Cells.Count} 列)");
        }
        
        var cell = row.Cells[colIndex];
        
        // Check if cell is merged
        bool isMerged = cell.CellFormat.HorizontalMerge != CellMerge.None || 
                       cell.CellFormat.VerticalMerge != CellMerge.None;
        
        if (isMerged)
        {
            throw new InvalidOperationException("無法拆分合併的單元格。請先使用 word_merge_table_cells 取消合併，或直接編輯單元格。");
        }
        
        // Split the cell using Aspose.Words API
        // Note: Aspose.Words doesn't have a direct Split method, so we need to manually create new cells
        // This is a simplified implementation that creates new cells in the same row
        
        try
        {
            // Get the cell's content
            string cellText = cell.GetText();
            
            // Remove the original cell
            var parentRow = cell.ParentRow;
            var cellIndex = parentRow.Cells.IndexOf(cell);
            
            // Create new cells
            for (int c = 0; c < cols; c++)
            {
                Cell newCell = new Cell(doc);
                
                // Copy formatting from original cell
                newCell.CellFormat.Width = cell.CellFormat.Width / cols;
                newCell.CellFormat.VerticalAlignment = cell.CellFormat.VerticalAlignment;
                newCell.CellFormat.Shading.BackgroundPatternColor = cell.CellFormat.Shading.BackgroundPatternColor;
                newCell.CellFormat.SetPaddings(
                    cell.CellFormat.TopPadding,
                    cell.CellFormat.BottomPadding,
                    cell.CellFormat.LeftPadding,
                    cell.CellFormat.RightPadding
                );
                
                // Add paragraph with text (split text if multiple cells)
                Paragraph para = new Paragraph(doc);
                if (cols == 1)
                {
                    // Single column - keep all text
                    Run run = new Run(doc, cellText);
                    para.AppendChild(run);
                }
                else
                {
                    // Multiple columns - split text or leave empty
                    if (c == 0 && !string.IsNullOrEmpty(cellText))
                    {
                        Run run = new Run(doc, cellText);
                        para.AppendChild(run);
                    }
                }
                newCell.AppendChild(para);
                
                // Insert the new cell
                if (c == 0)
                {
                    parentRow.Cells[cellIndex].Remove();
                    parentRow.Cells.Insert(cellIndex, newCell);
                }
                else
                {
                    parentRow.Cells.Insert(cellIndex + c, newCell);
                }
            }
            
            // If rows > 1, we need to add rows below and split vertically
            // This is more complex and requires handling merged cells properly
            if (rows > 1)
            {
                // For vertical split, we need to add new rows
                // This is a simplified version - for full implementation, 
                // we'd need to handle the merged cell structure more carefully
                for (int r = 1; r < rows; r++)
                {
                    // Find the row to insert after
                    int insertAfterRowIndex = rowIndex + r - 1;
                    if (insertAfterRowIndex < table.Rows.Count)
                    {
                        Row newRow = new Row(doc);
                        
                        // Copy structure from the row with split cell
                        var sourceRow = table.Rows[rowIndex];
                        for (int c = 0; c < sourceRow.Cells.Count; c++)
                        {
                            Cell newCell = new Cell(doc);
                            
                            // Copy formatting
                            var sourceCell = sourceRow.Cells[c];
                            newCell.CellFormat.Width = sourceCell.CellFormat.Width;
                            newCell.CellFormat.VerticalAlignment = sourceCell.CellFormat.VerticalAlignment;
                            newCell.CellFormat.Shading.BackgroundPatternColor = sourceCell.CellFormat.Shading.BackgroundPatternColor;
                            newCell.CellFormat.SetPaddings(
                                sourceCell.CellFormat.TopPadding,
                                sourceCell.CellFormat.BottomPadding,
                                sourceCell.CellFormat.LeftPadding,
                                sourceCell.CellFormat.RightPadding
                            );
                            
                            // Add empty paragraph
                            Paragraph para = new Paragraph(doc);
                            newCell.AppendChild(para);
                            
                            newRow.AppendChild(newCell);
                        }
                        
                        table.InsertAfter(newRow, table.Rows[insertAfterRowIndex]);
                    }
                }
            }
            
            doc.Save(outputPath);
            
            var result = $"成功拆分單元格\n";
            result += $"表格: {tableIndex}\n";
            result += $"單元格位置: [{rowIndex}, {colIndex}]\n";
            result += $"拆分為: {rows} 行 x {cols} 列\n";
            result += $"輸出: {outputPath}";
            
            return await Task.FromResult(result);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"拆分單元格時發生錯誤: {ex.Message}", ex);
        }
    }
}

