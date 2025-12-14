using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel merged cells (merge/unmerge/get)
/// Merges: ExcelMergeCellsTool, ExcelGetMergedCellsTool
/// </summary>
public class ExcelMergeCellsTool : IAsposeTool
{
    public string Description => @"Manage Excel merged cells. Supports 3 operations: merge, unmerge, get.

Usage examples:
- Merge cells: excel_merge_cells(operation='merge', path='book.xlsx', range='A1:C1')
- Unmerge cells: excel_merge_cells(operation='unmerge', path='book.xlsx', range='A1:C1')
- Get merged cells: excel_merge_cells(operation='get', path='book.xlsx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'merge': Merge cells (required params: path, range)
- 'unmerge': Unmerge cells (required params: path, range)
- 'get': Get merged cells info (required params: path)",
                @enum = new[] { "merge", "unmerge", "get" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            range = new
            {
                type = "string",
                description = "Cell range to merge/unmerge (e.g., 'A1:C3', required for merge/unmerge)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "merge" => await MergeCellsAsync(arguments, path, sheetIndex),
            "unmerge" => await UnmergeCellsAsync(arguments, path, sheetIndex),
            "get" => await GetMergedCellsAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Merges cells in a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with merged range</returns>
    private async Task<string> MergeCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        cellRange.Merge();
        workbook.Save(path);
        return await Task.FromResult($"範圍 {range} 已合併: {path}");
    }

    /// <summary>
    /// Unmerges cells in a range
    /// </summary>
    /// <param name="arguments">JSON arguments containing range</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> UnmergeCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var range = ArgumentHelper.GetString(arguments, "range", "range");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];
        var cells = worksheet.Cells;
        var cellRange = cells.CreateRange(range);

        cellRange.UnMerge();
        workbook.Save(path);
        return await Task.FromResult($"範圍 {range} 已取消合併: {path}");
    }

    /// <summary>
    /// Gets all merged cell ranges from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all merged ranges</returns>
    private async Task<string> GetMergedCellsAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var mergedCells = worksheet.Cells.MergedCells;
        if (mergedCells == null)
        {
            throw new InvalidOperationException($"無法取得合併單元格資訊：{worksheet.Name}");
        }
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的合併單元格資訊 ===\n");
        result.AppendLine($"總合併區域數: {mergedCells.Count}\n");

        if (mergedCells.Count == 0)
        {
            result.AppendLine("未找到合併單元格");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < mergedCells.Count; i++)
        {
            try
            {
                // MergedCells[i] returns a CellArea object or its string representation
                // Format might be "A1:B2" or "Aspose.Cells.CellArea(A5:C5)[4,0,4,2]"
                string? mergedCellName = null;
                int startRow = -1, startCol = -1, endRow = -1, endCol = -1;
                
                try
                {
                    // MergedCells[i] returns a CellArea struct or its string representation
                    // Try to get as CellArea struct first
                    object? mergedCellObj = mergedCells[i];
                    CellArea cellArea;
                    
                    if (mergedCellObj is CellArea)
                    {
                        cellArea = (CellArea)mergedCellObj;
                        // Use CellArea properties directly
                        startRow = cellArea.StartRow;
                        startCol = cellArea.StartColumn;
                        endRow = cellArea.EndRow;
                        endCol = cellArea.EndColumn;
                        
                        // Convert to cell names
                        var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                        var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                        mergedCellName = $"{startCellName}:{endCellName}";
                    }
                    else
                    {
                        // Try as string representation
                        mergedCellName = mergedCells[i]?.ToString();
                        
                        // Parse string format
                        // Format might be "A1:B2" or "Aspose.Cells.CellArea(A5:C5)[4,0,4,2]"
                        if (!string.IsNullOrWhiteSpace(mergedCellName))
                        {
                            // Check if it's CellArea format: "Aspose.Cells.CellArea(A5:C5)[4,0,4,2]"
                            if (mergedCellName.Contains("CellArea(") && mergedCellName.Contains("["))
                            {
                                // Extract indices from [4,0,4,2] first (before modifying mergedCellName)
                                var bracketStart = mergedCellName.IndexOf('[');
                                if (bracketStart > 0)
                                {
                                    var bracketPart = mergedCellName.Substring(bracketStart + 1);
                                    bracketPart = bracketPart.TrimEnd(']');
                                    var indices = bracketPart.Split(',');
                                    if (indices.Length >= 4)
                                    {
                                        if (int.TryParse(indices[0].Trim(), out startRow) &&
                                            int.TryParse(indices[1].Trim(), out startCol) &&
                                            int.TryParse(indices[2].Trim(), out endRow) &&
                                            int.TryParse(indices[3].Trim(), out endCol))
                                        {
                                            // Use parsed indices - success!
                                            // Convert to cell names for display
                                            var startCellName = CellsHelper.CellIndexToName(startRow, startCol);
                                            var endCellName = CellsHelper.CellIndexToName(endRow, endCol);
                                            mergedCellName = $"{startCellName}:{endCellName}";
                                        }
                                    }
                                }
                                
                                // Also extract range from CellArea format: "CellArea(A5:C5)[4,0,4,2]"
                                // This is for display purposes if indices parsing failed
                                if (startRow < 0)
                                {
                                    var startIdx = mergedCellName.IndexOf('(') + 1;
                                    var endIdx = mergedCellName.IndexOf(')', startIdx);
                                    if (startIdx > 0 && endIdx > startIdx)
                                    {
                                        var rangePart = mergedCellName.Substring(startIdx, endIdx - startIdx);
                                        mergedCellName = rangePart; // Now it's "A5:C5"
                                    }
                                }
                            }
                            
                            // If we still need to parse, try splitting by ':'
                            if (startRow < 0 && mergedCellName.Contains(':'))
                            {
                                var parts = mergedCellName.Split(':');
                                if (parts.Length == 2)
                                {
                                    var startCell = parts[0].Trim();
                                    var endCell = parts[1].Trim();
                                    
                                    try
                                    {
                                        CellsHelper.CellNameToIndex(startCell, out startRow, out startCol);
                                        CellsHelper.CellNameToIndex(endCell, out endRow, out endCol);
                                    }
                                    catch
                                    {
                                        // If parsing fails, keep mergedCellName as is
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Skip this merged cell if we can't access it
                    result.AppendLine($"【合併區域 {i}】");
                    result.AppendLine($"錯誤: 無法讀取合併單元格資訊 - {ex.Message}");
                    result.AppendLine();
                    continue;
                }
                
                if (string.IsNullOrWhiteSpace(mergedCellName) || startRow < 0)
                {
                    result.AppendLine($"【合併區域 {i}】");
                    result.AppendLine($"範圍: {mergedCellName ?? "未知"}");
                    result.AppendLine($"注意: 無法解析範圍格式");
                    result.AppendLine();
                    continue;
                }
                
                // Display merged cell information
                result.AppendLine($"【合併區域 {i}】");
                result.AppendLine($"範圍: {mergedCellName}");
                
                if (startRow >= 0 && endRow >= 0 && startCol >= 0 && endCol >= 0)
                {
                    result.AppendLine($"行數: {endRow - startRow + 1}");
                    result.AppendLine($"列數: {endCol - startCol + 1}");
                    
                    try
                    {
                        var cell = worksheet.Cells[startRow, startCol];
                        result.AppendLine($"值: {cell.Value ?? "(空白)"}");
                    }
                    catch
                    {
                        result.AppendLine($"值: (無法讀取)");
                    }
                }
                result.AppendLine();
            }
            catch (Exception ex)
            {
                // Skip this merged cell if there's an error
                result.AppendLine($"【合併區域 {i}】");
                result.AppendLine($"錯誤: 無法讀取合併單元格資訊 - {ex.Message}");
                result.AppendLine();
            }
        }

        return await Task.FromResult(result.ToString());
    }
}
