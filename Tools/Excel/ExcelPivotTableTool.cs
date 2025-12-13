using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel pivot tables (add, edit, delete, get, add_field, delete_field, refresh)
/// Merges: ExcelAddPivotTableTool, ExcelEditPivotTableTool, ExcelDeletePivotTableTool, 
/// ExcelGetPivotTablesTool, ExcelAddPivotTableFieldTool, ExcelDeletePivotTableFieldTool, ExcelRefreshPivotTableTool
/// </summary>
public class ExcelPivotTableTool : IAsposeTool
{
    public string Description => @"Manage Excel pivot tables. Supports 7 operations: add, edit, delete, get, add_field, delete_field, refresh.

Usage examples:
- Add pivot table: excel_pivot_table(operation='add', path='book.xlsx', sourceRange='A1:D10', destCell='F1')
- Edit pivot table: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0)
- Delete pivot table: excel_pivot_table(operation='delete', path='book.xlsx', pivotTableIndex=0)
- Get pivot tables: excel_pivot_table(operation='get', path='book.xlsx')
- Add field: excel_pivot_table(operation='add_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1', area='Row')
- Delete field: excel_pivot_table(operation='delete_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1')
- Refresh: excel_pivot_table(operation='refresh', path='book.xlsx', pivotTableIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a pivot table (required params: path, sourceRange, destCell)
- 'edit': Edit pivot table (required params: path, pivotTableIndex)
- 'delete': Delete a pivot table (required params: path, pivotTableIndex)
- 'get': Get all pivot tables (required params: path)
- 'add_field': Add field to pivot table (required params: path, pivotTableIndex, fieldName, area)
- 'delete_field': Delete field from pivot table (required params: path, pivotTableIndex, fieldName)
- 'refresh': Refresh pivot table data (required params: path, pivotTableIndex)",
                @enum = new[] { "add", "edit", "delete", "get", "add_field", "delete_field", "refresh" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for edit/refresh operations, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            sourceRange = new
            {
                type = "string",
                description = "Source data range (e.g., 'A1:D10', required for add)"
            },
            destCell = new
            {
                type = "string",
                description = "Destination cell for pivot table (e.g., 'F1', required for add)"
            },
            pivotTableIndex = new
            {
                type = "number",
                description = "Pivot table index (0-based, required for edit/delete/add_field/delete_field/refresh)"
            },
            name = new
            {
                type = "string",
                description = "Pivot table name (optional, for add/edit)"
            },
            refreshData = new
            {
                type = "boolean",
                description = "Refresh pivot table data (optional, for edit/refresh)"
            },
            fieldName = new
            {
                type = "string",
                description = "Field name from source data (required for add_field/delete_field)"
            },
            fieldType = new
            {
                type = "string",
                description = "Field type: 'Row', 'Column', 'Data', 'Page' (required for add_field/delete_field)",
                @enum = new[] { "Row", "Column", "Data", "Page" }
            },
            function = new
            {
                type = "string",
                description = "Aggregation function for data field: 'Sum', 'Count', 'Average', 'Max', 'Min' (optional, for add_field)",
                @enum = new[] { "Sum", "Count", "Average", "Max", "Min" }
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        return operation.ToLower() switch
        {
            "add" => await AddPivotTableAsync(arguments, path, sheetIndex),
            "edit" => await EditPivotTableAsync(arguments, path, sheetIndex),
            "delete" => await DeletePivotTableAsync(arguments, path, sheetIndex),
            "get" => await GetPivotTablesAsync(arguments, path, sheetIndex),
            "add_field" => await AddFieldAsync(arguments, path, sheetIndex),
            "delete_field" => await DeleteFieldAsync(arguments, path, sheetIndex),
            "refresh" => await RefreshPivotTableAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var sourceRange = arguments?["sourceRange"]?.GetValue<string>() ?? throw new ArgumentException("sourceRange is required for add operation");
        var destCell = arguments?["destCell"]?.GetValue<string>() ?? throw new ArgumentException("destCell is required for add operation");
        var name = arguments?["name"]?.GetValue<string>() ?? "PivotTable1";

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        var pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add($"={worksheet.Name}!{sourceRange}", destCell, name);
        var pivotTable = pivotTables[pivotIndex];

        // Add first field as row field and second field as data field by default
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
        
        pivotTable.CalculateData();

        workbook.Save(path);

        return await Task.FromResult($"Pivot table added to worksheet: {path}");
    }

    private async Task<string> EditPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required for edit operation");
        var name = arguments?["name"]?.GetValue<string>();
        var refreshData = arguments?["refreshData"]?.GetValue<bool>() ?? false;

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;
        
        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
        {
            throw new ArgumentException($"樞紐表索引 {pivotTableIndex} 超出範圍 (工作表共有 {pivotTables.Count} 個樞紐表)");
        }

        var pivotTable = pivotTables[pivotTableIndex];
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(name))
        {
            pivotTable.Name = name;
            changes.Add($"名稱: {name}");
        }

        if (refreshData)
        {
            pivotTable.CalculateData();
            changes.Add("數據已刷新");
        }

        workbook.Save(outputPath);

        var result = $"成功編輯樞紐表 #{pivotTableIndex}\n";
        if (changes.Count > 0)
        {
            result += "變更:\n";
            foreach (var change in changes)
            {
                result += $"  - {change}\n";
            }
        }
        else
        {
            result += "無變更。\n";
        }
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private async Task<string> DeletePivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required for delete operation");

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;
        
        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
        {
            throw new ArgumentException($"樞紐表索引 {pivotTableIndex} 超出範圍 (工作表共有 {pivotTables.Count} 個樞紐表)");
        }

        var pivotTable = pivotTables[pivotTableIndex];
        var pivotTableName = pivotTable.Name ?? $"樞紐表 {pivotTableIndex}";
        
        pivotTables.RemoveAt(pivotTableIndex);
        workbook.Save(path);
        
        var remainingCount = pivotTables.Count;
        
        return await Task.FromResult($"成功刪除樞紐表 #{pivotTableIndex} ({pivotTableName})\n工作表剩餘樞紐表數: {remainingCount}\n輸出: {path}");
    }

    private async Task<string> GetPivotTablesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的樞紐表資訊 ===\n");
        result.AppendLine($"總樞紐表數: {pivotTables.Count}\n");

        if (pivotTables.Count == 0)
        {
            result.AppendLine("未找到樞紐表");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < pivotTables.Count; i++)
        {
            var pivotTable = pivotTables[i];
            result.AppendLine($"【樞紐表 {i}】");
            result.AppendLine($"名稱: {pivotTable.Name ?? "(無名稱)"}");
            result.AppendLine($"數據源: {pivotTable.DataSource}");
            var dataBodyRange = pivotTable.DataBodyRange;
            if (dataBodyRange.StartRow >= 0)
            {
                result.AppendLine($"位置: 行 {dataBodyRange.StartRow}-{dataBodyRange.EndRow}, 列 {dataBodyRange.StartColumn}-{dataBodyRange.EndColumn}");
            }
            else
            {
                result.AppendLine($"位置: 未知");
            }
            
            if (pivotTable.RowFields != null && pivotTable.RowFields.Count > 0)
            {
                result.AppendLine($"行欄位數: {pivotTable.RowFields.Count}");
            }
            
            if (pivotTable.ColumnFields != null && pivotTable.ColumnFields.Count > 0)
            {
                result.AppendLine($"列欄位數: {pivotTable.ColumnFields.Count}");
            }
            
            if (pivotTable.DataFields != null && pivotTable.DataFields.Count > 0)
            {
                result.AppendLine($"數據欄位數: {pivotTable.DataFields.Count}");
            }
            
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    private async Task<string> AddFieldAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required for add_field operation");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required for add_field operation");
        var fieldType = arguments?["fieldType"]?.GetValue<string>() ?? throw new ArgumentException("fieldType is required for add_field operation");
        var function = arguments?["function"]?.GetValue<string>() ?? "Sum";

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        PowerPointHelper.ValidateCollectionIndex(pivotTableIndex, pivotTables, "資料透視表");

        var pivotTable = pivotTables[pivotTableIndex];
        
        var sourceRange = pivotTable.DataSource?.ToString();
        if (string.IsNullOrEmpty(sourceRange))
        {
            throw new ArgumentException("Pivot table data source is not available");
        }
        
        var sourceSheet = workbook.Worksheets[sheetIndex];
        var sourceRangeStr = sourceRange.Replace("=", "");
        var sourceParts = sourceRangeStr.Split(new[] { '!' }, StringSplitOptions.RemoveEmptyEntries);
        var rangeStr = sourceParts.Length > 1 ? sourceParts[1] : sourceParts[0];
        var sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
        
        int fieldIndex = -1;
        for (int col = sourceRangeObj.FirstColumn; col <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount; col++)
        {
            var headerCell = sourceSheet.Cells[sourceRangeObj.FirstRow, col];
            if (headerCell.Value?.ToString() == fieldName)
            {
                fieldIndex = col - sourceRangeObj.FirstColumn;
                break;
            }
        }

        if (fieldIndex < 0)
        {
            throw new ArgumentException($"Field '{fieldName}' not found in pivot table source data");
        }

        switch (fieldType.ToLower())
        {
            case "row":
                pivotTable.AddFieldToArea(PivotFieldType.Row, fieldIndex);
                break;
            case "column":
                pivotTable.AddFieldToArea(PivotFieldType.Column, fieldIndex);
                break;
            case "data":
                pivotTable.AddFieldToArea(PivotFieldType.Data, fieldIndex);
                if (pivotTable.DataFields.Count > 0)
                {
                    var dataField = pivotTable.DataFields[pivotTable.DataFields.Count - 1];
                    var functionType = function switch
                    {
                        "Count" => ConsolidationFunction.Count,
                        "Average" => ConsolidationFunction.Average,
                        "Max" => ConsolidationFunction.Max,
                        "Min" => ConsolidationFunction.Min,
                        _ => ConsolidationFunction.Sum
                    };
                    dataField.Function = functionType;
                }
                break;
            case "page":
                pivotTable.AddFieldToArea(PivotFieldType.Page, fieldIndex);
                break;
        }

        pivotTable.RefreshData();
        pivotTable.CalculateData();

        workbook.Save(path);
        return await Task.FromResult($"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex}: {path}");
    }

    private async Task<string> DeleteFieldAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int>() ?? throw new ArgumentException("pivotTableIndex is required for delete_field operation");
        var fieldName = arguments?["fieldName"]?.GetValue<string>() ?? throw new ArgumentException("fieldName is required for delete_field operation");
        var fieldType = arguments?["fieldType"]?.GetValue<string>() ?? throw new ArgumentException("fieldType is required for delete_field operation");

        using var workbook = new Workbook(path);
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        PowerPointHelper.ValidateCollectionIndex(pivotTableIndex, pivotTables, "資料透視表");

        var pivotTable = pivotTables[pivotTableIndex];
        
        var sourceRange = pivotTable.DataSource?.ToString();
        if (string.IsNullOrEmpty(sourceRange))
        {
            throw new ArgumentException("Pivot table data source is not available");
        }
        
        var sourceSheet = workbook.Worksheets[sheetIndex];
        var sourceRangeStr = sourceRange.Replace("=", "");
        var sourceParts = sourceRangeStr.Split(new[] { '!' }, StringSplitOptions.RemoveEmptyEntries);
        var rangeStr = sourceParts.Length > 1 ? sourceParts[1] : sourceParts[0];
        var sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
        
        int fieldIndex = -1;
        for (int col = sourceRangeObj.FirstColumn; col <= sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount; col++)
        {
            var headerCell = sourceSheet.Cells[sourceRangeObj.FirstRow, col];
            if (headerCell.Value?.ToString() == fieldName)
            {
                fieldIndex = col - sourceRangeObj.FirstColumn;
                break;
            }
        }

        if (fieldIndex < 0)
        {
            throw new ArgumentException($"Field '{fieldName}' not found in pivot table source data");
        }

        var fieldTypeEnum = fieldType.ToLower() switch
        {
            "row" => PivotFieldType.Row,
            "column" => PivotFieldType.Column,
            "data" => PivotFieldType.Data,
            "page" => PivotFieldType.Page,
            _ => throw new ArgumentException($"Invalid fieldType: {fieldType}")
        };

        pivotTable.RemoveField(fieldTypeEnum, fieldIndex);

        pivotTable.RefreshData();
        pivotTable.CalculateData();

        workbook.Save(path);
        return await Task.FromResult($"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex}: {path}");
    }

    private async Task<string> RefreshPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var pivotTableIndex = arguments?["pivotTableIndex"]?.GetValue<int?>();

        using var workbook = new Workbook(path);
        
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;
        
        if (pivotTables.Count == 0)
        {
            throw new InvalidOperationException($"工作表 '{worksheet.Name}' 中未找到樞紐表");
        }

        int refreshedCount = 0;

        if (pivotTableIndex.HasValue)
        {
            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
            {
                throw new ArgumentException($"樞紐表索引 {pivotTableIndex.Value} 超出範圍 (工作表共有 {pivotTables.Count} 個樞紐表)");
            }
            
            pivotTables[pivotTableIndex.Value].CalculateData();
            refreshedCount = 1;
        }
        else
        {
            foreach (var pivotTable in pivotTables)
            {
                pivotTable.CalculateData();
                refreshedCount++;
            }
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"成功刷新 {refreshedCount} 個樞紐表\n工作表: {worksheet.Name}\n輸出: {outputPath}");
    }
}

