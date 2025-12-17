using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel pivot tables (add, edit, delete, get, add_field, delete_field, refresh)
///     Merges: ExcelAddPivotTableTool, ExcelEditPivotTableTool, ExcelDeletePivotTableTool,
///     ExcelGetPivotTablesTool, ExcelAddPivotTableFieldTool, ExcelDeletePivotTableFieldTool, ExcelRefreshPivotTableTool
/// </summary>
public class ExcelPivotTableTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description =>
        @"Manage Excel pivot tables. Supports 7 operations: add, edit, delete, get, add_field, delete_field, refresh.

Usage examples:
- Add pivot table: excel_pivot_table(operation='add', path='book.xlsx', sourceRange='A1:D10', destCell='F1')
- Edit pivot table: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0)
- Delete pivot table: excel_pivot_table(operation='delete', path='book.xlsx', pivotTableIndex=0)
- Get pivot tables: excel_pivot_table(operation='get', path='book.xlsx')
- Add field: excel_pivot_table(operation='add_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1', area='Row')
- Delete field: excel_pivot_table(operation='delete_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1')
- Refresh: excel_pivot_table(operation='refresh', path='book.xlsx', pivotTableIndex=0) or excel_pivot_table(operation='refresh', path='book.xlsx') to refresh all";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
    /// </summary>
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
- 'delete_field': Delete field from pivot table (required params: path, pivotTableIndex, fieldName, fieldType)
- 'refresh': Refresh pivot table data (required params: path; optional: pivotTableIndex - if not provided, refreshes all pivot tables)",
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
                description =
                    "Output file path (optional, for add/edit/delete/add_field/delete_field/refresh operations, defaults to input path)"
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
                description =
                    "Pivot table index (0-based, required for edit/delete/add_field/delete_field; optional for refresh - if not provided, refreshes all pivot tables)"
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
                description =
                    "Field type: 'Row', 'Column', 'Data', 'Page' (required for add_field and delete_field operations, also accepts 'area' as alias)",
                @enum = new[] { "Row", "Column", "Data", "Page" }
            },
            area = new
            {
                type = "string",
                description =
                    "Alias for fieldType: 'Row', 'Column', 'Data', 'Page' (optional, for add_field and delete_field operations, use fieldType or area)",
                @enum = new[] { "Row", "Column", "Data", "Page" }
            },
            function = new
            {
                type = "string",
                description =
                    "Aggregation function for data field: 'Sum', 'Count', 'Average', 'Max', 'Min' (optional, for add_field)",
                @enum = new[] { "Sum", "Count", "Average", "Max", "Min" }
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

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

    /// <summary>
    ///     Adds a new pivot table to the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing sourceRange, destCell, and optional name</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with file path</returns>
    private async Task<string> AddPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange");
        var destCell = ArgumentHelper.GetString(arguments, "destCell");
        var name = ArgumentHelper.GetString(arguments, "name", "PivotTable1");

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        var pivotTables = worksheet.PivotTables;
        var pivotIndex = pivotTables.Add($"={worksheet.Name}!{sourceRange}", destCell, name);
        var pivotTable = pivotTables[pivotIndex];

        // Add default fields: first column as row field, second column as data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        pivotTable.CalculateData();

        workbook.Save(outputPath);

        return await Task.FromResult($"Pivot table added to worksheet: {outputPath}");
    }

    /// <summary>
    ///     Edits an existing pivot table (name, refresh data)
    /// </summary>
    /// <param name="arguments">JSON arguments containing pivotTableIndex and optional name, refreshData</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with changes made</returns>
    private async Task<string> EditPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");
        var name = ArgumentHelper.GetStringNullable(arguments, "name");
        var refreshData = ArgumentHelper.GetBool(arguments, "refreshData", false);

        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex];
        var changes = new List<string>();

        if (!string.IsNullOrEmpty(name))
        {
            pivotTable.Name = name;
            changes.Add($"Name: {name}");
        }

        if (refreshData)
        {
            pivotTable.CalculateData();
            changes.Add("Data refreshed");
        }

        workbook.Save(outputPath);

        var result = $"Successfully edited pivot table #{pivotTableIndex}\n";
        if (changes.Count > 0)
        {
            result += "Changes:\n";
            foreach (var change in changes) result += $"  - {change}\n";
        }
        else
        {
            result += "No changes.\n";
        }

        result += $"Output: {outputPath}";

        return await Task.FromResult(result);
    }

    /// <summary>
    ///     Deletes a pivot table from the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments containing pivotTableIndex</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with remaining pivot table count</returns>
    private async Task<string> DeletePivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");

        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex];
        var pivotTableName = pivotTable.Name ?? $"PivotTable {pivotTableIndex}";

        pivotTables.RemoveAt(pivotTableIndex);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        workbook.Save(outputPath);

        var remainingCount = pivotTables.Count;

        return await Task.FromResult(
            $"Successfully deleted pivot table #{pivotTableIndex} ({pivotTableName})\nRemaining pivot tables in worksheet: {remainingCount}\nOutput: {outputPath}");
    }

    /// <summary>
    ///     Gets information about all pivot tables in the worksheet
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with pivot table information</returns>
    private async Task<string> GetPivotTablesAsync(JsonObject? _, string path, int sheetIndex)
    {
        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;
        var result = new StringBuilder();

        result.AppendLine($"=== Pivot Table Information for Worksheet '{worksheet.Name}' ===\n");
        result.AppendLine($"Total pivot tables: {pivotTables.Count}\n");

        if (pivotTables.Count == 0)
        {
            result.AppendLine("No pivot tables found");
            return await Task.FromResult(result.ToString());
        }

        for (var i = 0; i < pivotTables.Count; i++)
        {
            var pivotTable = pivotTables[i];
            result.AppendLine($"【Pivot Table {i}】");
            result.AppendLine($"Name: {pivotTable.Name ?? "(no name)"}");
            result.AppendLine($"Data source: {pivotTable.DataSource}");
            var dataBodyRange = pivotTable.DataBodyRange;
            if (dataBodyRange.StartRow >= 0)
                result.AppendLine(
                    $"Location: Rows {dataBodyRange.StartRow}-{dataBodyRange.EndRow}, Columns {dataBodyRange.StartColumn}-{dataBodyRange.EndColumn}");
            else
                result.AppendLine("Location: Unknown");

            if (pivotTable.RowFields is { Count: > 0 } rowFields)
                result.AppendLine($"Row fields: {rowFields.Count}");

            if (pivotTable.ColumnFields is { Count: > 0 } columnFields)
                result.AppendLine($"Column fields: {columnFields.Count}");

            if (pivotTable.DataFields is { Count: > 0 } dataFields)
                result.AppendLine($"Data fields: {dataFields.Count}");

            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }

    /// <summary>
    ///     Adds a field to the pivot table
    /// </summary>
    /// <param name="arguments">JSON arguments containing pivotTableIndex, fieldName, fieldType, and optional function</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with field details</returns>
    private async Task<string> AddFieldAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        try
        {
            var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            // Support both "fieldType" and "area" parameter names for compatibility
            var fieldType = ArgumentHelper.GetStringNullable(arguments, "fieldType")
                            ?? ArgumentHelper.GetStringNullable(arguments, "area");
            if (string.IsNullOrEmpty(fieldType))
                throw new ArgumentException("fieldType (or area) parameter is required for add_field operation");
            var function = ArgumentHelper.GetString(arguments, "function", "Sum");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            PowerPointHelper.ValidateCollectionIndex(pivotTableIndex, pivotTables, "Pivot table");

            var pivotTable = pivotTables[pivotTableIndex];

            // Get data source (supports string or array formats: "=Sheet1!A1:C4", "Sheet1!A1:C4", or "A1:C4")
            string? sourceRangeStr = null;
            var dataSource = pivotTable.DataSource;

            if (dataSource is Array { Length: > 0 } dataSourceArray)
                sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
            else if (dataSource != null) sourceRangeStr = dataSource.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr))
                throw new ArgumentException("Pivot table data source is not available");

            var sourceSheet = workbook.Worksheets[sheetIndex];
            var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
            var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
            var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

            if (string.IsNullOrEmpty(rangeStr))
                throw new ArgumentException($"Invalid data source format: {sourceRangeStr}");

            Range sourceRangeObj;
            try
            {
                sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
            }
            catch (Exception rangeEx)
            {
                throw new ArgumentException(
                    $"Failed to parse pivot table data source range '{rangeStr}' from source '{sourceRangeStr}': {rangeEx.Message}");
            }

            var fieldIndex = -1;
            // Check if first row contains headers (common case: first row is header row)
            var headerRowIndex = sourceRangeObj.FirstRow;

            // Try to find field name in header row
            for (var col = sourceRangeObj.FirstColumn;
                 col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                 col++)
            {
                var headerCell = sourceSheet.Cells[headerRowIndex, col];
                var cellValue = headerCell.Value?.ToString()?.Trim();
                if (cellValue == fieldName || cellValue == fieldName.Trim())
                {
                    fieldIndex = col - sourceRangeObj.FirstColumn;
                    break;
                }
            }

            // If not found in first row, search all rows in the range (for cases where header might be elsewhere)
            if (fieldIndex < 0)
                for (var row = sourceRangeObj.FirstRow;
                     row < sourceRangeObj.FirstRow + sourceRangeObj.RowCount;
                     row++)
                {
                    for (var col = sourceRangeObj.FirstColumn;
                         col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                         col++)
                    {
                        var cell = sourceSheet.Cells[row, col];
                        var cellValue = cell.Value?.ToString()?.Trim();
                        if (cellValue == fieldName || cellValue == fieldName.Trim())
                        {
                            // Found the field, calculate its column index relative to range start
                            fieldIndex = col - sourceRangeObj.FirstColumn;
                            break;
                        }
                    }

                    if (fieldIndex >= 0) break;
                }

            if (fieldIndex < 0)
            {
                // Provide more detailed error message with available field names
                var availableFields = new List<string>();
                for (var col = sourceRangeObj.FirstColumn;
                     col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                     col++)
                {
                    var headerCell = sourceSheet.Cells[headerRowIndex, col];
                    var cellValue = headerCell.Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(cellValue))
                        availableFields.Add(cellValue);
                }

                var availableFieldsStr = availableFields.Count > 0
                    ? $" Available fields in header row: {string.Join(", ", availableFields)}"
                    : " No field names found in header row.";

                throw new ArgumentException(
                    $"Field '{fieldName}' not found in pivot table source data.{availableFieldsStr} Please check that the field name matches exactly (case-sensitive).");
            }

            try
            {
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
                            var dataField = pivotTable.DataFields[^1];
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
                    default:
                        throw new ArgumentException(
                            $"Invalid fieldType: {fieldType}. Valid values are: Row, Column, Data, Page");
                }

                // CalculateData updates the pivot table (RefreshData may cause issues)
                string? calcWarning = null;
                try
                {
                    pivotTable.CalculateData();
                }
                catch (Exception calcEx)
                {
                    calcWarning = calcEx.Message;
                }

                var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                try
                {
                    workbook.Save(outputPath);
                }
                catch (Exception saveEx)
                {
                    throw new ArgumentException(
                        $"Failed to save workbook after adding field '{fieldName}': {saveEx.Message}");
                }

                if (!string.IsNullOrEmpty(calcWarning))
                    return await Task.FromResult(
                        $"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex} (note: CalculateData warning: {calcWarning}): {outputPath}");
                return await Task.FromResult(
                    $"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex}: {outputPath}");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
                {
                    var outputPath2 = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                    try
                    {
                        workbook.Save(outputPath2);
                        return await Task.FromResult(
                            $"Field '{fieldName}' may already exist in {fieldType} area of pivot table #{pivotTableIndex}: {outputPath2}");
                    }
                    catch (Exception saveEx)
                    {
                        throw new ArgumentException(
                            $"Failed to add field '{fieldName}' to pivot table and save workbook: {ex.Message}. Save error: {saveEx.Message}");
                    }
                }

                throw new ArgumentException(
                    $"Failed to add field '{fieldName}' to pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {fieldType}");
            }
        }
        catch (Exception outerEx)
        {
            var fieldNameForError = ArgumentHelper.GetString(arguments, "fieldName", "unknown");
            throw new ArgumentException($"Failed to add field '{fieldNameForError}' to pivot table: {outerEx.Message}");
        }
    }

    /// <summary>
    ///     Removes a field from the pivot table
    /// </summary>
    /// <param name="arguments">JSON arguments containing pivotTableIndex, fieldName, and fieldType</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with field removal details</returns>
    private async Task<string> DeleteFieldAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        try
        {
            var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");
            var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
            var fieldType = ArgumentHelper.GetString(arguments, "fieldType");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            PowerPointHelper.ValidateCollectionIndex(pivotTableIndex, pivotTables, "Pivot table");

            var pivotTable = pivotTables[pivotTableIndex];

            // Get data source (supports string or array formats: "=Sheet1!A1:C4", "Sheet1!A1:C4", or "A1:C4")
            string? sourceRangeStr = null;
            var dataSource = pivotTable.DataSource;

            if (dataSource is Array { Length: > 0 } dataSourceArray)
                sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
            else if (dataSource != null) sourceRangeStr = dataSource.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr))
                throw new ArgumentException("Pivot table data source is not available");

            var sourceSheet = workbook.Worksheets[sheetIndex];
            var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
            var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
            var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

            if (string.IsNullOrEmpty(rangeStr))
                throw new ArgumentException($"Invalid data source format: {sourceRangeStr}");

            Range sourceRangeObj;
            try
            {
                sourceRangeObj = sourceSheet.Cells.CreateRange(rangeStr);
            }
            catch (Exception rangeEx)
            {
                throw new ArgumentException(
                    $"Failed to parse pivot table data source range '{rangeStr}' from source '{sourceRangeStr}': {rangeEx.Message}");
            }

            var fieldIndex = -1;
            // Check if first row contains headers (common case: first row is header row)
            var headerRowIndex = sourceRangeObj.FirstRow;

            // Try to find field name in header row
            for (var col = sourceRangeObj.FirstColumn;
                 col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                 col++)
            {
                var headerCell = sourceSheet.Cells[headerRowIndex, col];
                var cellValue = headerCell.Value?.ToString()?.Trim();
                if (cellValue == fieldName || cellValue == fieldName.Trim())
                {
                    fieldIndex = col - sourceRangeObj.FirstColumn;
                    break;
                }
            }

            // If not found in first row, search all rows in the range (for cases where header might be elsewhere)
            if (fieldIndex < 0)
                for (var row = sourceRangeObj.FirstRow;
                     row < sourceRangeObj.FirstRow + sourceRangeObj.RowCount;
                     row++)
                {
                    for (var col = sourceRangeObj.FirstColumn;
                         col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                         col++)
                    {
                        var cell = sourceSheet.Cells[row, col];
                        var cellValue = cell.Value?.ToString()?.Trim();
                        if (cellValue == fieldName || cellValue == fieldName.Trim())
                        {
                            // Found the field, calculate its column index relative to range start
                            fieldIndex = col - sourceRangeObj.FirstColumn;
                            break;
                        }
                    }

                    if (fieldIndex >= 0) break;
                }

            if (fieldIndex < 0)
            {
                // Provide more detailed error message with available field names
                var availableFields = new List<string>();
                for (var col = sourceRangeObj.FirstColumn;
                     col < sourceRangeObj.FirstColumn + sourceRangeObj.ColumnCount;
                     col++)
                {
                    var headerCell = sourceSheet.Cells[headerRowIndex, col];
                    var cellValue = headerCell.Value?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(cellValue))
                        availableFields.Add(cellValue);
                }

                var availableFieldsStr = availableFields.Count > 0
                    ? $" Available fields in header row: {string.Join(", ", availableFields)}"
                    : " No field names found in header row.";

                throw new ArgumentException(
                    $"Field '{fieldName}' not found in pivot table source data.{availableFieldsStr} Please check that the field name matches exactly (case-sensitive).");
            }

            var fieldTypeEnum = fieldType.ToLower() switch
            {
                "row" => PivotFieldType.Row,
                "column" => PivotFieldType.Column,
                "data" => PivotFieldType.Data,
                "page" => PivotFieldType.Page,
                _ => throw new ArgumentException($"Invalid fieldType: {fieldType}")
            };

            try
            {
                pivotTable.RemoveField(fieldTypeEnum, fieldIndex);

                // CalculateData updates the pivot table (RefreshData may cause issues)
                string? calcWarning = null;
                try
                {
                    pivotTable.CalculateData();
                }
                catch (Exception calcEx)
                {
                    calcWarning = calcEx.Message;
                }

                var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                try
                {
                    workbook.Save(outputPath);
                }
                catch (Exception saveEx)
                {
                    throw new ArgumentException(
                        $"Failed to save workbook after removing field '{fieldName}': {saveEx.Message}");
                }

                if (!string.IsNullOrEmpty(calcWarning))
                    return await Task.FromResult(
                        $"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex} (note: CalculateData warning: {calcWarning}): {outputPath}");
                return await Task.FromResult(
                    $"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex}: {outputPath}");
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
                {
                    var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                    try
                    {
                        workbook.Save(outputPath);
                        return await Task.FromResult(
                            $"Field '{fieldName}' may already be removed from {fieldType} area of pivot table #{pivotTableIndex}: {outputPath}");
                    }
                    catch (Exception saveEx)
                    {
                        throw new ArgumentException(
                            $"Failed to remove field '{fieldName}' from pivot table and save workbook: {ex.Message}. Save error: {saveEx.Message}");
                    }
                }

                throw new ArgumentException(
                    $"Failed to remove field '{fieldName}' from pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {fieldType}");
            }
        }
        catch (Exception outerEx)
        {
            var fieldNameForError = ArgumentHelper.GetString(arguments, "fieldName", "unknown");
            throw new ArgumentException(
                $"Failed to remove field '{fieldNameForError}' from pivot table: {outerEx.Message}");
        }
    }

    /// <summary>
    ///     Refreshes pivot table data (one or all tables)
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional pivotTableIndex (if null, refreshes all)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with refresh count</returns>
    private async Task<string> RefreshPivotTableAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pivotTableIndex = ArgumentHelper.GetIntNullable(arguments, "pivotTableIndex");

        using var workbook = new Workbook(path);

        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTables.Count == 0)
            throw new InvalidOperationException($"No pivot tables found in worksheet '{worksheet.Name}'");

        var refreshedCount = 0;

        if (pivotTableIndex.HasValue)
        {
            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

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

        return await Task.FromResult(
            $"Successfully refreshed {refreshedCount} pivot table(s)\nWorksheet: {worksheet.Name}\nOutput: {outputPath}");
    }
}