using System.Text.Json;
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
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel pivot tables. Supports 7 operations: add, edit, delete, get, add_field, delete_field, refresh.

NOTE: The 'add' operation creates a pivot table with default field settings:
- First column (index 0) is added as Row field
- Second column (index 1) is added as Data field
Use 'add_field' operation to customize field arrangement after creation.

Usage examples:
- Add pivot table: excel_pivot_table(operation='add', path='book.xlsx', sourceRange='A1:D10', destCell='F1')
- Edit pivot table: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0, name='NewName')
- Edit with style: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0, style='Medium6')
- Edit layout options: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0, showRowGrand=true, showColumnGrand=false)
- Edit with auto-fit: excel_pivot_table(operation='edit', path='book.xlsx', pivotTableIndex=0, autoFitColumns=true)
- Delete pivot table: excel_pivot_table(operation='delete', path='book.xlsx', pivotTableIndex=0)
- Get pivot tables: excel_pivot_table(operation='get', path='book.xlsx')
- Add field: excel_pivot_table(operation='add_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1', area='Row')
- Delete field: excel_pivot_table(operation='delete_field', path='book.xlsx', pivotTableIndex=0, fieldName='Column1', fieldType='Row')
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
            style = new
            {
                type = "string",
                description = @"Pivot table style (optional, for edit). Common styles:
- Light styles: 'Light1' to 'Light28'
- Medium styles: 'Medium1' to 'Medium28'
- Dark styles: 'Dark1' to 'Dark28'
- 'None' to remove style"
            },
            showRowGrand = new
            {
                type = "boolean",
                description = "Show row grand totals (optional, for edit)"
            },
            showColumnGrand = new
            {
                type = "boolean",
                description = "Show column grand totals (optional, for edit)"
            },
            autoFitColumns = new
            {
                type = "boolean",
                description = "Auto-fit column widths after editing (optional, for edit)"
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or parameters are invalid.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddPivotTableAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditPivotTableAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeletePivotTableAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetPivotTablesAsync(path, sheetIndex),
            "add_field" => await AddFieldAsync(path, outputPath, sheetIndex, arguments),
            "delete_field" => await DeleteFieldAsync(path, outputPath, sheetIndex, arguments),
            "refresh" => await RefreshPivotTableAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new pivot table to the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing sourceRange, destCell, and optional name.</param>
    /// <returns>Success message with file path.</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range.</exception>
    private Task<string> AddPivotTableAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sourceRange = ArgumentHelper.GetString(arguments, "sourceRange");
            var destCell = ArgumentHelper.GetString(arguments, "destCell");
            var name = ArgumentHelper.GetString(arguments, "name", "PivotTable1");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var pivotTables = worksheet.PivotTables;
            var pivotIndex = pivotTables.Add($"={worksheet.Name}!{sourceRange}", destCell, name);
            var pivotTable = pivotTables[pivotIndex];

            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            pivotTable.CalculateData();

            workbook.Save(outputPath);

            return $"Pivot table '{name}' added to worksheet. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing pivot table (name, style, layout, refresh data).
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">
    ///     JSON arguments containing pivotTableIndex and optional name, style, showRowGrand,
    ///     showColumnGrand, autoFitColumns, refreshData.
    /// </param>
    /// <returns>Success message with changes made.</returns>
    /// <exception cref="ArgumentException">Thrown when pivot table index is out of range or style is invalid.</exception>
    private Task<string> EditPivotTableAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");
            var name = ArgumentHelper.GetStringNullable(arguments, "name");
            var style = ArgumentHelper.GetStringNullable(arguments, "style");
            var showRowGrand = ArgumentHelper.GetBoolNullable(arguments, "showRowGrand");
            var showColumnGrand = ArgumentHelper.GetBoolNullable(arguments, "showColumnGrand");
            var autoFitColumns = ArgumentHelper.GetBool(arguments, "autoFitColumns", false);
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
                changes.Add($"name={name}");
            }

            if (!string.IsNullOrEmpty(style))
            {
                var styleType = ParsePivotTableStyle(style);
                pivotTable.PivotTableStyleType = styleType;
                changes.Add($"style={style}");
            }

            if (showRowGrand.HasValue)
            {
                pivotTable.RowGrand = showRowGrand.Value;
                changes.Add($"showRowGrand={showRowGrand.Value}");
            }

            if (showColumnGrand.HasValue)
            {
                pivotTable.ColumnGrand = showColumnGrand.Value;
                changes.Add($"showColumnGrand={showColumnGrand.Value}");
            }

            if (refreshData)
            {
                pivotTable.CalculateData();
                changes.Add("refreshed");
            }

            if (autoFitColumns)
            {
                worksheet.AutoFitColumns();
                changes.Add("autoFitColumns");
            }

            workbook.Save(outputPath);

            var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
            return $"Pivot table #{pivotTableIndex} edited ({changesStr}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Parses a style name string to PivotTableStyleType enum.
    /// </summary>
    /// <param name="style">Style name (e.g., "Light1", "Medium6", "Dark3", "None").</param>
    /// <returns>The corresponding PivotTableStyleType value.</returns>
    /// <exception cref="ArgumentException">Thrown when style name is invalid.</exception>
    private static PivotTableStyleType ParsePivotTableStyle(string style)
    {
        if (string.Equals(style, "None", StringComparison.OrdinalIgnoreCase))
            return PivotTableStyleType.None;

        if (Enum.TryParse<PivotTableStyleType>($"PivotTableStyle{style}", true, out var result))
            return result;

        if (Enum.TryParse(style, true, out result))
            return result;

        throw new ArgumentException(
            $"Invalid style: '{style}'. Valid formats: 'Light1'-'Light28', 'Medium1'-'Medium28', 'Dark1'-'Dark28', or 'None'");
    }

    /// <summary>
    ///     Deletes a pivot table from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing pivotTableIndex.</param>
    /// <returns>Success message with remaining pivot table count.</returns>
    /// <exception cref="ArgumentException">Thrown when pivot table index is out of range.</exception>
    private Task<string> DeletePivotTableAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
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
            workbook.Save(outputPath);

            return $"Pivot table #{pivotTableIndex} ({pivotTableName}) deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets information about all pivot tables in the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with pivot table information.</returns>
    /// <exception cref="ArgumentException">Thrown when sheet index is out of range.</exception>
    private Task<string> GetPivotTablesAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            if (pivotTables.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    worksheetName = worksheet.Name,
                    items = Array.Empty<object>(),
                    message = "No pivot tables found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var pivotTableList = new List<object>();
            for (var i = 0; i < pivotTables.Count; i++)
            {
                var pivotTable = pivotTables[i];

                // Format data source information
                string dataSourceInfo;
                if (pivotTable.DataSource is Array { Length: > 0 } dataSourceArray)
                {
                    var sourceParts = new List<string>();
                    foreach (var item in dataSourceArray)
                        if (item != null)
                            sourceParts.Add(item.ToString() ?? "");
                    dataSourceInfo = string.Join(", ", sourceParts);
                }
                else if (pivotTable.DataSource != null)
                {
                    dataSourceInfo = pivotTable.DataSource.ToString() ?? "Unknown";
                }
                else
                {
                    dataSourceInfo = "Unknown";
                }

                object locationInfo;
                var dataBodyRange = pivotTable.DataBodyRange;
                if (dataBodyRange.StartRow >= 0)
                {
                    var startCell = CellsHelper.CellIndexToName(dataBodyRange.StartRow, dataBodyRange.StartColumn);
                    var endCell = CellsHelper.CellIndexToName(dataBodyRange.EndRow, dataBodyRange.EndColumn);
                    locationInfo = new
                    {
                        range = $"{startCell}:{endCell}",
                        startRow = dataBodyRange.StartRow,
                        endRow = dataBodyRange.EndRow,
                        startColumn = dataBodyRange.StartColumn,
                        endColumn = dataBodyRange.EndColumn
                    };
                }
                else
                {
                    locationInfo = new
                    {
                        range = "Not calculated",
                        startRow = -1,
                        endRow = -1,
                        startColumn = -1,
                        endColumn = -1
                    };
                }

                // Format row fields information
                var rowFieldsList = new List<object>();
                if (pivotTable.RowFields is { Count: > 0 } rowFields)
                    foreach (PivotField field in rowFields)
                        rowFieldsList.Add(new
                        {
                            name = field.Name ?? $"Field {field.Position}",
                            position = field.Position
                        });

                // Format column fields information
                var columnFieldsList = new List<object>();
                if (pivotTable.ColumnFields is { Count: > 0 } columnFields)
                    foreach (PivotField field in columnFields)
                        columnFieldsList.Add(new
                        {
                            name = field.Name ?? $"Field {field.Position}",
                            position = field.Position
                        });

                // Format data fields information with aggregation functions
                var dataFieldsList = new List<object>();
                if (pivotTable.DataFields is { Count: > 0 } dataFields)
                    foreach (PivotField field in dataFields)
                        dataFieldsList.Add(new
                        {
                            name = field.Name ?? $"Field {field.Position}",
                            position = field.Position,
                            function = field.Function.ToString()
                        });

                pivotTableList.Add(new
                {
                    index = i,
                    name = pivotTable.Name ?? "(no name)",
                    dataSource = dataSourceInfo,
                    location = locationInfo,
                    rowFields = rowFieldsList,
                    columnFields = columnFieldsList,
                    dataFields = dataFieldsList
                });
            }

            var result = new
            {
                count = pivotTables.Count,
                worksheetName = worksheet.Name,
                items = pivotTableList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Adds a field to the pivot table.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing pivotTableIndex, fieldName, fieldType, and optional function.</param>
    /// <returns>Success message with field details.</returns>
    /// <exception cref="ArgumentException">Thrown when pivot table index is out of range or field not found.</exception>
    private Task<string> AddFieldAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
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

                if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
                    throw new ArgumentException(
                        $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

                var pivotTable = pivotTables[pivotTableIndex];

                string? sourceRangeStr = null;
                var dataSource = pivotTable.DataSource;

                if (dataSource is Array { Length: > 0 } dataSourceArray)
                    sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
                else if (dataSource != null) sourceRangeStr = dataSource.ToString();

                if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

                if (string.IsNullOrEmpty(sourceRangeStr))
                    throw new ArgumentException(
                        $"Pivot table data source is not available. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheet.Name}'");

                var sourceSheet = workbook.Worksheets[sheetIndex];
                var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
                var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
                var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

                if (string.IsNullOrEmpty(rangeStr))
                    throw new ArgumentException(
                        $"Invalid data source format: '{sourceRangeStr}'. Unable to parse range from data source. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheet.Name}'");

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
                    try
                    {
                        pivotTable.CalculateData();
                    }
                    catch (Exception calcEx)
                    {
                        Console.Error.WriteLine($"[WARN] CalculateData warning: {calcEx.Message}");
                    }

                    try
                    {
                        workbook.Save(outputPath);
                    }
                    catch (Exception saveEx)
                    {
                        throw new ArgumentException(
                            $"Failed to save workbook after adding field '{fieldName}': {saveEx.Message}");
                    }

                    return
                        $"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex}. Output: {outputPath}";
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
                        try
                        {
                            workbook.Save(outputPath);
                            return
                                $"Field '{fieldName}' may already exist in {fieldType} area of pivot table #{pivotTableIndex}. Output: {outputPath}";
                        }
                        catch (Exception saveEx)
                        {
                            throw new ArgumentException(
                                $"Failed to add field '{fieldName}' to pivot table and save workbook: {ex.Message}. Save error: {saveEx.Message}");
                        }

                    throw new ArgumentException(
                        $"Failed to add field '{fieldName}' to pivot table: {ex.Message}. Field index: {fieldIndex}, Field type: {fieldType}");
                }
            }
            catch (Exception outerEx)
            {
                var fieldNameForError = ArgumentHelper.GetString(arguments, "fieldName", "unknown");
                throw new ArgumentException(
                    $"Failed to add field '{fieldNameForError}' to pivot table: {outerEx.Message}");
            }
        });
    }

    /// <summary>
    ///     Removes a field from the pivot table.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing pivotTableIndex, fieldName, and fieldType.</param>
    /// <returns>Success message with field removal details.</returns>
    /// <exception cref="ArgumentException">Thrown when pivot table index is out of range or field not found.</exception>
    private Task<string> DeleteFieldAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            try
            {
                var pivotTableIndex = ArgumentHelper.GetInt(arguments, "pivotTableIndex");
                var fieldName = ArgumentHelper.GetString(arguments, "fieldName");
                var fieldType = ArgumentHelper.GetString(arguments, "fieldType");

                using var workbook = new Workbook(path);
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                var pivotTables = worksheet.PivotTables;

                if (pivotTableIndex < 0 || pivotTableIndex >= pivotTables.Count)
                    throw new ArgumentException(
                        $"Pivot table index {pivotTableIndex} is out of range (worksheet has {pivotTables.Count} pivot tables)");

                var pivotTable = pivotTables[pivotTableIndex];

                string? sourceRangeStr = null;
                var dataSource = pivotTable.DataSource;

                if (dataSource is Array { Length: > 0 } dataSourceArray)
                    sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
                else if (dataSource != null) sourceRangeStr = dataSource.ToString();

                if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

                if (string.IsNullOrEmpty(sourceRangeStr))
                    throw new ArgumentException(
                        $"Pivot table data source is not available. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheet.Name}'");

                var sourceSheet = workbook.Worksheets[sheetIndex];
                var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
                var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
                var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

                if (string.IsNullOrEmpty(rangeStr))
                    throw new ArgumentException(
                        $"Invalid data source format: '{sourceRangeStr}'. Unable to parse range from data source. Pivot table index: {pivotTableIndex}, Worksheet: '{worksheet.Name}'");

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
                    try
                    {
                        pivotTable.CalculateData();
                    }
                    catch (Exception calcEx)
                    {
                        Console.Error.WriteLine($"[WARN] CalculateData warning: {calcEx.Message}");
                    }

                    try
                    {
                        workbook.Save(outputPath);
                    }
                    catch (Exception saveEx)
                    {
                        throw new ArgumentException(
                            $"Failed to save workbook after removing field '{fieldName}': {saveEx.Message}");
                    }

                    return
                        $"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex}. Output: {outputPath}";
                }
                catch (Exception ex)
                {
                    if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
                        try
                        {
                            workbook.Save(outputPath);
                            return
                                $"Field '{fieldName}' may already be removed from {fieldType} area of pivot table #{pivotTableIndex}. Output: {outputPath}";
                        }
                        catch (Exception saveEx)
                        {
                            throw new ArgumentException(
                                $"Failed to remove field '{fieldName}' from pivot table and save workbook: {ex.Message}. Save error: {saveEx.Message}");
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
        });
    }

    /// <summary>
    ///     Refreshes pivot table data (one or all tables).
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing optional pivotTableIndex (if null, refreshes all).</param>
    /// <returns>Success message with refresh count.</returns>
    /// <exception cref="ArgumentException">Thrown when pivot table index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when no pivot tables exist.</exception>
    private Task<string> RefreshPivotTableAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
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

            return $"Refreshed {refreshedCount} pivot table(s) in worksheet '{worksheet.Name}'. Output: {outputPath}";
        });
    }
}