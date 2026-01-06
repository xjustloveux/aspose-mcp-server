using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel pivot tables (add, edit, delete, get, add_field, delete_field, refresh)
///     Merges: ExcelAddPivotTableTool, ExcelEditPivotTableTool, ExcelDeletePivotTableTool,
///     ExcelGetPivotTablesTool, ExcelAddPivotTableFieldTool, ExcelDeletePivotTableFieldTool, ExcelRefreshPivotTableTool
/// </summary>
[McpServerToolType]
public class ExcelPivotTableTool
{
    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelPivotTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelPivotTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel pivot table operation (add, edit, delete, get, add_field, delete_field, refresh).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, add_field, delete_field, refresh.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="sourceRange">Source data range (e.g., 'A1:D10', required for add).</param>
    /// <param name="destCell">Destination cell for pivot table (e.g., 'F1', required for add).</param>
    /// <param name="pivotTableIndex">
    ///     Pivot table index (0-based, required for edit/delete/add_field/delete_field; optional for
    ///     refresh).
    /// </param>
    /// <param name="name">Pivot table name (optional, for add/edit).</param>
    /// <param name="refreshData">Refresh pivot table data (optional, for edit/refresh).</param>
    /// <param name="style">Pivot table style (optional, for edit).</param>
    /// <param name="showRowGrand">Show row grand totals (optional, for edit).</param>
    /// <param name="showColumnGrand">Show column grand totals (optional, for edit).</param>
    /// <param name="autoFitColumns">Auto-fit column widths after editing (optional, for edit).</param>
    /// <param name="fieldName">Field name from source data (required for add_field/delete_field).</param>
    /// <param name="fieldType">Field type: 'Row', 'Column', 'Data', 'Page' (required for add_field and delete_field).</param>
    /// <param name="area">Alias for fieldType (optional, for add_field/delete_field).</param>
    /// <param name="function">
    ///     Aggregation function for data field: 'Sum', 'Count', 'Average', 'Max', 'Min' (optional, for
    ///     add_field).
    /// </param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_pivot_table")]
    [Description(
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
- Refresh: excel_pivot_table(operation='refresh', path='book.xlsx', pivotTableIndex=0) or excel_pivot_table(operation='refresh', path='book.xlsx') to refresh all")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a pivot table (required params: path, sourceRange, destCell)
- 'edit': Edit pivot table (required params: path, pivotTableIndex)
- 'delete': Delete a pivot table (required params: path, pivotTableIndex)
- 'get': Get all pivot tables (required params: path)
- 'add_field': Add field to pivot table (required params: path, pivotTableIndex, fieldName, area)
- 'delete_field': Delete field from pivot table (required params: path, pivotTableIndex, fieldName, fieldType)
- 'refresh': Refresh pivot table data (required params: path; optional: pivotTableIndex - if not provided, refreshes all pivot tables)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Source data range (e.g., 'A1:D10', required for add)")]
        string? sourceRange = null,
        [Description("Destination cell for pivot table (e.g., 'F1', required for add)")]
        string? destCell = null,
        [Description(
            "Pivot table index (0-based, required for edit/delete/add_field/delete_field; optional for refresh)")]
        int? pivotTableIndex = null,
        [Description("Pivot table name (optional, for add/edit)")]
        string? name = null,
        [Description("Refresh pivot table data (optional, for edit/refresh)")]
        bool refreshData = false,
        [Description(@"Pivot table style (optional, for edit). Common styles:
- Light styles: 'Light1' to 'Light28'
- Medium styles: 'Medium1' to 'Medium28'
- Dark styles: 'Dark1' to 'Dark28'
- 'None' to remove style")]
        string? style = null,
        [Description("Show row grand totals (optional, for edit)")]
        bool? showRowGrand = null,
        [Description("Show column grand totals (optional, for edit)")]
        bool? showColumnGrand = null,
        [Description("Auto-fit column widths after editing (optional, for edit)")]
        bool autoFitColumns = false,
        [Description("Field name from source data (required for add_field/delete_field)")]
        string? fieldName = null,
        [Description("Field type: 'Row', 'Column', 'Data', 'Page' (required for add_field and delete_field)")]
        string? fieldType = null,
        [Description("Alias for fieldType: 'Row', 'Column', 'Data', 'Page' (optional, for add_field/delete_field)")]
        string? area = null,
        [Description(
            "Aggregation function for data field: 'Sum', 'Count', 'Average', 'Max', 'Min' (optional, for add_field)")]
        string function = "Sum")
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddPivotTable(ctx, outputPath, sheetIndex, sourceRange, destCell, name),
            "edit" => EditPivotTable(ctx, outputPath, sheetIndex, pivotTableIndex, name, style, showRowGrand,
                showColumnGrand, autoFitColumns, refreshData),
            "delete" => DeletePivotTable(ctx, outputPath, sheetIndex, pivotTableIndex),
            "get" => GetPivotTables(ctx, sheetIndex),
            "add_field" => AddField(ctx, outputPath, sheetIndex, pivotTableIndex, fieldName, fieldType ?? area,
                function),
            "delete_field" => DeleteField(ctx, outputPath, sheetIndex, pivotTableIndex, fieldName, fieldType ?? area),
            "refresh" => RefreshPivotTable(ctx, outputPath, sheetIndex, pivotTableIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new pivot table to the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="sourceRange">The source data range.</param>
    /// <param name="destCell">The destination cell for the pivot table.</param>
    /// <param name="name">The name for the pivot table.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sourceRange or destCell is not provided.</exception>
    private static string AddPivotTable(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? sourceRange, string? destCell, string? name)
    {
        if (string.IsNullOrEmpty(sourceRange))
            throw new ArgumentException("sourceRange is required for add operation");
        if (string.IsNullOrEmpty(destCell))
            throw new ArgumentException("destCell is required for add operation");

        var pivotName = name ?? "PivotTable1";
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var pivotTables = worksheet.PivotTables;
        var pivotIndex = pivotTables.Add($"={worksheet.Name}!{sourceRange}", destCell, pivotName);
        var pivotTable = pivotTables[pivotIndex];

        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        pivotTable.CalculateData();

        ctx.Save(outputPath);
        return $"Pivot table '{pivotName}' added to worksheet. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing pivot table (name, style, layout, refresh data).
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="pivotTableIndex">The pivot table index.</param>
    /// <param name="name">The new name for the pivot table.</param>
    /// <param name="style">The style to apply to the pivot table.</param>
    /// <param name="showRowGrand">Whether to show row grand totals.</param>
    /// <param name="showColumnGrand">Whether to show column grand totals.</param>
    /// <param name="autoFitColumns">Whether to auto-fit column widths.</param>
    /// <param name="refreshData">Whether to refresh the pivot table data.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when pivotTableIndex is not provided or out of range.</exception>
    private static string EditPivotTable(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? pivotTableIndex, string? name, string? style, bool? showRowGrand, bool? showColumnGrand,
        bool autoFitColumns, bool refreshData)
    {
        if (!pivotTableIndex.HasValue)
            throw new ArgumentException("pivotTableIndex is required for edit operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex.Value];
        List<string> changes = [];

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

        ctx.Save(outputPath);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "no changes";
        return $"Pivot table #{pivotTableIndex.Value} edited ({changesStr}). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Parses a style name string to PivotTableStyleType enum.
    /// </summary>
    /// <param name="style">The style name string to parse.</param>
    /// <returns>The corresponding PivotTableStyleType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the style name is invalid.</exception>
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
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="pivotTableIndex">The pivot table index to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when pivotTableIndex is not provided or out of range.</exception>
    private static string DeletePivotTable(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? pivotTableIndex)
    {
        if (!pivotTableIndex.HasValue)
            throw new ArgumentException("pivotTableIndex is required for delete operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var pivotTables = worksheet.PivotTables;

        if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
            throw new ArgumentException(
                $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

        var pivotTable = pivotTables[pivotTableIndex.Value];
        var pivotTableName = pivotTable.Name ?? $"PivotTable {pivotTableIndex.Value}";

        pivotTables.RemoveAt(pivotTableIndex.Value);
        ctx.Save(outputPath);

        return $"Pivot table #{pivotTableIndex.Value} ({pivotTableName}) deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets information about all pivot tables in the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the pivot table information.</returns>
    private static string GetPivotTables(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
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

        List<object> pivotTableList = [];
        for (var i = 0; i < pivotTables.Count; i++)
        {
            var pivotTable = pivotTables[i];

            string dataSourceInfo;
            if (pivotTable.DataSource is Array { Length: > 0 } dataSourceArray)
            {
                List<string> sourceParts = [];
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

            List<object> rowFieldsList = [];
            if (pivotTable.RowFields is { Count: > 0 } rowFields)
                foreach (PivotField field in rowFields)
                    rowFieldsList.Add(new
                    {
                        name = field.Name ?? $"Field {field.Position}",
                        position = field.Position
                    });

            List<object> columnFieldsList = [];
            if (pivotTable.ColumnFields is { Count: > 0 } columnFields)
                foreach (PivotField field in columnFields)
                    columnFieldsList.Add(new
                    {
                        name = field.Name ?? $"Field {field.Position}",
                        position = field.Position
                    });

            List<object> dataFieldsList = [];
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
    }

    /// <summary>
    ///     Adds a field to the pivot table.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="pivotTableIndex">The pivot table index.</param>
    /// <param name="fieldName">The name of the field to add.</param>
    /// <param name="fieldType">The type of field (Row, Column, Data, Page).</param>
    /// <param name="function">The aggregation function for data fields.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or field is not found.</exception>
    private static string AddField(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? pivotTableIndex, string? fieldName, string? fieldType, string function)
    {
        if (!pivotTableIndex.HasValue)
            throw new ArgumentException("pivotTableIndex is required for add_field operation");
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for add_field operation");
        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType (or area) parameter is required for add_field operation");

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            var pivotTable = pivotTables[pivotTableIndex.Value];

            string? sourceRangeStr = null;
            var dataSource = pivotTable.DataSource;

            if (dataSource is Array { Length: > 0 } dataSourceArray)
                sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
            else if (dataSource != null) sourceRangeStr = dataSource.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr))
                throw new ArgumentException(
                    $"Pivot table data source is not available. Pivot table index: {pivotTableIndex.Value}, Worksheet: '{worksheet.Name}'");

            var sourceSheet = workbook.Worksheets[sheetIndex];
            var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
            var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
            var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

            if (string.IsNullOrEmpty(rangeStr))
                throw new ArgumentException(
                    $"Invalid data source format: '{sourceRangeStr}'. Unable to parse range from data source. Pivot table index: {pivotTableIndex.Value}, Worksheet: '{worksheet.Name}'");

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
            var headerRowIndex = sourceRangeObj.FirstRow;

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
                            fieldIndex = col - sourceRangeObj.FirstColumn;
                            break;
                        }
                    }

                    if (fieldIndex >= 0) break;
                }

            if (fieldIndex < 0)
            {
                List<string> availableFields = [];
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
                    ctx.Save(outputPath);
                }
                catch (Exception saveEx)
                {
                    throw new ArgumentException(
                        $"Failed to save workbook after adding field '{fieldName}': {saveEx.Message}");
                }

                return
                    $"Field '{fieldName}' added as {fieldType} field to pivot table #{pivotTableIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate"))
                    try
                    {
                        ctx.Save(outputPath);
                        return
                            $"Field '{fieldName}' may already exist in {fieldType} area of pivot table #{pivotTableIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
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
            throw new ArgumentException(
                $"Failed to add field '{fieldName}' to pivot table: {outerEx.Message}");
        }
    }

    /// <summary>
    ///     Removes a field from the pivot table.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="pivotTableIndex">The pivot table index.</param>
    /// <param name="fieldName">The name of the field to remove.</param>
    /// <param name="fieldType">The type of field (Row, Column, Data, Page).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or field is not found.</exception>
    private static string DeleteField(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? pivotTableIndex, string? fieldName, string? fieldType)
    {
        if (!pivotTableIndex.HasValue)
            throw new ArgumentException("pivotTableIndex is required for delete_field operation");
        if (string.IsNullOrEmpty(fieldName))
            throw new ArgumentException("fieldName is required for delete_field operation");
        if (string.IsNullOrEmpty(fieldType))
            throw new ArgumentException("fieldType is required for delete_field operation");

        try
        {
            var workbook = ctx.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var pivotTables = worksheet.PivotTables;

            if (pivotTableIndex.Value < 0 || pivotTableIndex.Value >= pivotTables.Count)
                throw new ArgumentException(
                    $"Pivot table index {pivotTableIndex.Value} is out of range (worksheet has {pivotTables.Count} pivot tables)");

            var pivotTable = pivotTables[pivotTableIndex.Value];

            string? sourceRangeStr = null;
            var dataSource = pivotTable.DataSource;

            if (dataSource is Array { Length: > 0 } dataSourceArray)
                sourceRangeStr = dataSourceArray.GetValue(0)?.ToString();
            else if (dataSource != null) sourceRangeStr = dataSource.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr)) sourceRangeStr = pivotTable.DataSource?.ToString();

            if (string.IsNullOrEmpty(sourceRangeStr))
                throw new ArgumentException(
                    $"Pivot table data source is not available. Pivot table index: {pivotTableIndex.Value}, Worksheet: '{worksheet.Name}'");

            var sourceSheet = workbook.Worksheets[sheetIndex];
            var cleanSourceRange = sourceRangeStr.Replace("=", "").Trim();
            var sourceParts = cleanSourceRange.Split(['!'], StringSplitOptions.RemoveEmptyEntries);
            var rangeStr = sourceParts.Length > 1 ? sourceParts[1].Trim() : sourceParts[0].Trim();

            if (string.IsNullOrEmpty(rangeStr))
                throw new ArgumentException(
                    $"Invalid data source format: '{sourceRangeStr}'. Unable to parse range from data source. Pivot table index: {pivotTableIndex.Value}, Worksheet: '{worksheet.Name}'");

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
            var headerRowIndex = sourceRangeObj.FirstRow;

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
                            fieldIndex = col - sourceRangeObj.FirstColumn;
                            break;
                        }
                    }

                    if (fieldIndex >= 0) break;
                }

            if (fieldIndex < 0)
            {
                List<string> availableFields = [];
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
                    ctx.Save(outputPath);
                }
                catch (Exception saveEx)
                {
                    throw new ArgumentException(
                        $"Failed to save workbook after removing field '{fieldName}': {saveEx.Message}");
                }

                return
                    $"Field '{fieldName}' removed from {fieldType} area of pivot table #{pivotTableIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("not found") || ex.Message.Contains("does not exist"))
                    try
                    {
                        ctx.Save(outputPath);
                        return
                            $"Field '{fieldName}' may already be removed from {fieldType} area of pivot table #{pivotTableIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
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
            throw new ArgumentException(
                $"Failed to remove field '{fieldName}' from pivot table: {outerEx.Message}");
        }
    }

    /// <summary>
    ///     Refreshes pivot table data (one or all tables).
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="pivotTableIndex">The pivot table index, or null to refresh all pivot tables.</param>
    /// <returns>A message indicating the number of pivot tables refreshed.</returns>
    /// <exception cref="InvalidOperationException">Thrown when no pivot tables are found.</exception>
    /// <exception cref="ArgumentException">Thrown when pivotTableIndex is out of range.</exception>
    private static string RefreshPivotTable(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int? pivotTableIndex)
    {
        var workbook = ctx.Document;
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

        ctx.Save(outputPath);

        return
            $"Refreshed {refreshedCount} pivot table(s) in worksheet '{worksheet.Name}'. {ctx.GetOutputMessage(outputPath)}";
    }
}