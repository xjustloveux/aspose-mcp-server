using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

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
    ///     Handler registry for pivot table operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

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
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.PivotTable");
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
    public string Execute( // NOSONAR S107 - MCP protocol requires multiple parameters
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

        var parameters = BuildParameters(operation, sheetIndex, sourceRange, destCell, pivotTableIndex, name,
            style, showRowGrand, showColumnGrand, autoFitColumns, refreshData, fieldName, fieldType ?? area, function);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters( // NOSONAR S107
        string operation,
        int sheetIndex,
        string? sourceRange,
        string? destCell,
        int? pivotTableIndex,
        string? name,
        string? style,
        bool? showRowGrand,
        bool? showColumnGrand,
        bool autoFitColumns,
        bool refreshData,
        string? fieldName,
        string? fieldType,
        string function)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, sourceRange, destCell, name),
            "edit" => BuildEditParameters(parameters, pivotTableIndex, name, style, showRowGrand, showColumnGrand,
                autoFitColumns, refreshData),
            "delete" or "refresh" => BuildIndexParameters(parameters, pivotTableIndex),
            "get" => parameters,
            "add_field" => BuildAddFieldParameters(parameters, pivotTableIndex, fieldName, fieldType, function),
            "delete_field" => BuildDeleteFieldParameters(parameters, pivotTableIndex, fieldName, fieldType),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add pivot table operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="sourceRange">The source data range (e.g., 'A1:D10').</param>
    /// <param name="destCell">The destination cell for the pivot table (e.g., 'F1').</param>
    /// <param name="name">The pivot table name.</param>
    /// <returns>OperationParameters configured for the add operation.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? sourceRange,
        string? destCell, string? name)
    {
        if (sourceRange != null) parameters.Set("sourceRange", sourceRange);
        if (destCell != null) parameters.Set("destCell", destCell);
        if (name != null) parameters.Set("name", name);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit pivot table operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pivotTableIndex">The pivot table index (0-based).</param>
    /// <param name="name">The new pivot table name.</param>
    /// <param name="style">The pivot table style (e.g., 'Medium6').</param>
    /// <param name="showRowGrand">Whether to show row grand totals.</param>
    /// <param name="showColumnGrand">Whether to show column grand totals.</param>
    /// <param name="autoFitColumns">Whether to auto-fit column widths after editing.</param>
    /// <param name="refreshData">Whether to refresh pivot table data.</param>
    /// <returns>OperationParameters configured for the edit operation.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int? pivotTableIndex,
        string? name, string? style, bool? showRowGrand, bool? showColumnGrand, bool autoFitColumns, bool refreshData)
    {
        if (pivotTableIndex.HasValue) parameters.Set("pivotTableIndex", pivotTableIndex.Value);
        if (name != null) parameters.Set("name", name);
        if (style != null) parameters.Set("style", style);
        if (showRowGrand.HasValue) parameters.Set("showRowGrand", showRowGrand.Value);
        if (showColumnGrand.HasValue) parameters.Set("showColumnGrand", showColumnGrand.Value);
        parameters.Set("autoFitColumns", autoFitColumns);
        parameters.Set("refreshData", refreshData);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for index-based operations (delete, refresh).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pivotTableIndex">The pivot table index (0-based).</param>
    /// <returns>OperationParameters configured for index-based operations.</returns>
    private static OperationParameters BuildIndexParameters(OperationParameters parameters, int? pivotTableIndex)
    {
        if (pivotTableIndex.HasValue) parameters.Set("pivotTableIndex", pivotTableIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pivotTableIndex">The pivot table index (0-based).</param>
    /// <param name="fieldName">The field name from source data.</param>
    /// <param name="fieldType">The field type: 'Row', 'Column', 'Data', 'Page'.</param>
    /// <param name="function">The aggregation function for data field (e.g., 'Sum', 'Count', 'Average').</param>
    /// <returns>OperationParameters configured for the add field operation.</returns>
    private static OperationParameters BuildAddFieldParameters(OperationParameters parameters, int? pivotTableIndex,
        string? fieldName, string? fieldType, string function)
    {
        if (pivotTableIndex.HasValue) parameters.Set("pivotTableIndex", pivotTableIndex.Value);
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (fieldType != null) parameters.Set("fieldType", fieldType);
        parameters.Set("function", function);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete field operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="pivotTableIndex">The pivot table index (0-based).</param>
    /// <param name="fieldName">The field name to delete.</param>
    /// <param name="fieldType">The field type: 'Row', 'Column', 'Data', 'Page'.</param>
    /// <returns>OperationParameters configured for the delete field operation.</returns>
    private static OperationParameters BuildDeleteFieldParameters(OperationParameters parameters, int? pivotTableIndex,
        string? fieldName, string? fieldType)
    {
        if (pivotTableIndex.HasValue) parameters.Set("pivotTableIndex", pivotTableIndex.Value);
        if (fieldName != null) parameters.Set("fieldName", fieldName);
        if (fieldType != null) parameters.Set("fieldType", fieldType);
        return parameters;
    }
}
