using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel ranges (write, edit, get, clear, copy, move, copy_format)
///     Merges: ExcelWriteRangeTool, ExcelEditRangeTool, ExcelGetRangeTool, ExcelClearRangeTool,
///     ExcelCopyRangeTool, ExcelMoveRangeTool, ExcelCopyFormatTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Range")]
[McpServerToolType]
public class ExcelRangeTool
{
    /// <summary>
    ///     Handler registry for range operations.
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
    ///     Initializes a new instance of the <see cref="ExcelRangeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelRangeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Range");
    }

    /// <summary>
    ///     Executes an Excel range operation (write, edit, get, clear, copy, move, copy_format).
    /// </summary>
    /// <param name="operation">The operation to perform: write, edit, get, clear, copy, move, copy_format.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0).</param>
    /// <param name="sourceSheetIndex">Source sheet index (0-based, optional, for copy/move, default: same as sheetIndex).</param>
    /// <param name="destSheetIndex">Destination sheet index (0-based, optional, for copy/move, default: same as source).</param>
    /// <param name="startCell">Starting cell (e.g., 'A1', required for write).</param>
    /// <param name="range">
    ///     Source cell range (e.g., 'A1:C5', required for edit/get/clear operations, optional for
    ///     copy_format).
    /// </param>
    /// <param name="sourceRange">
    ///     Source range (e.g., 'A1:C5', required for copy/move, optional for copy_format as alternative
    ///     to range).
    /// </param>
    /// <param name="destCell">Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format).</param>
    /// <param name="destRange">Destination range (e.g., 'E1:G5', required for copy_format, or use destCell).</param>
    /// <param name="data">Data to write as JSON array.</param>
    /// <param name="clearRange">Clear range before writing (optional, for edit, default: false).</param>
    /// <param name="includeFormulas">Include formulas instead of values (optional, for get, default: false).</param>
    /// <param name="calculateFormulas">Recalculate all formulas before getting values (optional, for get, default: false).</param>
    /// <param name="includeFormat">Include format information (optional, for get, default: false).</param>
    /// <param name="clearContent">Clear cell content (optional, for clear, default: true).</param>
    /// <param name="clearFormat">Clear cell format (optional, for clear, default: false).</param>
    /// <param name="copyOptions">Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All').</param>
    /// <param name="copyValue">Copy cell values as well (optional, for copy_format, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_range",
        Title = "Excel Range Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel ranges. Supports 7 operations: write, edit, get, clear, copy, move, copy_format.

Usage examples:
- Write range: excel_range(operation='write', path='book.xlsx', startCell='A1', data=[['A','B'],['C','D']])
- Edit range: excel_range(operation='edit', path='book.xlsx', range='A1:B2', data=[['X','Y']])
- Get range: excel_range(operation='get', path='book.xlsx', range='A1:B2')
- Clear range: excel_range(operation='clear', path='book.xlsx', range='A1:B2')
- Copy range: excel_range(operation='copy', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Move range: excel_range(operation='move', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Copy format: excel_range(operation='copy_format', path='book.xlsx', sourceRange='A1:B2', destCell='C1')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'write': Write data to range (required params: path, startCell, data)
- 'edit': Edit range data (required params: path, range, data)
- 'get': Get range data (required params: path, range)
- 'clear': Clear range (required params: path, range)
- 'copy': Copy range (required params: path, sourceRange, destCell)
- 'move': Move range (required params: path, sourceRange, destCell)
- 'copy_format': Copy format only (required params: path, range or sourceRange, destRange or destCell)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Source sheet index (0-based, optional, for copy/move, default: same as sheetIndex)")]
        int? sourceSheetIndex = null,
        [Description("Destination sheet index (0-based, optional, for copy/move, default: same as source)")]
        int? destSheetIndex = null,
        [Description("Starting cell (e.g., 'A1', required for write)")]
        string? startCell = null,
        [Description(
            "Source cell range (e.g., 'A1:C5', required for edit/get/clear operations, optional for copy_format)")]
        string? range = null,
        [Description(
            "Source range (e.g., 'A1:C5', required for copy/move, optional for copy_format as alternative to range)")]
        string? sourceRange = null,
        [Description("Destination cell (top-left cell, e.g., 'E1', required for copy/move, optional for copy_format)")]
        string? destCell = null,
        [Description("Destination range (e.g., 'E1:G5', required for copy_format, or use destCell)")]
        string? destRange = null,
        [Description(@"Data to write. Supports two formats:
1) 2D array: [['row1_col1', 'row1_col2'], ['row2_col1', 'row2_col2']]
2) Object array: [{'cell': 'A1', 'value': '10'}, {'cell': 'B1', 'value': '20'}]")]
        string? data = null,
        [Description("Clear range before writing (optional, for edit, default: false)")]
        bool clearRange = false,
        [Description("Include formulas instead of values (optional, for get, default: false)")]
        bool includeFormulas = false,
        [Description("Recalculate all formulas before getting values (optional, for get, default: false)")]
        bool calculateFormulas = false,
        [Description("Include format information (optional, for get, default: false)")]
        bool includeFormat = false,
        [Description("Clear cell content (optional, for clear, default: true)")]
        bool clearContent = true,
        [Description("Clear cell format (optional, for clear, default: false)")]
        bool clearFormat = false,
        [Description("Copy options: 'All', 'Values', 'Formats', 'Formulas' (optional, for copy, default: 'All')")]
        string copyOptions = "All",
        [Description("Copy cell values as well (optional, for copy_format, default: false)")]
        bool copyValue = false)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, sourceSheetIndex, destSheetIndex, startCell,
            range, sourceRange, destCell, destRange, data, clearRange, includeFormulas, calculateFormulas,
            includeFormat, clearContent, clearFormat, copyOptions, copyValue);

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
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        int? sourceSheetIndex,
        int? destSheetIndex,
        string? startCell,
        string? range,
        string? sourceRange,
        string? destCell,
        string? destRange,
        string? data,
        bool clearRange,
        bool includeFormulas,
        bool calculateFormulas,
        bool includeFormat,
        bool clearContent,
        bool clearFormat,
        string copyOptions,
        bool copyValue)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "write" => BuildWriteParameters(parameters, startCell, data),
            "edit" => BuildEditParameters(parameters, range, data, clearRange),
            "get" => BuildGetParameters(parameters, range, includeFormulas, calculateFormulas, includeFormat),
            "clear" => BuildClearParameters(parameters, range, clearContent, clearFormat),
            "copy" => BuildCopyParameters(parameters, sourceSheetIndex, destSheetIndex, sourceRange, destCell,
                copyOptions),
            "move" => BuildMoveParameters(parameters, sourceSheetIndex, destSheetIndex, sourceRange, destCell),
            "copy_format" => BuildCopyFormatParameters(parameters, range, sourceRange, destRange, destCell, copyValue),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the write range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="startCell">The starting cell address (e.g., 'A1').</param>
    /// <param name="data">The data to write as JSON array.</param>
    /// <returns>OperationParameters configured for the write operation.</returns>
    private static OperationParameters BuildWriteParameters(OperationParameters parameters, string? startCell,
        string? data)
    {
        if (startCell != null) parameters.Set("startCell", startCell);
        if (data != null) parameters.Set("data", data);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="range">The cell range (e.g., 'A1:C5').</param>
    /// <param name="data">The data to write as JSON array.</param>
    /// <param name="clearRange">Whether to clear range before writing.</param>
    /// <returns>OperationParameters configured for the edit operation.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, string? range, string? data,
        bool clearRange)
    {
        if (range != null) parameters.Set("range", range);
        if (data != null) parameters.Set("data", data);
        parameters.Set("clearRange", clearRange);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="range">The cell range (e.g., 'A1:C5').</param>
    /// <param name="includeFormulas">Whether to include formulas instead of values.</param>
    /// <param name="calculateFormulas">Whether to recalculate formulas before getting values.</param>
    /// <param name="includeFormat">Whether to include format information.</param>
    /// <returns>OperationParameters configured for the get operation.</returns>
    private static OperationParameters BuildGetParameters(OperationParameters parameters, string? range,
        bool includeFormulas, bool calculateFormulas, bool includeFormat)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("includeFormulas", includeFormulas);
        parameters.Set("calculateFormulas", calculateFormulas);
        parameters.Set("includeFormat", includeFormat);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the clear range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="range">The cell range (e.g., 'A1:C5').</param>
    /// <param name="clearContent">Whether to clear cell content.</param>
    /// <param name="clearFormat">Whether to clear cell format.</param>
    /// <returns>OperationParameters configured for the clear operation.</returns>
    private static OperationParameters BuildClearParameters(OperationParameters parameters, string? range,
        bool clearContent, bool clearFormat)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("clearContent", clearContent);
        parameters.Set("clearFormat", clearFormat);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="sourceSheetIndex">The source sheet index (0-based).</param>
    /// <param name="destSheetIndex">The destination sheet index (0-based).</param>
    /// <param name="sourceRange">The source range (e.g., 'A1:C5').</param>
    /// <param name="destCell">The destination cell (top-left cell, e.g., 'E1').</param>
    /// <param name="copyOptions">The copy options: 'All', 'Values', 'Formats', 'Formulas'.</param>
    /// <returns>OperationParameters configured for the copy operation.</returns>
    private static OperationParameters BuildCopyParameters(OperationParameters parameters, int? sourceSheetIndex,
        int? destSheetIndex, string? sourceRange, string? destCell, string copyOptions)
    {
        if (sourceSheetIndex != null) parameters.Set("sourceSheetIndex", sourceSheetIndex);
        if (destSheetIndex != null) parameters.Set("destSheetIndex", destSheetIndex);
        if (sourceRange != null) parameters.Set("sourceRange", sourceRange);
        if (destCell != null) parameters.Set("destCell", destCell);
        parameters.Set("copyOptions", copyOptions);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the move range operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="sourceSheetIndex">The source sheet index (0-based).</param>
    /// <param name="destSheetIndex">The destination sheet index (0-based).</param>
    /// <param name="sourceRange">The source range (e.g., 'A1:C5').</param>
    /// <param name="destCell">The destination cell (top-left cell, e.g., 'E1').</param>
    /// <returns>OperationParameters configured for the move operation.</returns>
    private static OperationParameters BuildMoveParameters(OperationParameters parameters, int? sourceSheetIndex,
        int? destSheetIndex, string? sourceRange, string? destCell)
    {
        if (sourceSheetIndex != null) parameters.Set("sourceSheetIndex", sourceSheetIndex);
        if (destSheetIndex != null) parameters.Set("destSheetIndex", destSheetIndex);
        if (sourceRange != null) parameters.Set("sourceRange", sourceRange);
        if (destCell != null) parameters.Set("destCell", destCell);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy format operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="range">The source range (alternative to sourceRange).</param>
    /// <param name="sourceRange">The source range (e.g., 'A1:C5').</param>
    /// <param name="destRange">The destination range (e.g., 'E1:G5').</param>
    /// <param name="destCell">The destination cell (alternative to destRange).</param>
    /// <param name="copyValue">Whether to copy cell values as well.</param>
    /// <returns>OperationParameters configured for the copy format operation.</returns>
    private static OperationParameters BuildCopyFormatParameters(OperationParameters parameters, string? range,
        string? sourceRange, string? destRange, string? destCell, bool copyValue)
    {
        var effectiveRange = range ?? sourceRange;
        var effectiveDestTarget = destRange ?? destCell;
        if (effectiveRange != null) parameters.Set("range", effectiveRange);
        if (effectiveDestTarget != null) parameters.Set("destTarget", effectiveDestTarget);
        parameters.Set("copyValue", copyValue);
        return parameters;
    }
}
