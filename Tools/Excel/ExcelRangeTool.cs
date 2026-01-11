using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Excel.Range;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel ranges (write, edit, get, clear, copy, move, copy_format)
///     Merges: ExcelWriteRangeTool, ExcelEditRangeTool, ExcelGetRangeTool, ExcelClearRangeTool,
///     ExcelCopyRangeTool, ExcelMoveRangeTool, ExcelCopyFormatTool
/// </summary>
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
        _handlerRegistry = ExcelRangeHandlerRegistry.Create();
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
    [McpServerTool(Name = "excel_range")]
    [Description(@"Manage Excel ranges. Supports 7 operations: write, edit, get, clear, copy, move, copy_format.

Usage examples:
- Write range: excel_range(operation='write', path='book.xlsx', startCell='A1', data=[['A','B'],['C','D']])
- Edit range: excel_range(operation='edit', path='book.xlsx', range='A1:B2', data=[['X','Y']])
- Get range: excel_range(operation='get', path='book.xlsx', range='A1:B2')
- Clear range: excel_range(operation='clear', path='book.xlsx', range='A1:B2')
- Copy range: excel_range(operation='copy', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Move range: excel_range(operation='move', path='book.xlsx', sourceRange='A1:B2', destCell='C1')
- Copy format: excel_range(operation='copy_format', path='book.xlsx', sourceRange='A1:B2', destCell='C1')")]
    public string Execute(
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

        if (operation.ToLowerInvariant() == "get")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
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

        switch (operation.ToLowerInvariant())
        {
            case "write":
                if (startCell != null) parameters.Set("startCell", startCell);
                if (data != null) parameters.Set("data", data);
                break;

            case "edit":
                if (range != null) parameters.Set("range", range);
                if (data != null) parameters.Set("data", data);
                parameters.Set("clearRange", clearRange);
                break;

            case "get":
                if (range != null) parameters.Set("range", range);
                parameters.Set("includeFormulas", includeFormulas);
                parameters.Set("calculateFormulas", calculateFormulas);
                parameters.Set("includeFormat", includeFormat);
                break;

            case "clear":
                if (range != null) parameters.Set("range", range);
                parameters.Set("clearContent", clearContent);
                parameters.Set("clearFormat", clearFormat);
                break;

            case "copy":
                if (sourceSheetIndex != null) parameters.Set("sourceSheetIndex", sourceSheetIndex);
                if (destSheetIndex != null) parameters.Set("destSheetIndex", destSheetIndex);
                if (sourceRange != null) parameters.Set("sourceRange", sourceRange);
                if (destCell != null) parameters.Set("destCell", destCell);
                parameters.Set("copyOptions", copyOptions);
                break;

            case "move":
                if (sourceSheetIndex != null) parameters.Set("sourceSheetIndex", sourceSheetIndex);
                if (destSheetIndex != null) parameters.Set("destSheetIndex", destSheetIndex);
                if (sourceRange != null) parameters.Set("sourceRange", sourceRange);
                if (destCell != null) parameters.Set("destCell", destCell);
                break;

            case "copy_format":
                // Handle range/sourceRange and destRange/destCell aliasing
                var effectiveRange = range ?? sourceRange;
                var effectiveDestTarget = destRange ?? destCell;
                if (effectiveRange != null) parameters.Set("range", effectiveRange);
                if (effectiveDestTarget != null) parameters.Set("destTarget", effectiveDestTarget);
                parameters.Set("copyValue", copyValue);
                break;
        }

        return parameters;
    }
}
